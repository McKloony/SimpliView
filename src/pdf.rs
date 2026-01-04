use crate::document::{Document, PageData};
use std::sync::atomic::{AtomicBool, Ordering};
use windows::{
    core::*,
    Data::Pdf::*,
    Foundation::Size,
    Storage::*,
    Storage::Streams::*,
};

// Maximum render dimension to prevent memory issues with very large PDFs
// 2048 pixels is sufficient for most displays while keeping memory usage reasonable
const MAX_RENDER_DIMENSION: f64 = 2048.0;

/// PDF loader that handles password-protected documents via Windows.Data.Pdf WinRT API.
///
/// Password handling approach:
/// - First load attempt is made without password (or with provided password)
/// - If load fails with a password-related error, sets `needs_password` flag
/// - Caller (App) then prompts user for password and retries with provided password
/// - WinRT handles the actual decryption internally
pub struct PdfLoader {
    needs_password: AtomicBool,
}

impl PdfLoader {
    pub fn new() -> Self {
        Self {
            needs_password: AtomicBool::new(false),
        }
    }

    /// Loads a PDF document, optionally with a password for encrypted files.
    ///
    /// # Arguments
    /// * `path` - Absolute path to the PDF file
    /// * `password` - Optional password for encrypted PDFs (supports Unicode)
    ///
    /// # Returns
    /// * `Ok(Document)` - Successfully loaded document with rendered pages
    /// * `Err` - Load failed; check `needs_password()` to determine if password is required
    pub fn load(&self, path: &str, password: Option<&str>) -> Result<Document> {
        self.needs_password.store(false, Ordering::SeqCst);

        // Open file using WinRT StorageFile
        let path_hstring: HSTRING = path.into();
        let file = StorageFile::GetFileFromPathAsync(&path_hstring)?.get()?;

        // Load PDF document - with or without password
        let pdf_doc = if let Some(pwd) = password {
            // Password provided - attempt to load with it
            let pwd_hstring: HSTRING = pwd.into();
            match PdfDocument::LoadFromFileWithPasswordAsync(&file, &pwd_hstring)?.get() {
                Ok(doc) => doc,
                Err(e) => {
                    // Wrong password - flag for retry
                    if is_password_error(&e) {
                        self.needs_password.store(true, Ordering::SeqCst);
                    }
                    return Err(e);
                }
            }
        } else {
            // No password - try loading without one first
            match PdfDocument::LoadFromFileAsync(&file)?.get() {
                Ok(doc) => doc,
                Err(e) => {
                    // Check if this is a password-protected file
                    if is_password_error(&e) {
                        self.needs_password.store(true, Ordering::SeqCst);
                    }
                    return Err(e);
                }
            }
        };

        // Get page count
        let page_count = pdf_doc.PageCount()? as usize;

        // Render each page to bitmap
        let mut pages = Vec::with_capacity(page_count);

        for i in 0..page_count {
            let page = pdf_doc.GetPage(i as u32)?;

            // Get original page size
            let page_size: Size = page.Size()?;
            let orig_width = page_size.Width as f64;
            let orig_height = page_size.Height as f64;

            // Create in-memory stream for rendering
            let stream = InMemoryRandomAccessStream::new()?;

            // Check if page needs to be scaled down
            let max_dim = orig_width.max(orig_height);
            if max_dim > MAX_RENDER_DIMENSION {
                // Calculate scaled dimensions maintaining aspect ratio
                let scale = MAX_RENDER_DIMENSION / max_dim;
                let render_width = (orig_width * scale) as u32;
                let render_height = (orig_height * scale) as u32;

                // Create render options with size limit
                let options = PdfPageRenderOptions::new()?;
                options.SetDestinationWidth(render_width)?;
                options.SetDestinationHeight(render_height)?;

                // Render with options
                page.RenderWithOptionsToStreamAsync(&stream, &options)?.get()?;
            } else {
                // Render at original size
                page.RenderToStreamAsync(&stream)?.get()?;
            }

            // Read pixel data from the stream
            let (pixel_data, actual_width, actual_height) = self.read_stream_to_pixels(&stream)?;

            // Store actual rendered dimensions (pixel data dimensions must match)
            pages.push(PageData {
                width: actual_width as f32,
                height: actual_height as f32,
                wic_bitmap: None,
                pixel_data: Some(pixel_data),
                stride: actual_width * 4,
            });

            // Close the page
            page.Close()?;
        }

        Ok(Document::new_pdf(pages))
    }

    fn read_stream_to_pixels(
        &self,
        stream: &InMemoryRandomAccessStream,
    ) -> Result<(Vec<u8>, u32, u32)> {
        unsafe {
            use windows::Win32::Graphics::Imaging::*;
            use windows::Win32::System::Com::*;

            let factory: IWICImagingFactory =
                CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?;

            // Get stream as IStream
            stream.Seek(0)?;
            let size = stream.Size()? as usize;

            // Read data from WinRT stream
            let reader = DataReader::CreateDataReader(&stream.GetInputStreamAt(0)?)?;
            reader.LoadAsync(size as u32)?.get()?;

            let mut buffer = vec![0u8; size];
            reader.ReadBytes(&mut buffer)?;

            // Create WIC stream from memory
            let wic_stream = factory.CreateStream()?;
            wic_stream.InitializeFromMemory(&buffer)?;

            // Create decoder
            let decoder = factory.CreateDecoderFromStream(
                &wic_stream,
                std::ptr::null(),
                WICDecodeMetadataCacheOnDemand,
            )?;

            // Get frame
            let frame = decoder.GetFrame(0)?;

            // Convert to BGRA
            let converter = factory.CreateFormatConverter()?;
            converter.Initialize(
                &frame,
                &GUID_WICPixelFormat32bppPBGRA,
                WICBitmapDitherTypeNone,
                None,
                0.0,
                WICBitmapPaletteTypeMedianCut,
            )?;

            // Get actual dimensions
            let mut actual_width = 0u32;
            let mut actual_height = 0u32;
            converter.GetSize(&mut actual_width, &mut actual_height)?;

            // Read pixels
            let stride = actual_width * 4;
            let buffer_size = (stride * actual_height) as usize;
            let mut pixel_data = vec![0u8; buffer_size];

            converter.CopyPixels(
                std::ptr::null(),
                stride,
                &mut pixel_data,
            )?;

            Ok((pixel_data, actual_width, actual_height))
        }
    }

    pub fn needs_password(&self) -> bool {
        self.needs_password.load(Ordering::SeqCst)
    }
}

/// Checks if the error indicates a password-protected or incorrectly-passworded PDF.
///
/// Windows.Data.Pdf returns these error codes for password issues:
/// - 0x80070005 (E_ACCESSDENIED): File is encrypted and requires a password
/// - 0x8007052B (ERROR_WRONG_PASSWORD): Provided password is incorrect
///
/// Note: We deliberately do NOT include 0x8007000D (E_INVALID_DATA) or 0x80004005 (E_FAIL)
/// as those typically indicate corrupted/malformed PDFs, not password issues.
fn is_password_error(e: &Error) -> bool {
    let code = e.code().0 as u32;
    code == 0x80070005      // E_ACCESSDENIED - password required
        || code == 0x8007052B   // ERROR_WRONG_PASSWORD - wrong password provided
}
