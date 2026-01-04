use crate::document::Document;
use std::cell::RefCell;
use windows::{
    core::*,
    Win32::{
        Foundation::{GENERIC_READ, GENERIC_WRITE},
        Graphics::Imaging::*,
        System::Com::{StructuredStorage::IPropertyBag2, *},
    },
};

thread_local! {
    static WIC_FACTORY: RefCell<Option<IWICImagingFactory>> = const { RefCell::new(None) };
}

fn get_wic_factory() -> Result<IWICImagingFactory> {
    WIC_FACTORY.with(|cell| {
        let mut opt = cell.borrow_mut();
        if opt.is_none() {
            let factory: IWICImagingFactory = unsafe {
                CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?
            };
            *opt = Some(factory);
        }
        Ok(opt.as_ref().unwrap().clone())
    })
}

pub struct ClipboardBitmapData {
    pub width: u32,
    pub height: u32,
    pub data: Vec<u8>,
}

pub struct WicLoader {
    _marker: std::marker::PhantomData<()>,
}

impl WicLoader {
    pub fn new() -> Result<Self> {
        // Initialize WIC factory
        let _ = get_wic_factory()?;
        Ok(Self {
            _marker: std::marker::PhantomData,
        })
    }

    pub fn load(&self, path: &str) -> Result<Document> {
        let factory = get_wic_factory()?;

        unsafe {
            // Create decoder from file
            let path_wide: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
            let decoder = factory.CreateDecoderFromFilename(
                PCWSTR(path_wide.as_ptr()),
                None,
                GENERIC_READ,
                WICDecodeMetadataCacheOnDemand,
            )?;

            // Get frame
            let frame = decoder.GetFrame(0)?;

            // Get dimensions
            let mut width = 0u32;
            let mut height = 0u32;
            frame.GetSize(&mut width, &mut height)?;

            // Convert to BGRA format
            let converter = factory.CreateFormatConverter()?;
            converter.Initialize(
                &frame,
                &GUID_WICPixelFormat32bppPBGRA,
                WICBitmapDitherTypeNone,
                None,
                0.0,
                WICBitmapPaletteTypeMedianCut,
            )?;

            // Create WIC bitmap
            let wic_bitmap = factory.CreateBitmapFromSource(&converter, WICBitmapCacheOnLoad)?;

            Ok(Document::new_image(wic_bitmap, width, height))
        }
    }

    pub fn save(&self, doc: &Document, path: &str, page: usize) -> Result<()> {
        let factory = get_wic_factory()?;

        // Determine output format from extension
        let ext = std::path::Path::new(path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png")
            .to_lowercase();

        let container_format = match ext.as_str() {
            "jpg" | "jpeg" => &GUID_ContainerFormatJpeg,
            "png" => &GUID_ContainerFormatPng,
            "bmp" => &GUID_ContainerFormatBmp,
            "tif" | "tiff" => &GUID_ContainerFormatTiff,
            "webp" => &GUID_ContainerFormatWebp,
            _ => &GUID_ContainerFormatPng,
        };

        unsafe {
            // Create stream for output
            let path_wide: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
            let stream = factory.CreateStream()?;
            stream.InitializeFromFilename(PCWSTR(path_wide.as_ptr()), GENERIC_WRITE.0)?;

            // Create encoder
            let encoder = factory.CreateEncoder(container_format, std::ptr::null())?;
            encoder.Initialize(&stream, WICBitmapEncoderNoCache)?;

            // Create frame
            let mut frame: Option<IWICBitmapFrameEncode> = None;
            let mut props: Option<IPropertyBag2> = None;
            encoder.CreateNewFrame(&mut frame, &mut props)?;

            let frame = frame.ok_or_else(Error::from_win32)?;
            frame.Initialize(props.as_ref())?;

            // Get source bitmap
            if let Some(wic_bitmap) = doc.get_wic_bitmap(page) {
                let mut width = 0u32;
                let mut height = 0u32;
                wic_bitmap.GetSize(&mut width, &mut height)?;

                frame.SetSize(width, height)?;

                // Set pixel format
                let mut pixel_format = GUID_WICPixelFormat32bppBGRA;
                frame.SetPixelFormat(&mut pixel_format)?;

                // Write pixels
                frame.WriteSource(wic_bitmap, std::ptr::null())?;
            } else if let Some((data, width, height, stride)) = doc.get_pixel_data(page) {
                frame.SetSize(width, height)?;

                let mut pixel_format = GUID_WICPixelFormat32bppBGRA;
                frame.SetPixelFormat(&mut pixel_format)?;

                frame.WritePixels(height, stride, data)?;
            } else {
                return Err(Error::from_win32());
            }

            frame.Commit()?;
            encoder.Commit()?;
        }

        Ok(())
    }

    #[allow(dead_code)]
    pub fn create_bitmap_from_data(
        &self,
        width: u32,
        height: u32,
        data: &[u8],
    ) -> Result<IWICBitmap> {
        let factory = get_wic_factory()?;

        unsafe {
            let stride = width * 4;
            factory.CreateBitmapFromMemory(
                width,
                height,
                &GUID_WICPixelFormat32bppPBGRA,
                stride,
                data,
            )
        }
    }

    pub fn get_bitmap_for_clipboard(
        &self,
        doc: &Document,
        page: usize,
        rotation: i32,
    ) -> Result<ClipboardBitmapData> {
        let factory = get_wic_factory()?;

        unsafe {
            // Get source bitmap
            let source = if let Some(wic_bitmap) = doc.get_wic_bitmap(page) {
                wic_bitmap.clone()
            } else if let Some((data, width, height, stride)) = doc.get_pixel_data(page) {
                factory.CreateBitmapFromMemory(
                    width,
                    height,
                    &GUID_WICPixelFormat32bppBGRA,
                    stride,
                    data,
                )?
            } else {
                return Err(Error::from_win32());
            };

            // Apply rotation if needed
            let rotated: IWICBitmapSource = if rotation != 0 {
                let transform = match rotation {
                    90 => WICBitmapTransformRotate90,
                    180 => WICBitmapTransformRotate180,
                    270 => WICBitmapTransformRotate270,
                    _ => WICBitmapTransformRotate0,
                };
                let flip_rotator = factory.CreateBitmapFlipRotator()?;
                flip_rotator.Initialize(&source, transform)?;
                flip_rotator.cast()?
            } else {
                source.cast()?
            };

            // Convert to non-premultiplied BGRA for clipboard
            let converter = factory.CreateFormatConverter()?;
            converter.Initialize(
                &rotated,
                &GUID_WICPixelFormat32bppBGRA,
                WICBitmapDitherTypeNone,
                None,
                0.0,
                WICBitmapPaletteTypeMedianCut,
            )?;

            // Get dimensions
            let mut width = 0u32;
            let mut height = 0u32;
            converter.GetSize(&mut width, &mut height)?;

            // Read pixels
            let stride = width * 4;
            let buffer_size = (stride * height) as usize;
            let mut data = vec![0u8; buffer_size];
            converter.CopyPixels(std::ptr::null(), stride, &mut data)?;

            // Flip vertically for DIB format (bottom-up)
            let row_size = stride as usize;
            let mut flipped = vec![0u8; buffer_size];
            for y in 0..height as usize {
                let src_row = y * row_size;
                let dst_row = (height as usize - 1 - y) * row_size;
                flipped[dst_row..dst_row + row_size].copy_from_slice(&data[src_row..src_row + row_size]);
            }

            Ok(ClipboardBitmapData {
                width,
                height,
                data: flipped,
            })
        }
    }
}
