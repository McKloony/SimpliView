//! Gemeinsame Utility-Funktionen

use windows::{
    core::*,
    Win32::{
        Graphics::Gdi::*,
        Graphics::Imaging::*,
        System::Com::*,
    },
};

/// Creates a DWORD from two WORDs (low and high)
#[inline]
pub const fn make_long(lo: u16, hi: u16) -> u32 {
    (lo as u32) | ((hi as u32) << 16)
}

/// Loads a PNG image from memory and returns an HBITMAP
pub fn load_png_from_memory(data: &[u8]) -> Result<HBITMAP> {
    unsafe {
        let factory: IWICImagingFactory =
            CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?;

        let stream = factory.CreateStream()?;
        stream.InitializeFromMemory(data)?;

        let decoder = factory.CreateDecoderFromStream(
            &stream,
            std::ptr::null(),
            WICDecodeMetadataCacheOnDemand,
        )?;

        let frame = decoder.GetFrame(0)?;
        let mut width = 0u32;
        let mut height = 0u32;
        frame.GetSize(&mut width, &mut height)?;

        let converter = factory.CreateFormatConverter()?;
        converter.Initialize(
            &frame,
            &GUID_WICPixelFormat32bppBGRA,
            WICBitmapDitherTypeNone,
            None,
            0.0,
            WICBitmapPaletteTypeMedianCut,
        )?;

        let hdc = GetDC(None);
        let bmi = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width as i32,
                biHeight: -(height as i32),
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0 as u32,
                ..Default::default()
            },
            bmiColors: [RGBQUAD::default()],
        };

        let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
        let bitmap = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &mut bits, None, 0)?;
        let _ = ReleaseDC(None, hdc);

        let stride = width * 4;
        let buffer_size = stride * height;
        converter.CopyPixels(
            std::ptr::null(),
            stride,
            std::slice::from_raw_parts_mut(bits as *mut u8, buffer_size as usize),
        )?;

        Ok(bitmap)
    }
}
