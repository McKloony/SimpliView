use crate::icons;
use parking_lot::Mutex;
use std::sync::Arc;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::*,
        Graphics::Imaging::*,
        System::Com::*,
        UI::WindowsAndMessaging::*,
    },
};

// Context menu command IDs
pub const IDM_FIT_TO_PAGE: u32 = 200;
pub const IDM_ROTATE_LEFT: u32 = 201;
pub const IDM_ROTATE_RIGHT: u32 = 202;

pub struct ContextMenu {
    menu: HMENU,
    pending_command: Arc<Mutex<Option<u32>>>,
    bitmaps: Vec<HBITMAP>, // Keep bitmaps alive
}

impl ContextMenu {
    pub fn new() -> Result<Self> {
        unsafe {
            let menu = CreatePopupMenu()?;
            let mut bitmaps = Vec::new();

            // Load icons as bitmaps
            let bmp_fit = Self::load_png_as_bitmap(icons::ICON_FIT_TO_SIZE)?;
            let bmp_rotate_left = Self::load_png_as_bitmap(icons::ICON_ROTATE_LEFT)?;
            let bmp_rotate_right = Self::load_png_as_bitmap(icons::ICON_ROTATE_RIGHT)?;

            // Add menu items with icons
            Self::append_menu_item_with_icon(menu, IDM_FIT_TO_PAGE, w!("Fit to Page"), bmp_fit);
            let _ = AppendMenuW(menu, MF_SEPARATOR, 0, None);
            Self::append_menu_item_with_icon(menu, IDM_ROTATE_LEFT, w!("Rotate Left"), bmp_rotate_left);
            Self::append_menu_item_with_icon(menu, IDM_ROTATE_RIGHT, w!("Rotate Right"), bmp_rotate_right);

            // Store bitmaps to keep them alive
            bitmaps.push(bmp_fit);
            bitmaps.push(bmp_rotate_left);
            bitmaps.push(bmp_rotate_right);

            Ok(Self {
                menu,
                pending_command: Arc::new(Mutex::new(None)),
                bitmaps,
            })
        }
    }

    fn append_menu_item_with_icon(menu: HMENU, id: u32, text: PCWSTR, bitmap: HBITMAP) {
        unsafe {
            let _ = AppendMenuW(menu, MF_STRING, id as usize, text);

            let mii = MENUITEMINFOW {
                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_BITMAP,
                hbmpItem: bitmap,
                ..Default::default()
            };
            let _ = SetMenuItemInfoW(menu, id, false, &mii);
        }
    }

    fn load_png_as_bitmap(data: &[u8]) -> Result<HBITMAP> {
        unsafe {
            let factory: IWICImagingFactory =
                CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?;
            let stream = factory.CreateStream()?;
            stream.InitializeFromMemory(data)?;
            let decoder = factory.CreateDecoderFromStream(&stream, std::ptr::null(), WICDecodeMetadataCacheOnDemand)?;
            let frame = decoder.GetFrame(0)?;

            let mut width = 0u32;
            let mut height = 0u32;
            frame.GetSize(&mut width, &mut height)?;

            // Convert to premultiplied BGRA for proper alpha blending in menus
            let converter = factory.CreateFormatConverter()?;
            converter.Initialize(
                &frame,
                &GUID_WICPixelFormat32bppPBGRA,
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

    pub fn show(&self, hwnd: HWND, x: i32, y: i32) {
        unsafe {
            // If coordinates are -1, -1, use cursor position
            let (screen_x, screen_y) = if x == -1 && y == -1 {
                let mut pt = POINT::default();
                let _ = GetCursorPos(&mut pt);
                (pt.x, pt.y)
            } else {
                (x, y)
            };

            let cmd = TrackPopupMenu(
                self.menu,
                TPM_RETURNCMD | TPM_RIGHTBUTTON,
                screen_x,
                screen_y,
                0,
                hwnd,
                None,
            );

            if cmd.0 != 0 {
                // Store the command for processing
                *self.pending_command.lock() = Some(cmd.0 as u32);

                // Also send WM_COMMAND to the window
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(cmd.0 as usize), LPARAM(0));
            }
        }
    }

    pub fn poll_command(&self) -> Option<u32> {
        self.pending_command.lock().take()
    }

    pub fn set_document_loaded(&self, loaded: bool) {
        unsafe {
            let flag = if loaded { MF_ENABLED } else { MF_GRAYED };
            let _ = EnableMenuItem(self.menu, IDM_FIT_TO_PAGE, flag);
            let _ = EnableMenuItem(self.menu, IDM_ROTATE_LEFT, flag);
            let _ = EnableMenuItem(self.menu, IDM_ROTATE_RIGHT, flag);
        }
    }
}

impl Drop for ContextMenu {
    fn drop(&mut self) {
        unsafe {
            let _ = DestroyMenu(self.menu);
            for bitmap in &self.bitmaps {
                let _ = DeleteObject(*bitmap);
            }
        }
    }
}
