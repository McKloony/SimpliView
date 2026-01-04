use crate::app::AppState;
use parking_lot::Mutex;
use std::sync::Arc;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::*,
        System::LibraryLoader::*,
        UI::{
            Controls::*,
            HiDpi::*,
            WindowsAndMessaging::*,
        },
    },
};

const CLASS_NAME: PCWSTR = w!("SimpliViewWindow");

pub struct Window {
    hwnd: HWND,
    instance: HMODULE,
}

impl Window {
    pub fn new(title: &str, _state: Arc<Mutex<AppState>>) -> Result<Self> {
        unsafe {
            // Enable Per-Monitor DPI awareness
            let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

            // Initialize common controls
            let icc = INITCOMMONCONTROLSEX {
                dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
                dwICC: ICC_BAR_CLASSES | ICC_STANDARD_CLASSES,
            };
            InitCommonControlsEx(&icc);

            let instance = GetModuleHandleW(None)?;

            // Register window class
            // Load large icon (typically 32x32)
            let icon_big = LoadImageW(
                instance,
                PCWSTR(1 as *const u16), // ID 1 from .rc file
                IMAGE_ICON,
                GetSystemMetrics(SM_CXICON),
                GetSystemMetrics(SM_CYICON),
                LR_DEFAULTCOLOR | LR_SHARED,
            )
            .map(|h| HICON(h.0))
            .unwrap_or_else(|_| LoadIconW(None, IDI_APPLICATION).unwrap_or(HICON::default()));

            // Load small icon (typically 16x16)
            let icon_small = LoadImageW(
                instance,
                PCWSTR(1 as *const u16),
                IMAGE_ICON,
                GetSystemMetrics(SM_CXSMICON),
                GetSystemMetrics(SM_CYSMICON),
                LR_DEFAULTCOLOR | LR_SHARED,
            )
            .map(|h| HICON(h.0))
            .unwrap_or(icon_big);

            let wc = WNDCLASSEXW {
                cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
                style: CS_HREDRAW | CS_VREDRAW,
                lpfnWndProc: Some(wnd_proc),
                cbClsExtra: 0,
                cbWndExtra: 0,
                hInstance: instance,
                hIcon: icon_big,
                hCursor: LoadCursorW(None, IDC_ARROW)?,
                hbrBackground: HBRUSH::default(),
                lpszMenuName: PCWSTR::null(),
                lpszClassName: CLASS_NAME,
                hIconSm: icon_small,
            };

            let atom = RegisterClassExW(&wc);
            if atom == 0 {
                return Err(Error::from_win32());
            }

            // Get monitor work area for initial window size
            let (x, y, width, height) = get_initial_window_rect();

            // Create the window
            let title_wide: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                CLASS_NAME,
                PCWSTR(title_wide.as_ptr()),
                WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                x,
                y,
                width,
                height,
                None,
                None,
                instance,
                None,
            );

            if hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

            Ok(Self { hwnd, instance })
        }
    }

    pub fn hwnd(&self) -> HWND {
        self.hwnd
    }

    pub fn instance(&self) -> HMODULE {
        self.instance
    }

    pub fn show(&self) {
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_SHOW);
            let _ = UpdateWindow(self.hwnd);
            // Bring to foreground
            let _ = SetForegroundWindow(self.hwnd);
        }
    }

    pub fn set_title(&self, title: &str) {
        let title_wide: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
        unsafe {
            let _ = SetWindowTextW(self.hwnd, PCWSTR(title_wide.as_ptr()));
        }
    }

    /// Set the App pointer for message handling
    #[allow(dead_code)]
    pub fn set_app_ptr(&self, app_ptr: *mut std::ffi::c_void) {
        unsafe {
            SetWindowLongPtrW(self.hwnd, GWLP_USERDATA, app_ptr as isize);
        }
    }
}

fn get_initial_window_rect() -> (i32, i32, i32, i32) {
    unsafe {
        // Get cursor position to determine which monitor to use
        let mut cursor_pos = POINT::default();
        let _ = GetCursorPos(&mut cursor_pos);

        // Get the monitor containing the cursor
        let monitor = MonitorFromPoint(cursor_pos, MONITOR_DEFAULTTOPRIMARY);

        // Get monitor info
        let mut mi = MONITORINFO {
            cbSize: std::mem::size_of::<MONITORINFO>() as u32,
            ..Default::default()
        };

        if GetMonitorInfoW(monitor, &mut mi).as_bool() {
            // Use work area (excludes taskbar)
            let work = mi.rcWork;
            let width = work.right - work.left;
            let height = work.bottom - work.top;

            // Slightly inset from full work area for a "near-maximized" look
            let margin = 8;
            (
                work.left + margin,
                work.top + margin,
                width - margin * 2,
                height - margin * 2,
            )
        } else {
            // Fallback to default size
            (CW_USEDEFAULT, CW_USEDEFAULT, 1024, 768)
        }
    }
}

extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        // Try to get App handler and delegate message
        let app_ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut std::ffi::c_void;
        if !app_ptr.is_null() {
            // Call the App's message handler via function pointer stored after App pointer
            let handler = app_ptr as *mut crate::app::App;
            if let Some(result) = (*handler).handle_window_message(msg, wparam, lparam) {
                return result;
            }
        }

        match msg {
            WM_CREATE => {
                LRESULT(0)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                LRESULT(0)
            }
            WM_ERASEBKGND => {
                // Prevent flicker by not erasing background
                // D2D will handle the background
                LRESULT(1)
            }
            WM_SIZE => {
                // Rebar auto-sizes when receiving WM_SIZE
                let rebar = FindWindowExW(hwnd, None, w!("ReBarWindow32"), None);
                if rebar.0 != 0 {
                    SendMessageW(rebar, WM_SIZE, WPARAM(0), LPARAM(0));
                }
                // Statusbar auto-sizes when receiving WM_SIZE with 0
                let statusbar = FindWindowExW(hwnd, None, w!("msctls_statusbar32"), None);
                if statusbar.0 != 0 {
                    SendMessageW(statusbar, WM_SIZE, WPARAM(0), LPARAM(0));
                }
                // Invalidate to repaint content area
                let _ = InvalidateRect(hwnd, None, false);
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_PAINT => {
                // Default paint handler - App should handle this via message callback
                let mut ps = PAINTSTRUCT::default();
                let hdc = BeginPaint(hwnd, &mut ps);
                // Fill with a neutral background color to prevent artifacts
                let brush = CreateSolidBrush(COLORREF(0x1E1E1E)); // Dark gray
                let _ = FillRect(hdc, &ps.rcPaint, brush);
                let _ = DeleteObject(brush);
                let _ = EndPaint(hwnd, &ps);
                LRESULT(0)
            }
            WM_GETMINMAXINFO => {
                let mmi = &mut *(lparam.0 as *mut MINMAXINFO);
                mmi.ptMinTrackSize.x = 400;
                mmi.ptMinTrackSize.y = 300;
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}
