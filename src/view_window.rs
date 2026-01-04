use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::*,
        UI::WindowsAndMessaging::*,
    },
};

const VIEW_CLASS_NAME: PCWSTR = w!("SimpliViewCanvas");

pub struct ViewWindow {
    hwnd: HWND,
}

impl ViewWindow {
    pub fn new(parent: HWND, instance: HMODULE) -> Result<Self> {
        unsafe {
            let wc = WNDCLASSEXW {
                cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
                style: CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS,
                lpfnWndProc: Some(view_wnd_proc),
                cbClsExtra: 0,
                cbWndExtra: 0,
                hInstance: instance,
                hIcon: HICON::default(),
                hCursor: LoadCursorW(None, IDC_ARROW)?,
                hbrBackground: HBRUSH::default(), // D2D handles background
                lpszMenuName: PCWSTR::null(),
                lpszClassName: VIEW_CLASS_NAME,
                hIconSm: HICON::default(),
            };

            let _ = RegisterClassExW(&wc);

            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                VIEW_CLASS_NAME,
                None,
                WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_HSCROLL | WS_VSCROLL,
                0,
                0,
                0,
                0,
                parent,
                HMENU(3000isize),
                instance,
                None,
            );

            if hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

            Ok(Self { hwnd })
        }
    }

    pub fn hwnd(&self) -> HWND {
        self.hwnd
    }

    pub fn resize(&self, x: i32, y: i32, width: i32, height: i32) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                None,
                x,
                y,
                width,
                height,
                SWP_NOZORDER,
            );
        }
    }

    pub fn invalidate(&self) {
        unsafe {
            let _ = InvalidateRect(self.hwnd, None, false);
        }
    }
}

extern "system" fn view_wnd_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        // Forward relevant messages to parent
        match msg {
            WM_PAINT => {
                // Let parent handle paint via message forwarding or just validate
                // We notify the App to paint this window.
                let parent = GetParent(hwnd);
                SendMessageW(parent, WM_APP_VIEW_PAINT, WPARAM(0), LPARAM(hwnd.0));
                ValidateRect(hwnd, None); // Mark as painted
                LRESULT(0)
            }
            WM_ERASEBKGND => LRESULT(1), // Prevent flicker
            
            // Forward input to parent for handling
            WM_LBUTTONDOWN | WM_LBUTTONUP | WM_RBUTTONDOWN | WM_MOUSEMOVE | WM_MOUSEWHEEL
            | WM_KEYDOWN | WM_KEYUP | WM_HSCROLL | WM_VSCROLL => {
                let parent = GetParent(hwnd);
                SendMessageW(parent, msg, wparam, lparam)
            }
            
            WM_SETCURSOR => {
                 let parent = GetParent(hwnd);
                 if SendMessageW(parent, msg, wparam, lparam).0 == 1 {
                     LRESULT(1)
                 } else {
                     DefWindowProcW(hwnd, msg, wparam, lparam)
                 }
            }
            
            WM_CONTEXTMENU => {
                 let parent = GetParent(hwnd);
                 SendMessageW(parent, msg, wparam, lparam)
            }

            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

pub const WM_APP_VIEW_PAINT: u32 = WM_USER + 100;