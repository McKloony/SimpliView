use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Dwm::*,
        System::Registry::*,
        UI::WindowsAndMessaging::*,
    },
};

pub struct Theme;

impl Theme {
    /// Check if Windows is using dark mode
    #[allow(dead_code)]
    pub fn is_system_dark_mode() -> bool {
        unsafe {
            // Read from registry: HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize
            let mut hkey = HKEY::default();
            let result = RegOpenKeyExW(
                HKEY_CURRENT_USER,
                w!("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"),
                0,
                KEY_READ,
                &mut hkey,
            );

            if result.is_err() {
                return false;
            }

            let mut value: u32 = 1;
            let mut size = std::mem::size_of::<u32>() as u32;

            let result = RegQueryValueExW(
                hkey,
                w!("AppsUseLightTheme"),
                None,
                None,
                Some(&mut value as *mut u32 as *mut u8),
                Some(&mut size),
            );

            let _ = RegCloseKey(hkey);

            if result.is_ok() {
                value == 0 // 0 = dark mode, 1 = light mode
            } else {
                false
            }
        }
    }

    /// Apply dark/light theme to window
    pub fn apply_to_window(hwnd: HWND, is_dark: bool) {
        unsafe {
            // Use DwmSetWindowAttribute to enable dark mode title bar on Windows 10/11
            let use_dark: BOOL = if is_dark { TRUE } else { FALSE };

            // DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 10 20H1+)
            let _ = DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                &use_dark as *const _ as *const _,
                std::mem::size_of::<BOOL>() as u32,
            );

            // Force window to redraw with new theme
            let _ = SetWindowPos(
                hwnd,
                None,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED,
            );
        }
    }

    /// Get appropriate colors for the current theme
    #[allow(dead_code)]
    pub fn get_colors(is_dark: bool) -> ThemeColors {
        if is_dark {
            ThemeColors {
                background: 0x1E1E1E,
                text: 0xFFFFFF,
                accent: 0x0078D4,
                border: 0x3F3F3F,
            }
        } else {
            ThemeColors {
                background: 0xF3F3F3,
                text: 0x000000,
                accent: 0x0078D4,
                border: 0xD0D0D0,
            }
        }
    }
}

#[derive(Clone, Copy)]
#[allow(dead_code)]
pub struct ThemeColors {
    pub background: u32,
    pub text: u32,
    pub accent: u32,
    pub border: u32,
}

#[allow(dead_code)]
impl ThemeColors {
    pub fn background_rgb(&self) -> (u8, u8, u8) {
        (
            ((self.background >> 16) & 0xFF) as u8,
            ((self.background >> 8) & 0xFF) as u8,
            (self.background & 0xFF) as u8,
        )
    }

    pub fn text_rgb(&self) -> (u8, u8, u8) {
        (
            ((self.text >> 16) & 0xFF) as u8,
            ((self.text >> 8) & 0xFF) as u8,
            (self.text & 0xFF) as u8,
        )
    }
}
