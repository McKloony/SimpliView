use crate::icons;
use crate::utils::{load_png_from_memory, make_long};
use parking_lot::Mutex;
use std::sync::Arc;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::{DeleteObject, HBITMAP, InvalidateRect},
        UI::{Controls::*, WindowsAndMessaging::*},
    },
};

// Command IDs
pub const ID_OPEN: u16 = 100;
pub const ID_EXPORT: u16 = 101;
pub const ID_ROTATE_LEFT: u16 = 102;
pub const ID_ROTATE_RIGHT: u16 = 103;
pub const ID_PREV_PAGE: u16 = 104;
pub const ID_NEXT_PAGE: u16 = 105;
pub const ID_INFO: u16 = 106;
pub const ID_CLOSE: u16 = 107;
pub const ID_PRINT: u16 = 108;
pub const ID_SPRING: u16 = 9999;

#[derive(Clone, Copy, Debug)]
#[allow(dead_code)]
pub enum ToolbarCommand {
    Open,
    Export,
    RotateLeft,
    RotateRight,
    PrevPage,
    NextPage,
    Print,
    Info,
    Close,
}

#[allow(dead_code)]
pub enum ToolbarType {
    Top,
    Bottom,
}

pub struct Toolbar {
    rebar_hwnd: HWND,
    toolbar_hwnd: HWND,
    image_list: HIMAGELIST,
    pending_command: Arc<Mutex<Option<ToolbarCommand>>>,
    is_dark: bool,
    toolbar_type: ToolbarType,
}

impl Toolbar {
    pub fn new(parent: HWND, instance: HMODULE, toolbar_type: ToolbarType) -> Result<Self> {
        unsafe {
            let (rebar_id, toolbar_id, ccs_style) = match toolbar_type {
                ToolbarType::Top => (
                    HMENU(999isize), 
                    HMENU(1000isize), 
                    CCS_TOP as u32
                ),
                ToolbarType::Bottom => (
                    HMENU(1999isize), 
                    HMENU(2000isize), 
                    CCS_BOTTOM as u32 | CCS_NOPARENTALIGN as u32
                ),
            };

            // Create rebar control
            let rebar_hwnd = CreateWindowExW(
                WS_EX_TOOLWINDOW,
                w!("ReBarWindow32"),
                None,
                WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
                WINDOW_STYLE(
                    RBS_VARHEIGHT |
                    CCS_NODIVIDER as u32 |
                    ccs_style
                ),
                0,
                0,
                0,
                0,
                parent,
                rebar_id,
                instance,
                None,
            );

            if rebar_hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

            // Create toolbar
            let toolbar_hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("ToolbarWindow32"),
                None,
                WS_CHILD | WS_VISIBLE | WINDOW_STYLE(
                    TBSTYLE_FLAT |
                    TBSTYLE_LIST |
                    TBSTYLE_TOOLTIPS |
                    TBSTYLE_TRANSPARENT |
                    CCS_NODIVIDER as u32 |
                    CCS_NORESIZE as u32 |
                    CCS_NOPARENTALIGN as u32
                ),
                0,
                0,
                0,
                0,
                rebar_hwnd,
                toolbar_id,
                instance,
                None,
            );

            if toolbar_hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

            // Set toolbar extended style
            SendMessageW(
                toolbar_hwnd,
                TB_SETEXTENDEDSTYLE,
                WPARAM(0),
                LPARAM(TBSTYLE_EX_MIXEDBUTTONS as isize),
            );

            // Set smaller button size
            SendMessageW(
                toolbar_hwnd,
                TB_SETBUTTONSIZE,
                WPARAM(0),
                LPARAM(make_long(0, 24) as isize),
            );

            // Create and set image list
            let image_list = Self::create_image_list()?;
            SendMessageW(
                toolbar_hwnd,
                TB_SETIMAGELIST,
                WPARAM(0),
                LPARAM(image_list.0),
            );

            // Add buttons
            let buttons = Self::create_buttons(&toolbar_type);
            SendMessageW(
                toolbar_hwnd,
                TB_ADDBUTTONS,
                WPARAM(buttons.len()),
                LPARAM(buttons.as_ptr() as isize),
            );

            // Auto-size toolbar
            SendMessageW(toolbar_hwnd, TB_AUTOSIZE, WPARAM(0), LPARAM(0));

            // Get toolbar size for rebar band
            let mut tb_size = SIZE::default();
            SendMessageW(toolbar_hwnd, TB_GETMAXSIZE, WPARAM(0), LPARAM(&mut tb_size as *mut _ as isize));

            // Add toolbar to rebar
            let rbbi = REBARBANDINFOW {
                cbSize: std::mem::size_of::<REBARBANDINFOW>() as u32,
                fMask: RBBIM_STYLE | RBBIM_CHILD | RBBIM_CHILDSIZE | RBBIM_SIZE,
                fStyle: RBBS_CHILDEDGE | RBBS_GRIPPERALWAYS,
                hwndChild: toolbar_hwnd,
                cxMinChild: tb_size.cx as u32,
                cyMinChild: tb_size.cy as u32,
                cx: tb_size.cx as u32,
                ..Default::default()
            };

            SendMessageW(
                rebar_hwnd,
                RB_INSERTBANDW,
                WPARAM(usize::MAX),
                LPARAM(&rbbi as *const _ as isize),
            );

            // Apply theme
            let _ = SetWindowTheme(rebar_hwnd, w!("Explorer"), None);
            let _ = SetWindowTheme(toolbar_hwnd, w!("Explorer"), None);

            Ok(Self {
                rebar_hwnd,
                toolbar_hwnd,
                image_list,
                pending_command: Arc::new(Mutex::new(None)),
                is_dark: false,
                toolbar_type,
            })
        }
    }

    fn create_image_list() -> Result<HIMAGELIST> {
        unsafe {
            let image_list = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 8, 0);
            if image_list.0 == 0 {
                return Err(Error::from_win32());
            }

            let icon_data: &[&[u8]] = &[
                icons::ICON_FOLDER_OPEN,
                icons::ICON_FOLDER_OUT,
                icons::ICON_ROTATE_LEFT,
                icons::ICON_ROTATE_RIGHT,
                icons::ICON_NAV_LEFT,
                icons::ICON_NAV_RIGHT,
                icons::ICON_INFORMATION,
                icons::ICON_CLOSE,
                icons::ICON_PRINT,
            ];

            for data in icon_data {
                let bitmap = load_png_from_memory(data)?;
                ImageList_Add(image_list, bitmap, HBITMAP::default());
                let _ = DeleteObject(bitmap);
            }

            Ok(image_list)
        }
    }

    fn create_buttons(toolbar_type: &ToolbarType) -> Vec<TBBUTTON> {
        let mut buttons = Vec::new();

        let add_button = |buttons: &mut Vec<TBBUTTON>, id: i32, image: i32, text: &str| {
            let text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
            let text_ptr = Box::leak(text_wide.into_boxed_slice()).as_ptr() as isize;

            buttons.push(TBBUTTON {
                iBitmap: image,
                idCommand: id,
                fsState: TBSTATE_ENABLED as u8,
                fsStyle: (BTNS_BUTTON | BTNS_SHOWTEXT) as u8,
                bReserved: [0; 6],
                dwData: 0,
                iString: text_ptr,
            });
        };

        match toolbar_type {
            ToolbarType::Top => {
                add_button(&mut buttons, ID_OPEN as i32, 0, "Öffnen");
                add_button(&mut buttons, ID_EXPORT as i32, 1, "Exportieren");

                buttons.push(TBBUTTON {
                    iBitmap: 0,
                    idCommand: 0,
                    fsState: 0,
                    fsStyle: BTNS_SEP as u8,
                    bReserved: [0; 6],
                    dwData: 0,
                    iString: 0,
                });

                add_button(&mut buttons, ID_ROTATE_LEFT as i32, 2, "Links");
                add_button(&mut buttons, ID_ROTATE_RIGHT as i32, 3, "Rechts");

                buttons.push(TBBUTTON {
                    iBitmap: 0,
                    idCommand: 0,
                    fsState: 0,
                    fsStyle: BTNS_SEP as u8,
                    bReserved: [0; 6],
                    dwData: 0,
                    iString: 0,
                });

                add_button(&mut buttons, ID_PREV_PAGE as i32, 4, "Zurück");
                add_button(&mut buttons, ID_NEXT_PAGE as i32, 5, "Weiter");

                // Separator before Print
                buttons.push(TBBUTTON {
                    iBitmap: 0,
                    idCommand: 0,
                    fsState: 0,
                    fsStyle: BTNS_SEP as u8,
                    bReserved: [0; 6],
                    dwData: 0,
                    iString: 0,
                });

                // Print button
                add_button(&mut buttons, ID_PRINT as i32, 8, "Drucken");

                // Spring Separator
                buttons.push(TBBUTTON {
                    iBitmap: 0,
                    idCommand: ID_SPRING as i32,
                    fsState: 0,
                    fsStyle: BTNS_SEP as u8,
                    bReserved: [0; 6],
                    dwData: 0,
                    iString: 0,
                });

                // Info Button
                add_button(&mut buttons, ID_INFO as i32, 6, "Info");

                // Close Button
                add_button(&mut buttons, ID_CLOSE as i32, 7, "Schließen");
            }
            ToolbarType::Bottom => {
                add_button(&mut buttons, ID_ROTATE_LEFT as i32, 2, "Rotate Left");
                add_button(&mut buttons, ID_ROTATE_RIGHT as i32, 3, "Rotate Right");

                buttons.push(TBBUTTON {
                    iBitmap: 0,
                    idCommand: 0,
                    fsState: 0,
                    fsStyle: BTNS_SEP as u8,
                    bReserved: [0; 6],
                    dwData: 0,
                    iString: 0,
                });

                add_button(&mut buttons, ID_PREV_PAGE as i32, 4, "Prev");
                add_button(&mut buttons, ID_NEXT_PAGE as i32, 5, "Next");
            }
        }

        buttons
    }

    pub fn height(&self) -> i32 {
        unsafe {
            let mut rect = RECT::default();
            let _ = GetWindowRect(self.rebar_hwnd, &mut rect);
            rect.bottom - rect.top
        }
    }

    pub fn resize(&self, parent_width: i32, y: i32) {
        unsafe {
            let mut rebar_rect = RECT::default();
            let _ = GetWindowRect(self.rebar_hwnd, &mut rebar_rect);
            let height = rebar_rect.bottom - rebar_rect.top;
            let effective_height = if height > 0 { height } else { 28 };

            let _ = SetWindowPos(
                self.rebar_hwnd,
                None,
                0,
                y,
                parent_width,
                effective_height,
                SWP_NOZORDER,
            );

            if let ToolbarType::Top = self.toolbar_type {
                // Reset spring
                let mut tbbi_reset = TBBUTTONINFOW {
                    cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                    dwMask: TBIF_SIZE,
                    cx: 0,
                    ..Default::default()
                };
                SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_SPRING as usize), LPARAM(&mut tbbi_reset as *mut _ as isize));
                
                SendMessageW(self.toolbar_hwnd, TB_AUTOSIZE, WPARAM(0), LPARAM(0));

                let button_count = SendMessageW(self.toolbar_hwnd, TB_BUTTONCOUNT, WPARAM(0), LPARAM(0)).0 as usize;
                let mut used_width = 0;
                
                for i in 0..button_count {
                     let mut btn = TBBUTTON::default();
                     SendMessageW(self.toolbar_hwnd, TB_GETBUTTON, WPARAM(i), LPARAM(&mut btn as *mut _ as isize));
                     
                     if btn.idCommand != ID_SPRING as i32 {
                         let mut rect = RECT::default();
                         SendMessageW(self.toolbar_hwnd, TB_GETITEMRECT, WPARAM(i), LPARAM(&mut rect as *mut _ as isize));
                         used_width += rect.right - rect.left;
                     }
                }

                let spring_width = (parent_width - used_width - 15).max(0);
                
                let mut tbbi_spring = TBBUTTONINFOW {
                    cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                    dwMask: TBIF_SIZE,
                    cx: spring_width as u16,
                    ..Default::default()
                };
                SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_SPRING as usize), LPARAM(&mut tbbi_spring as *mut _ as isize));
            } else {
                SendMessageW(self.toolbar_hwnd, TB_AUTOSIZE, WPARAM(0), LPARAM(0));
            }

            let _ = InvalidateRect(self.rebar_hwnd, None, true);
        }
    }

    pub fn poll_command(&self) -> Option<ToolbarCommand> {
        self.pending_command.lock().take()
    }

    pub fn set_dark_theme(&mut self, is_dark: bool) {
        self.is_dark = is_dark;
        unsafe {
            let theme = if is_dark { w!("DarkMode_Explorer") } else { w!("Explorer") };
            let _ = SetWindowTheme(self.rebar_hwnd, theme, None);
            let _ = SetWindowTheme(self.toolbar_hwnd, theme, None);

            let _ = InvalidateRect(self.rebar_hwnd, None, true);
            let _ = InvalidateRect(self.toolbar_hwnd, None, true);
        }
    }

    #[allow(dead_code)]
    pub fn hwnd(&self) -> HWND {
        self.rebar_hwnd
    }

    pub fn set_document_loaded(&self, loaded: bool) {
        unsafe {
            let enable = if loaded { 1isize } else { 0isize };
            // Enable/disable document-dependent buttons
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_EXPORT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ROTATE_LEFT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ROTATE_RIGHT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_PREV_PAGE as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_NEXT_PAGE as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_PRINT as usize), LPARAM(enable));
        }
    }

    pub fn set_open_enabled(&self, enabled: bool) {
        unsafe {
            let enable = if enabled { 1isize } else { 0isize };
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_OPEN as usize), LPARAM(enable));
        }
    }

    pub fn handle_notify(&self, lparam: LPARAM) -> Option<LRESULT> {
        unsafe {
            let nmhdr = &*(lparam.0 as *const NMHDR);

            // Only handle notifications from our toolbar
            if nmhdr.hwndFrom != self.toolbar_hwnd {
                return None;
            }

            // Handle tooltip requests
            if nmhdr.code == TBN_GETINFOTIPW {
                let nmtbgit = &mut *(lparam.0 as *mut NMTBGETINFOTIPW);

                let tooltip_text: &str = match nmtbgit.iItem {
                    x if x == ID_OPEN as i32 => "Datei öffnen (Strg+O)",
                    x if x == ID_EXPORT as i32 => "Exportieren (Strg+E)",
                    x if x == ID_ROTATE_LEFT as i32 => "Nach links drehen (Strg+Links)",
                    x if x == ID_ROTATE_RIGHT as i32 => "Nach rechts drehen (Strg+Rechts)",
                    x if x == ID_PREV_PAGE as i32 => "Vorherige Seite (Bild↑)",
                    x if x == ID_NEXT_PAGE as i32 => "Nächste Seite (Bild↓)",
                    x if x == ID_PRINT as i32 => "Drucken (Strg+P)",
                    x if x == ID_INFO as i32 => "Dokumentinformationen",
                    x if x == ID_CLOSE as i32 => "Programm beenden (Alt+F4)",
                    _ => return None,
                };

                // Convert to wide string and copy to buffer
                let wide: Vec<u16> = tooltip_text.encode_utf16().collect();
                let max_chars = (nmtbgit.cchTextMax as usize).saturating_sub(1);
                let copy_len = wide.len().min(max_chars);

                if !nmtbgit.pszText.is_null() && copy_len > 0 {
                    std::ptr::copy_nonoverlapping(
                        wide.as_ptr(),
                        nmtbgit.pszText.0,
                        copy_len,
                    );
                    // Null-terminate
                    *nmtbgit.pszText.0.add(copy_len) = 0;
                }

                return Some(LRESULT(0));
            }

            None
        }
    }
}

impl Drop for Toolbar {
    fn drop(&mut self) {
        unsafe {
            if self.image_list != HIMAGELIST::default() {
                let _ = ImageList_Destroy(self.image_list);
            }
        }
    }
}