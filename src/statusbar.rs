use crate::icons;
use crate::utils::{load_png_from_memory, make_long};
use parking_lot::Mutex;
use std::sync::Arc;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::*,
        UI::{Controls::*, WindowsAndMessaging::*},
    },
};

// Command IDs
pub const ID_ZOOM_OUT: u16 = 300;
pub const ID_ZOOM_IN: u16 = 301;
pub const ID_ZOOM_FIT: u16 = 302;
pub const ID_ZOOM_HEIGHT: u16 = 303;
pub const ID_ZOOM_WIDTH: u16 = 304;

pub const ID_FILENAME: u16 = 400;
pub const ID_FILEINFO: u16 = 401;
pub const ID_ZOOM_TEXT: u16 = 399;
pub const ID_SPRING: u16 = 9999;
pub const ID_SPRING_RIGHT: u16 = 9998;

pub struct StatusBar {
    rebar_hwnd: HWND,
    toolbar_hwnd: HWND,
    image_list: HIMAGELIST,
    pending_zoom_command: Arc<Mutex<Option<f32>>>,
    current_zoom: f32,
    filename: String,
    info_text: String,
}

impl StatusBar {
    pub fn new(parent: HWND, instance: HMODULE) -> Result<Self> {
        unsafe {
            let rebar_hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("ReBarWindow32"),
                None,
                WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
                WINDOW_STYLE(
                    CCS_NODIVIDER as u32 |
                    CCS_NORESIZE as u32 |
                    CCS_NOPARENTALIGN as u32 |
                    CCS_BOTTOM as u32
                ),
                0,
                0,
                0,
                0,
                parent,
                HMENU(2001isize),
                instance,
                None,
            );

            if rebar_hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

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
                HMENU(2000isize),
                instance,
                None,
            );

            if toolbar_hwnd.0 == 0 {
                return Err(Error::from_win32());
            }

            SendMessageW(
                toolbar_hwnd,
                TB_SETEXTENDEDSTYLE,
                WPARAM(0),
                LPARAM(TBSTYLE_EX_MIXEDBUTTONS as isize),
            );

            let image_list = Self::create_image_list()?;
            SendMessageW(toolbar_hwnd, TB_SETIMAGELIST, WPARAM(0), LPARAM(image_list.0));

            let buttons = Self::create_buttons();
            SendMessageW(
                toolbar_hwnd,
                TB_ADDBUTTONS,
                WPARAM(buttons.len()),
                LPARAM(buttons.as_ptr() as isize),
            );

            SendMessageW(
                toolbar_hwnd,
                TB_SETBUTTONSIZE,
                WPARAM(0),
                LPARAM(make_long(0, 24) as isize),
            );

            SendMessageW(toolbar_hwnd, TB_AUTOSIZE, WPARAM(0), LPARAM(0));

            let mut tb_size = SIZE::default();
            SendMessageW(toolbar_hwnd, TB_GETMAXSIZE, WPARAM(0), LPARAM(&mut tb_size as *mut _ as isize));

            let rbbi = REBARBANDINFOW {
                cbSize: std::mem::size_of::<REBARBANDINFOW>() as u32,
                fMask: RBBIM_STYLE | RBBIM_CHILD | RBBIM_CHILDSIZE | RBBIM_SIZE,
                fStyle: RBBS_CHILDEDGE,
                hwndChild: toolbar_hwnd,
                cxMinChild: tb_size.cx as u32,
                cyMinChild: tb_size.cy as u32,
                cx: tb_size.cx as u32,
                ..Default::default()
            };

            SendMessageW(rebar_hwnd, RB_INSERTBANDW, WPARAM(usize::MAX), LPARAM(&rbbi as *const _ as isize));

            let _ = SetWindowTheme(rebar_hwnd, w!("Explorer"), None);
            let _ = SetWindowTheme(toolbar_hwnd, w!("Explorer"), None);

            Ok(Self {
                rebar_hwnd,
                toolbar_hwnd,
                image_list,
                pending_zoom_command: Arc::new(Mutex::new(None)),
                current_zoom: 1.0,
                filename: String::from("Dateiname |"),
                info_text: String::from("Bildinformation"),
            })
        }
    }

    fn create_image_list() -> Result<HIMAGELIST> {
        unsafe {
            let image_list = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 6, 0);
            if image_list.0 == 0 {
                return Err(Error::from_win32());
            }

            let icon_data: &[&[u8]] = &[
                icons::ICON_ZOOM_OUT,
                icons::ICON_ZOOM_IN,
                icons::ICON_FIT_TO_SIZE,
                icons::ICON_FIT_TO_HEIGHT,
                icons::ICON_FIT_TO_WIDTH,
                icons::ICON_DOCUMENT_EMPTY,
                icons::ICON_DOCUMENT_INFORMATION,
            ];

            for data in icon_data {
                let bitmap = load_png_from_memory(data)?;
                ImageList_Add(image_list, bitmap, HBITMAP::default());
                let _ = DeleteObject(bitmap);
            }

            Ok(image_list)
        }
    }

    fn create_buttons() -> Vec<TBBUTTON> {
        let mut buttons = Vec::new();

        let add_text_button = |buttons: &mut Vec<TBBUTTON>, id: i32, image: i32, text: &str| {
            let text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
            let text_ptr = Box::leak(text_wide.into_boxed_slice()).as_ptr() as isize;
            buttons.push(TBBUTTON {
                iBitmap: image,
                idCommand: id,
                fsState: TBSTATE_ENABLED as u8,
                fsStyle: (BTNS_BUTTON | BTNS_SHOWTEXT | BTNS_AUTOSIZE) as u8,
                bReserved: [0; 6],
                dwData: 0,
                iString: text_ptr,
            });
        };

        let add_icon_button = |buttons: &mut Vec<TBBUTTON>, id: i32, image: i32| {
            buttons.push(TBBUTTON {
                iBitmap: image,
                idCommand: id,
                fsState: TBSTATE_ENABLED as u8,
                fsStyle: BTNS_BUTTON as u8,
                bReserved: [0; 6],
                dwData: 0,
                iString: 0,
            });
        };

        // 0: Filename
        add_text_button(&mut buttons, ID_FILENAME as i32, 5, "Dateiname |");
        
        // 1: File Info
        add_text_button(&mut buttons, ID_FILEINFO as i32, 6, "Bildinformation");

        // 2: Spring Separator (Left)
        buttons.push(TBBUTTON {
            iBitmap: 0,
            idCommand: ID_SPRING as i32,
            fsState: 0,
            fsStyle: BTNS_SEP as u8,
            bReserved: [0; 6],
            dwData: 0,
            iString: 0,
        });

        // 3: Zoom - (Icon + Text "Zoom")
        add_text_button(&mut buttons, ID_ZOOM_OUT as i32, 0, "Zoom");

        // 4: Zoom % (Centered text placeholder)
        add_text_button(&mut buttons, ID_ZOOM_TEXT as i32, -1, " 100 % ");

        // 5: Zoom + (Icon + Text "Zoom")
        add_text_button(&mut buttons, ID_ZOOM_IN as i32, 1, "Zoom");

        // 6: Spring Separator (Right)
        buttons.push(TBBUTTON {
            iBitmap: 0,
            idCommand: ID_SPRING_RIGHT as i32,
            fsState: 0,
            fsStyle: BTNS_SEP as u8,
            bReserved: [0; 6],
            dwData: 0,
            iString: 0,
        });

        // 7: Fit to page (Icon only)
        add_icon_button(&mut buttons, ID_ZOOM_FIT as i32, 2);
        // 8: Fit Horizontal (Icon only)
        add_icon_button(&mut buttons, ID_ZOOM_WIDTH as i32, 4);
        // 9: Fit Vertical (Icon only)
        add_icon_button(&mut buttons, ID_ZOOM_HEIGHT as i32, 3);

        buttons
    }

    pub fn height(&self) -> i32 {
        unsafe {
            let mut rect = RECT::default();
            let _ = GetWindowRect(self.rebar_hwnd, &mut rect);
            let h = rect.bottom - rect.top;
            if h > 0 { h } else { 28 } // Compact height
        }
    }

    pub fn resize(&self, parent_width: i32, parent_height: i32) {
        unsafe {
            let height = self.height();
            let _ = SetWindowPos(self.rebar_hwnd, None, 0, parent_height - height, parent_width, height, SWP_NOZORDER);

            // 1. Reset to autosize to get natural widths
            let mut tbbi_reset = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_SIZE,
                cx: 0,
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILENAME as usize), LPARAM(&mut tbbi_reset as *mut _ as isize));
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILEINFO as usize), LPARAM(&mut tbbi_reset as *mut _ as isize));

            SendMessageW(self.toolbar_hwnd, TB_AUTOSIZE, WPARAM(0), LPARAM(0));

            // 2. Measure sections
            let mut w_left = 0;
            for i in 0..2 {
                let mut r = RECT::default();
                if SendMessageW(self.toolbar_hwnd, TB_GETITEMRECT, WPARAM(i), LPARAM(&mut r as *mut _ as isize)).0 != 0 {
                    w_left += r.right - r.left;
                }
            }

            let mut w_center = 0;
            for i in 3..6 {
                let mut r = RECT::default();
                if SendMessageW(self.toolbar_hwnd, TB_GETITEMRECT, WPARAM(i), LPARAM(&mut r as *mut _ as isize)).0 != 0 {
                    w_center += r.right - r.left;
                }
            }

            let mut w_right = 0;
            for i in 7..10 {
                let mut r = RECT::default();
                if SendMessageW(self.toolbar_hwnd, TB_GETITEMRECT, WPARAM(i), LPARAM(&mut r as *mut _ as isize)).0 != 0 {
                    w_right += r.right - r.left;
                }
            }

            // 3. Constrain Left section if needed
            let fixed_non_spring = w_center + w_right + 40;
            let available_left = (parent_width - fixed_non_spring).max(0);
            
            if w_left > available_left {
                let max_w0 = (available_left as f32 * 0.4) as i32;
                let final_w0 = w_left.min(max_w0);
                let final_w1 = (available_left - final_w0).max(0);
                
                let mut tbbi0 = TBBUTTONINFOW {
                    cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                    dwMask: TBIF_SIZE,
                    cx: final_w0 as u16,
                    ..Default::default()
                };
                SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILENAME as usize), LPARAM(&mut tbbi0 as *mut _ as isize));

                let mut tbbi1 = TBBUTTONINFOW {
                    cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                    dwMask: TBIF_SIZE,
                    cx: final_w1 as u16,
                    ..Default::default()
                };
                SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILEINFO as usize), LPARAM(&mut tbbi1 as *mut _ as isize));
                w_left = final_w0 + final_w1;
            }

            // 4. Calculate Springs to align Zoom group to the right
            let available_for_springs = (parent_width - w_left - w_center - w_right - 20).max(0);
            let sl = available_for_springs;  // Left spring takes all available space
            let sr = 0;                       // Right spring is zero (zoom group aligned right)

            let mut tbbi_sl = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_SIZE,
                cx: sl as u16,
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_SPRING as usize), LPARAM(&mut tbbi_sl as *mut _ as isize));

            let mut tbbi_sr = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_SIZE,
                cx: sr as u16,
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_SPRING_RIGHT as usize), LPARAM(&mut tbbi_sr as *mut _ as isize));
            
            let _ = InvalidateRect(self.rebar_hwnd, None, true);
        }
    }

    pub fn set_zoom(&mut self, zoom: f32) {
        self.current_zoom = zoom;
        let percent = (zoom * 100.0).round() as i32;
        let text = format!(" {:03} % ", percent);
        self.update_zoom_text(&text);
    }

    fn update_zoom_text(&self, text: &str) {
        unsafe {
            let text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
            let text_ptr = Box::leak(text_wide.into_boxed_slice()).as_ptr() as isize;
            let tbbi = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_TEXT,
                pszText: PWSTR(text_ptr as *mut u16),
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_ZOOM_TEXT as usize), LPARAM(&tbbi as *const _ as isize));
            let _ = InvalidateRect(self.toolbar_hwnd, None, true);
        }
    }

    /// Truncates filename to max length, preserving extension with "..." prefix
    /// Example: "VeryLongFileName.pdf" -> "VeryLongFileNam...pdf"
    fn truncate_filename(filename: &str, max_base_len: usize) -> String {
        let path = std::path::Path::new(filename);
        let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
        let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or(filename);

        if stem.chars().count() <= max_base_len {
            filename.to_string()
        } else {
            let truncated: String = stem.chars().take(max_base_len).collect();
            if ext.is_empty() {
                format!("{}...", truncated)
            } else {
                format!("{}...{}", truncated, ext)
            }
        }
    }

    /// Shows filename immediately while document is loading
    pub fn set_loading_file(&mut self, filename: &str) {
        let display_name = Self::truncate_filename(filename, 30);
        self.filename = format!("{} |", display_name);
        self.info_text = String::from(" Lade...");
        self.update_info_display();
    }

    pub fn set_file_info(
        &mut self,
        filename: &str,
        dimensions: &str,
        file_size: u64,
        current_page: usize,
        total_pages: usize,
    ) {
        let display_name = Self::truncate_filename(filename, 30);
        self.filename = format!("{} |", display_name);
        let size_str = if file_size >= 1024 * 1024 {
            format!("{:.1} MB", file_size as f64 / (1024.0 * 1024.0))
        } else if file_size >= 1024 {
            format!("{:.1} KB", file_size as f64 / 1024.0)
        } else {
            format!("{} B", file_size)
        };
        let page_str = if total_pages > 1 {
            format!(" | Page {}/{}", current_page + 1, total_pages)
        } else {
            String::new()
        };
        self.info_text = format!(" {} | {}{}", dimensions, size_str, page_str);
        self.update_info_display();
    }

    fn update_info_display(&self) {
        unsafe {
            let fname_wide: Vec<u16> = self.filename.encode_utf16().chain(std::iter::once(0)).collect();
            let info_wide: Vec<u16> = self.info_text.encode_utf16().chain(std::iter::once(0)).collect();
            let tbbi_f = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_TEXT,
                pszText: PWSTR(fname_wide.as_ptr() as *mut u16),
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILENAME as usize), LPARAM(&tbbi_f as *const _ as isize));
            let tbbi_i = TBBUTTONINFOW {
                cbSize: std::mem::size_of::<TBBUTTONINFOW>() as u32,
                dwMask: TBIF_TEXT,
                pszText: PWSTR(info_wide.as_ptr() as *mut u16),
                ..Default::default()
            };
            SendMessageW(self.toolbar_hwnd, TB_SETBUTTONINFOW, WPARAM(ID_FILEINFO as usize), LPARAM(&tbbi_i as *const _ as isize));
            
             let parent = GetParent(self.rebar_hwnd);
             let mut parent_rect = RECT::default();
             let _ = GetClientRect(parent, &mut parent_rect);
             self.resize(parent_rect.right, parent_rect.bottom);
        }
    }

    pub fn poll_zoom_command(&self) -> Option<f32> {
        self.pending_zoom_command.lock().take()
    }

    pub fn set_dark_theme(&self, _is_dark: bool) {
        unsafe {
            let _ = SetWindowTheme(self.rebar_hwnd, w!("Explorer"), None);
            let _ = SetWindowTheme(self.toolbar_hwnd, w!("Explorer"), None);
            let _ = InvalidateRect(self.rebar_hwnd, None, true);
        }
    }

    #[allow(dead_code)]
    pub fn toolbar_hwnd(&self) -> HWND {
        self.toolbar_hwnd
    }

    pub fn set_document_loaded(&self, loaded: bool) {
        unsafe {
            let enable = if loaded { 1isize } else { 0isize };
            // Enable/disable document-dependent buttons
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_OUT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_IN as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_FIT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_HEIGHT as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_WIDTH as usize), LPARAM(enable));
            SendMessageW(self.toolbar_hwnd, TB_ENABLEBUTTON, WPARAM(ID_ZOOM_TEXT as usize), LPARAM(enable));
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
                    x if x == ID_ZOOM_OUT as i32 => "Verkleinern (- / Strg+Mausrad)",
                    x if x == ID_ZOOM_IN as i32 => "Vergrößern (+ / Strg+Mausrad)",
                    x if x == ID_ZOOM_TEXT as i32 => "Zoom zurücksetzen (/)",
                    x if x == ID_ZOOM_FIT as i32 => "An Fenster anpassen (*)",
                    x if x == ID_ZOOM_HEIGHT as i32 => "An Höhe anpassen",
                    x if x == ID_ZOOM_WIDTH as i32 => "An Breite anpassen",
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

            // Handle custom draw for centered text
            if nmhdr.code == NM_CUSTOMDRAW {
                let nmcd = &*(lparam.0 as *const NMTBCUSTOMDRAW);

                match nmcd.nmcd.dwDrawStage {
                    CDDS_PREPAINT => {
                        return Some(LRESULT(CDRF_NOTIFYITEMDRAW as isize));
                    }
                    CDDS_ITEMPREPAINT => {
                        // Check if this is the zoom text button
                        if nmcd.nmcd.dwItemSpec == ID_ZOOM_TEXT as usize {
                            let hdc = nmcd.nmcd.hdc;
                            let rect = nmcd.nmcd.rc;
                            let item_state = nmcd.nmcd.uItemState;

                            // Draw button background with theme for pressed/hot states
                            let theme = OpenThemeData(self.toolbar_hwnd, w!("Toolbar"));
                            if !theme.is_invalid() {
                                let state = if (item_state.0 & CDIS_SELECTED.0) != 0 {
                                    TS_PRESSED
                                } else if (item_state.0 & CDIS_HOT.0) != 0 {
                                    TS_HOT
                                } else {
                                    TS_NORMAL
                                };

                                let _ = DrawThemeBackground(
                                    theme,
                                    hdc,
                                    TP_BUTTON.0,
                                    state.0,
                                    &rect,
                                    None,
                                );
                                let _ = CloseThemeData(theme);
                            }

                            // Get toolbar font and select it
                            let toolbar_font = HFONT(SendMessageW(
                                self.toolbar_hwnd,
                                WM_GETFONT,
                                WPARAM(0),
                                LPARAM(0),
                            ).0);

                            let old_font = if !toolbar_font.is_invalid() {
                                SelectObject(hdc, toolbar_font)
                            } else {
                                HGDIOBJ::default()
                            };

                            // Get current zoom text
                            let zoom_percent = (self.current_zoom * 100.0).round() as i32;
                            let text = format!("{:03} %", zoom_percent);
                            let mut text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();

                            // Draw centered text
                            SetBkMode(hdc, TRANSPARENT);
                            SetTextColor(hdc, COLORREF(0x00000000)); // Black text

                            let mut draw_rect = rect;
                            DrawTextW(
                                hdc,
                                &mut text_wide,
                                &mut draw_rect,
                                DT_CENTER | DT_VCENTER | DT_SINGLELINE,
                            );

                            // Restore old font
                            if !old_font.is_invalid() {
                                SelectObject(hdc, old_font);
                            }

                            return Some(LRESULT(CDRF_SKIPDEFAULT as isize));
                        }
                        return Some(LRESULT(CDRF_DODEFAULT as isize));
                    }
                    _ => {}
                }
            }
            None
        }
    }
}

impl Drop for StatusBar {
    fn drop(&mut self) {
        unsafe {
            if self.image_list != HIMAGELIST::default() {
                let _ = ImageList_Destroy(self.image_list);
            }
        }
    }
}
