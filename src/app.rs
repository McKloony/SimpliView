use crate::{
    d2d::D2DRenderer,
    dialogs::FileDialogs,
    document::{Document, PageLayout},
    menu::ContextMenu,
    pdf::PdfLoader,
    scroll::{ScrollAction, ScrollManager, LINE_SCROLL_PIXELS, clamp_scroll},
    statusbar::StatusBar,
    theme::Theme,
    toolbar::{Toolbar, ToolbarCommand, ToolbarType},
    view_window::{ViewWindow, WM_APP_VIEW_PAINT},
    wic::WicLoader,
    window::Window,
};
use parking_lot::Mutex;
use std::sync::Arc;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::{Direct2D::Common::*, Gdi::*},
        Storage::FileSystem::*,
        System::{DataExchange::*, Memory::*, Ole::CF_DIB},
        UI::{
            Controls::Dialogs::*,
            Input::KeyboardAndMouse::*,
            WindowsAndMessaging::*,
        },
    },
};

#[allow(dead_code)]
pub const WM_APP_DOCUMENT_LOADED: u32 = WM_APP + 1;
#[allow(dead_code)]
pub const WM_APP_DOCUMENT_ERROR: u32 = WM_APP + 2;

#[derive(Clone)]
pub struct AppState {
    pub document: Option<Document>,
    pub zoom: f32,
    pub rotation: i32, // 0, 90, 180, 270
    pub current_page: usize,
    pub total_pages: usize,
    pub file_path: Option<String>,
    #[allow(dead_code)]
    pub is_dark_theme: bool,
    pub fit_to_page: bool,
    // Folder navigation
    pub folder_files: Vec<String>,
    pub folder_file_index: usize,
    pub folder_navigation_mode: bool, // true = navigate files, false = navigate pages
    // Scroll state
    pub scroll_x: i32,        // Horizontal scroll offset (in scaled pixels)
    pub scroll_y: i32,        // Vertical scroll offset (in scaled pixels)
    pub content_width: i32,   // Total content width at current zoom
    pub content_height: i32,  // Total content height at current zoom
    // Multi-page view mode
    pub multi_page_view: bool,           // true = show all pages stacked, false = single page
    pub page_layout: Option<PageLayout>, // Cached layout for multi-page view
}

impl Default for AppState {
    fn default() -> Self {
        Self {
            document: None,
            zoom: 1.0,
            rotation: 0,
            current_page: 0,
            total_pages: 1,
            file_path: None,
            is_dark_theme: false,
            fit_to_page: true, // Default to fit to page
            folder_files: Vec::new(),
            folder_file_index: 0,
            folder_navigation_mode: true,
            scroll_x: 0,
            scroll_y: 0,
            content_width: 0,
            content_height: 0,
            multi_page_view: true, // Default to multi-page view for PDFs
            page_layout: None,
        }
    }
}

struct WaitCursorGuard {
    previous: Option<HCURSOR>,
}

impl WaitCursorGuard {
    fn new() -> Self {
        unsafe {
            if let Ok(wait_cursor) = LoadCursorW(None, IDC_WAIT) {
                let previous = SetCursor(wait_cursor);
                Self { previous: Some(previous) }
            } else {
                Self { previous: None }
            }
        }
    }
}

impl Drop for WaitCursorGuard {
    fn drop(&mut self) {
        if let Some(prev) = self.previous {
            unsafe {
                SetCursor(prev);
            }
        }
    }
}

pub struct App {
    window: Window,
    view_window: ViewWindow,
    renderer: D2DRenderer,
    top_toolbar: Toolbar,
    statusbar: StatusBar,
    context_menu: ContextMenu,
    wic_loader: WicLoader,
    pdf_loader: PdfLoader,
    dialogs: FileDialogs,
    state: Arc<Mutex<AppState>>,
    scroll_manager: ScrollManager,
    file_to_open: Option<String>,
    opened_from_cmdline: bool,
    open_disabled: bool,
    // Drag-to-pan state
    is_dragging: bool,
    drag_start_mouse: (i32, i32),
    drag_start_scroll: (i32, i32),
}

impl App {
    #[allow(clippy::arc_with_non_send_sync)]
    pub fn new(file_to_open: Option<String>, restricted_path: Option<String>) -> Result<Self> {
        // Always use light mode - using Arc for internal state sharing within App
        let state = Arc::new(Mutex::new(AppState {
            is_dark_theme: false,
            ..Default::default()
        }));

        // Create main window
        let window = Window::new("SimpliView", state.clone())?;

        // Create top toolbar
        let top_toolbar = Toolbar::new(window.hwnd(), window.instance(), ToolbarType::Top)?;

        // Create view window (canvas) for Direct2D rendering
        let view_window = ViewWindow::new(window.hwnd(), window.instance())?;

        // Create status bar
        let statusbar = StatusBar::new(window.hwnd(), window.instance())?;

        // Initialize Direct2D renderer targeting the view window
        let renderer = D2DRenderer::new(view_window.hwnd())?;

        // Create context menu
        let context_menu = ContextMenu::new()?;

        // Initialize image and PDF loaders
        let wic_loader = WicLoader::new()?;
        let pdf_loader = PdfLoader::new();

        // Initialize file dialogs
        let dialogs = FileDialogs::new(restricted_path);

        // Create scroll manager attached to the view window
        let scroll_manager = ScrollManager::new(view_window.hwnd());

        let opened_from_cmdline = file_to_open.is_some();
        let open_disabled = file_to_open.is_some();

        Ok(Self {
            window,
            view_window,
            renderer,
            top_toolbar,
            statusbar,
            context_menu,
            wic_loader,
            pdf_loader,
            dialogs,
            state,
            scroll_manager,
            file_to_open,
            opened_from_cmdline,
            open_disabled,
            is_dragging: false,
            drag_start_mouse: (0, 0),
            drag_start_scroll: (0, 0),
        })
    }

    pub fn run(&mut self) -> Result<()> {
        // Register self pointer with window for message handling
        let hwnd = self.window.hwnd();
        unsafe {
            use windows::Win32::UI::WindowsAndMessaging::{SetWindowLongPtrW, GWLP_USERDATA};
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, self as *mut App as isize);
        }

        // Show and update window
        self.window.show();

        // Apply initial theme
        self.apply_theme();

        // Disable document-dependent buttons until a document is loaded
        self.top_toolbar.set_document_loaded(false);
        self.top_toolbar.set_navigation_enabled(false);
        self.statusbar.set_document_loaded(false);
        self.context_menu.set_document_loaded(false);

        // Disable Open button if file was passed via command line
        if self.open_disabled {
            self.top_toolbar.set_open_enabled(false);
        }

        // If a file was passed via command line, open it
        if let Some(path) = self.file_to_open.take() {
            self.open_document(&path);
        }

        // Main Message loop
        unsafe {
            let mut msg = MSG::default();
            loop {
                // GetMessage blocks until a message arrives
                let ret = GetMessageW(&mut msg, None, 0, 0);
                if ret.0 == 0 || ret.0 == -1 {
                    break;
                }

                // Handle accelerators (keyboard shortcuts)
                if !self.handle_accelerator(&msg) {
                    let _ = TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }

                // Process any pending app logic
                self.process_app_messages();
            }
        }

        // Clear the pointer before exiting
        unsafe {
            use windows::Win32::UI::WindowsAndMessaging::{SetWindowLongPtrW, GWLP_USERDATA};
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        }

        Ok(())
    }

    fn handle_accelerator(&mut self, msg: &MSG) -> bool {
        if msg.message == WM_KEYDOWN {
            let key = VIRTUAL_KEY(msg.wParam.0 as u16);
            let ctrl = unsafe { GetKeyState(VK_CONTROL.0 as i32) } < 0;

            match (ctrl, key) {
                // Ctrl+O -> Open (disabled if file was passed via command line)
                (true, VK_O) => { if !self.open_disabled { self.cmd_open(); } return true; }
                // Ctrl+E -> Export
                (true, VK_E) => { self.cmd_export(); return true; }
                // Ctrl+C -> Copy to clipboard
                (true, VK_C) => { self.cmd_copy_to_clipboard(); return true; }
                // Ctrl+P -> Print
                (true, VK_P) => { self.cmd_print(); return true; }
                // Ctrl+Left -> Rotate left
                (true, VK_LEFT) => { self.cmd_rotate_left(); return true; }
                // Ctrl+Right -> Rotate right
                (true, VK_RIGHT) => { self.cmd_rotate_right(); return true; }
                // Numpad + -> Zoom in
                (false, VK_ADD) => { self.cmd_zoom_in(); return true; }
                // Numpad - -> Zoom out
                (false, VK_SUBTRACT) => { self.cmd_zoom_out(); return true; }
                // Numpad / -> Zoom 100%
                (false, VK_DIVIDE) => { self.cmd_zoom_reset(); return true; }
                // Numpad * -> Fit to page
                (false, VK_MULTIPLY) => { self.cmd_fit_to_page(); return true; }
                // Navigation
                (false, VK_LEFT) | (false, VK_PRIOR) => { self.cmd_prev_page(); return true; }
                (false, VK_RIGHT) | (false, VK_NEXT) => { self.cmd_next_page(); return true; }
                (false, VK_HOME) => { self.cmd_first_page(); return true; }
                (false, VK_END) => { self.cmd_last_page(); return true; }
                _ => {}
            }
        }
        // Handle character input for shortcuts that might be mapped to chars
        if msg.message == WM_CHAR {
            let ch = msg.wParam.0 as u32;
            match ch {
                0x2B => { self.cmd_zoom_in(); return true; } // '+'
                0x2D => { self.cmd_zoom_out(); return true; } // '-'
                0x2F => { self.cmd_zoom_reset(); return true; } // '/'
                0x2A => { self.cmd_fit_to_page(); return true; } // '*'
                _ => {}
            }
        }
        false
    }

    fn process_app_messages(&mut self) {
        // Check for top toolbar commands
        if let Some(cmd) = self.top_toolbar.poll_command() {
            self.execute_toolbar_command(cmd);
        }

        // Check for statusbar zoom commands
        if let Some(zoom_delta) = self.statusbar.poll_zoom_command() {
            if zoom_delta > 0.0 {
                self.cmd_zoom_in();
            } else if zoom_delta < 0.0 {
                self.cmd_zoom_out();
            } else {
                self.cmd_fit_to_page();
            }
        }

        // Check for context menu commands
        if let Some(cmd) = self.context_menu.poll_command() {
            match cmd {
                0 => self.cmd_fit_to_page(),
                1 => self.cmd_rotate_left(),
                2 => self.cmd_rotate_right(),
                _ => {}
            }
        }
    }

    fn execute_toolbar_command(&mut self, cmd: ToolbarCommand) {
        match cmd {
            ToolbarCommand::Open => self.cmd_open(),
            ToolbarCommand::Export => self.cmd_export(),
            ToolbarCommand::RotateLeft => self.cmd_rotate_left(),
            ToolbarCommand::RotateRight => self.cmd_rotate_right(),
            ToolbarCommand::PrevPage => self.cmd_prev_page(),
            ToolbarCommand::NextPage => self.cmd_next_page(),
            ToolbarCommand::Print => self.cmd_print(),
            ToolbarCommand::Info => self.cmd_info(),
            ToolbarCommand::Close => self.cmd_close(),
        }
    }

    pub fn handle_window_message(&mut self, msg: u32, wparam: WPARAM, lparam: LPARAM) -> Option<LRESULT> {
        match msg {
            WM_SIZE => {
                let width = (lparam.0 & 0xFFFF) as i32;
                let height = ((lparam.0 >> 16) & 0xFFFF) as i32;
                self.on_resize(width, height);
                Some(LRESULT(0))
            }
            WM_PAINT => {
                // Main window paint - just validate. View window handles its own paint.
                unsafe {
                     let _ = ValidateRect(self.window.hwnd(), None);
                }
                Some(LRESULT(0))
            }
            WM_APP_VIEW_PAINT => {
                // Custom message from ViewWindow requesting paint
                self.on_paint();
                Some(LRESULT(0))
            }
            WM_CONTEXTMENU => {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                self.context_menu.show(self.window.hwnd(), x, y);
                Some(LRESULT(0))
            }
            WM_MOUSEWHEEL => {
                self.handle_mouse_wheel(wparam);
                Some(LRESULT(0))
            }
            WM_HSCROLL => {
                self.handle_scroll(SB_HORZ, wparam);
                Some(LRESULT(0))
            }
            WM_VSCROLL => {
                self.handle_scroll(SB_VERT, wparam);
                Some(LRESULT(0))
            }
            WM_LBUTTONDOWN => {
                self.handle_lbutton_down(lparam);
                Some(LRESULT(0))
            }
            WM_LBUTTONUP => {
                self.handle_lbutton_up();
                Some(LRESULT(0))
            }
            WM_MOUSEMOVE => {
                self.handle_mouse_move(lparam);
                Some(LRESULT(0))
            }
            WM_CAPTURECHANGED => {
                self.handle_capture_changed();
                Some(LRESULT(0))
            }
            WM_SETCURSOR => {
                if self.handle_set_cursor(lparam) {
                    Some(LRESULT(1))
                } else {
                    None
                }
            }
            WM_COMMAND => {
                let cmd_id = (wparam.0 & 0xFFFF) as u16;
                self.handle_command(cmd_id);
                Some(LRESULT(0))
            }
            WM_NOTIFY => {
                if let Some(result) = self.statusbar.handle_notify(lparam) {
                    return Some(result);
                }
                if let Some(result) = self.top_toolbar.handle_notify(lparam) {
                    return Some(result);
                }
                None
            }
            WM_DPICHANGED => {
                self.on_dpi_changed(lparam);
                Some(LRESULT(0))
            }
            WM_SETTINGCHANGE => {
                Some(LRESULT(0))
            }
            _ => None,
        }
    }

    fn handle_command(&mut self, cmd_id: u16) {
        match cmd_id {
            100 => self.cmd_open(),
            101 => self.cmd_export(),
            102 => self.cmd_rotate_left(),
            103 => self.cmd_rotate_right(),
            104 => self.cmd_prev_page(),
            105 => self.cmd_next_page(),
            106 => self.cmd_info(),
            107 => self.cmd_close(),
            108 => self.cmd_print(),
            // Context menu commands
            200 => self.cmd_fit_to_page(),
            201 => self.cmd_rotate_left(),
            202 => self.cmd_rotate_right(),
            // Statusbar zoom commands
            300 => self.cmd_zoom_out(),
            301 => self.cmd_zoom_in(),
            302 => self.cmd_fit_to_page(),
            303 => self.cmd_fit_to_height(),
            304 => self.cmd_fit_to_width(),
            399 => self.cmd_zoom_reset(),
            _ => {}
        }
    }

    fn on_resize(&mut self, width: i32, height: i32) {
        // Layout:
        // [Top Toolbar]
        // [View Window]
        // [Status Bar]

        let top_height = self.top_toolbar.height();
        let status_height = self.statusbar.height();
        let view_height = (height - top_height - status_height).max(0);

        self.top_toolbar.resize(width, 0);
        self.view_window.resize(0, top_height, width, view_height);
        self.statusbar.resize(width, height);

        if width > 0 && view_height > 0 {
            let _ = self.renderer.resize(width as u32, view_height as u32);
        }

        // If fit-to-page is enabled, we must recalculate zoom when window size changes
        if self.state.lock().fit_to_page {
            self.calculate_fit_zoom();
        }

        // Update content size (scrolling range) and scrollbars
        self.update_content_size();

        self.invalidate();
    }

    fn on_paint(&mut self) {
        // Lock state once for the frame
        let state = self.state.lock().clone();

        if self.renderer.begin_draw().is_ok() {
            // Clear background (anthracite)
            let bg_color = D2D1_COLOR_F { r: 0.22, g: 0.23, b: 0.25, a: 1.0 };
            self.renderer.clear(bg_color);

            if let Some(ref doc) = state.document {
                // Use multi-page view for documents with multiple pages
                if state.multi_page_view && state.total_pages > 1 {
                    if let Some(ref layout) = state.page_layout {
                        let _ = self.renderer.draw_document_multipage(
                            doc,
                            layout,
                            state.zoom,
                            state.rotation,
                            state.scroll_x,
                            state.scroll_y,
                        );
                    }
                } else {
                    // Single page view (original behavior)
                    let _ = self.renderer.draw_document(
                        doc,
                        state.zoom,
                        state.rotation,
                        state.current_page,
                        state.scroll_x,
                        state.scroll_y,
                    );
                }
            }

            // Draw 1px separator line at the bottom (above statusbar)
            let separator_color = D2D1_COLOR_F { r: 0.75, g: 0.75, b: 0.75, a: 1.0 };
            self.renderer.draw_bottom_separator(separator_color);

            let _ = self.renderer.end_draw();
        }
    }

    fn on_dpi_changed(&mut self, lparam: LPARAM) {
        unsafe {
            let rect = &*(lparam.0 as *const RECT);
            let _ = SetWindowPos(
                self.window.hwnd(),
                None,
                rect.left,
                rect.top,
                rect.right - rect.left,
                rect.bottom - rect.top,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
        self.invalidate();
    }

    fn invalidate(&self) {
        self.view_window.invalidate();
    }

    /// Handle WM_HSCROLL / WM_VSCROLL messages
    fn handle_scroll(&mut self, bar: SCROLLBAR_CONSTANTS, wparam: WPARAM) {
        let scroll_code = (wparam.0 & 0xFFFF) as u16;

        let action = match ScrollAction::from_scroll_code(scroll_code) {
            Some(a) => a,
            None => return,
        };

        if action == ScrollAction::EndScroll {
            return;
        }

        let (render_w, render_h) = self.renderer.size();
        let viewport_width = render_w as i32;
        let viewport_height = render_h as i32;

        let mut state = self.state.lock();
        let is_multipage = state.multi_page_view && state.total_pages > 1;
        let is_vertical = bar == SB_VERT;

        let (content_size, viewport_size, scroll_pos) = if bar == SB_HORZ {
            (state.content_width, viewport_width, &mut state.scroll_x)
        } else {
            (state.content_height, viewport_height, &mut state.scroll_y)
        };

        let track_pos = if action == ScrollAction::ThumbTrack || action == ScrollAction::ThumbPosition {
            self.scroll_manager.get_track_pos(bar)
        } else {
            0
        };

        let new_pos = ScrollManager::calculate_new_pos(
            action,
            *scroll_pos,
            viewport_size,
            content_size,
            viewport_size,
            track_pos,
        );

        if new_pos != *scroll_pos {
            *scroll_pos = new_pos;
            drop(state); // Drop lock before update

            self.scroll_manager.set_pos(bar, new_pos);
            self.invalidate();

            // In multi-page mode, update the current page indicator on vertical scroll
            if is_multipage && is_vertical {
                self.update_current_page_from_scroll();
            }
        }
    }

    fn handle_mouse_wheel(&mut self, wparam: WPARAM) {
        let delta = ((wparam.0 >> 16) & 0xFFFF) as i16;
        let ctrl_down = unsafe { GetKeyState(VK_CONTROL.0 as i32) } < 0;
        let shift_down = unsafe { GetKeyState(VK_SHIFT.0 as i32) } < 0;

        if ctrl_down {
            // Ctrl+Wheel = Zoom
            if delta > 0 { self.cmd_zoom_in(); } else { self.cmd_zoom_out(); }
            return;
        }

        // Get system wheel scroll lines
        let scroll_lines = unsafe {
            let mut lines: u32 = 3;
            let _ = SystemParametersInfoW(
                SPI_GETWHEELSCROLLLINES,
                0,
                Some(&mut lines as *mut _ as *mut _),
                SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
            );
            lines.max(1) as i32
        };

        let notches = delta as i32 / 120;
        let scroll_amount = -notches * scroll_lines * LINE_SCROLL_PIXELS;

        let (render_w, render_h) = self.renderer.size();
        let viewport_width = render_w as i32;
        let viewport_height = render_h as i32;

        let mut state = self.state.lock();
        let is_multipage = state.multi_page_view && state.total_pages > 1;

        if shift_down {
            // Shift+Wheel = Horizontal scroll
            let max_x = (state.content_width - viewport_width).max(0);
            let new_x = (state.scroll_x + scroll_amount).clamp(0, max_x);
            if new_x != state.scroll_x {
                state.scroll_x = new_x;
                drop(state);
                self.scroll_manager.set_pos(SB_HORZ, new_x);
                self.invalidate();
            }
        } else {
            // Normal Wheel = Vertical scroll
            let max_y = (state.content_height - viewport_height).max(0);
            let new_y = (state.scroll_y + scroll_amount).clamp(0, max_y);
            if new_y != state.scroll_y {
                state.scroll_y = new_y;
                drop(state);
                self.scroll_manager.set_pos(SB_VERT, new_y);
                self.invalidate();

                // In multi-page mode, update the current page indicator
                if is_multipage {
                    self.update_current_page_from_scroll();
                }
            }
        }
    }

    /// Handle left mouse button down - start drag-to-pan if content is scrollable
    fn handle_lbutton_down(&mut self, lparam: LPARAM) {
        let x = (lparam.0 & 0xFFFF) as i16 as i32;
        let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;

        // Check if content is larger than viewport (scrolling is possible)
        let (viewport_width, viewport_height) = self.renderer.size();
        let state = self.state.lock();
        let can_scroll_h = state.content_width > viewport_width as i32;
        let can_scroll_v = state.content_height > viewport_height as i32;

        if can_scroll_h || can_scroll_v {
            // Start dragging
            self.is_dragging = true;
            self.drag_start_mouse = (x, y);
            self.drag_start_scroll = (state.scroll_x, state.scroll_y);
            drop(state);

            // Capture mouse to receive events even outside window
            unsafe {
                SetCapture(self.view_window.hwnd());
                // Change cursor to grabbing hand (IDC_SIZEALL for all-directions movement)
                if let Ok(cursor) = LoadCursorW(None, IDC_SIZEALL) {
                    SetCursor(cursor);
                }
            }
        }
    }

    /// Handle left mouse button up - end drag-to-pan
    fn handle_lbutton_up(&mut self) {
        if self.is_dragging {
            self.is_dragging = false;

            // Release mouse capture
            unsafe {
                let _ = ReleaseCapture();
                // Restore arrow cursor
                if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                    SetCursor(cursor);
                }
            }
        }
    }

    /// Handle mouse move - pan if dragging
    fn handle_mouse_move(&mut self, lparam: LPARAM) {
        if !self.is_dragging {
            return;
        }

        let x = (lparam.0 & 0xFFFF) as i16 as i32;
        let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;

        // Calculate delta from drag start
        let delta_x = self.drag_start_mouse.0 - x;
        let delta_y = self.drag_start_mouse.1 - y;

        // Calculate new scroll position (inverted - dragging down should scroll up)
        let (viewport_width, viewport_height) = self.renderer.size();
        let mut state = self.state.lock();

        let max_x = (state.content_width - viewport_width as i32).max(0);
        let max_y = (state.content_height - viewport_height as i32).max(0);

        let new_x = (self.drag_start_scroll.0 + delta_x).clamp(0, max_x);
        let new_y = (self.drag_start_scroll.1 + delta_y).clamp(0, max_y);

        let changed = new_x != state.scroll_x || new_y != state.scroll_y;
        let is_multipage = state.multi_page_view && state.total_pages > 1;

        if changed {
            state.scroll_x = new_x;
            state.scroll_y = new_y;
            drop(state);

            self.scroll_manager.set_pos(SB_HORZ, new_x);
            self.scroll_manager.set_pos(SB_VERT, new_y);
            self.invalidate();

            // In multi-page mode, update the current page indicator
            if is_multipage {
                self.update_current_page_from_scroll();
            }
        }
    }

    fn handle_capture_changed(&mut self) {
        if self.is_dragging {
            self.is_dragging = false;
            // Capture is already lost/changed, just reset cursor
            unsafe {
                if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                    SetCursor(cursor);
                }
            }
        }
    }

    fn handle_set_cursor(&self, _lparam: LPARAM) -> bool {
        if self.is_dragging {
            unsafe {
                if let Ok(cursor) = LoadCursorW(None, IDC_SIZEALL) {
                    SetCursor(cursor);
                    return true;
                }
            }
        }
        false
    }

    /// Update current_page based on scroll position (for multi-page mode)
    fn update_current_page_from_scroll(&mut self) {
        let most_visible = self.get_most_visible_page();
        let mut state = self.state.lock();

        if state.current_page != most_visible {
            state.current_page = most_visible;
            let total = state.total_pages;
            let path = state.file_path.clone();
            drop(state);

            // Update status bar with current page
            self.update_page_display(most_visible, total, path.as_deref());
        }
    }

    /// Update content size based on document, zoom, and rotation.
    /// Manages scrollbar visibility and range.
    fn update_content_size(&mut self) {
        let (render_w, render_h) = self.renderer.size();
        let viewport_width = render_w as i32;
        let viewport_height = render_h as i32;

        let mut state = self.state.lock();

        if let Some(ref doc) = state.document {
            // Check if we should use multi-page view
            let use_multipage = state.multi_page_view && state.total_pages > 1;

            if use_multipage {
                // Multi-page view: compute full document layout
                let layout = doc.compute_layout(state.zoom, state.rotation);
                state.content_width = layout.max_width;
                state.content_height = layout.total_height;
                state.page_layout = Some(layout);
            } else {
                // Single page view: use current page dimensions
                let (doc_width, doc_height) = doc.page_dimensions(state.current_page);

                // Determine dimensions based on rotation
                let (w, h) = if state.rotation == 90 || state.rotation == 270 {
                    (doc_height, doc_width)
                } else {
                    (doc_width, doc_height)
                };

                // Calculate effective content size in pixels
                state.content_width = (w * state.zoom) as i32;
                state.content_height = (h * state.zoom) as i32;
                state.page_layout = None;
            }

            // Clamp scroll positions to prevent scrolling past content
            state.scroll_x = clamp_scroll(state.scroll_x, viewport_width, state.content_width);
            state.scroll_y = clamp_scroll(state.scroll_y, viewport_height, state.content_height);

            let scroll_x = state.scroll_x;
            let scroll_y = state.scroll_y;
            let content_w = state.content_width;
            let content_h = state.content_height;
            let fit_to_page = state.fit_to_page;
            drop(state);

            // Update scrollbars:
            // - If "Fit to Page" is active in single-page mode, hide scrollbars
            // - In multi-page mode, always show vertical scrollbar if content exceeds viewport
            // - Otherwise, update based on content vs viewport size
            if fit_to_page && content_h <= viewport_height && content_w <= viewport_width {
                self.scroll_manager.hide_both();
            } else {
                self.scroll_manager.update_both(
                    viewport_width,
                    viewport_height,
                    content_w,
                    content_h,
                    scroll_x,
                    scroll_y,
                );
            }
        } else {
            // No document loaded
            state.content_width = 0;
            state.content_height = 0;
            state.scroll_x = 0;
            state.scroll_y = 0;
            state.page_layout = None;
            drop(state);
            self.scroll_manager.hide_both();
        }
    }

    fn apply_theme(&mut self) {
        Theme::apply_to_window(self.window.hwnd(), false);
        self.top_toolbar.set_dark_theme(false);
        self.statusbar.set_dark_theme(false);
        self.invalidate();
    }

    // --- Command Handlers ---

    fn cmd_info(&self) {
        crate::dialogs::show_info(
            self.window.hwnd(),
            "SimpliView",
            "SimpliView - Release 1.1.1\n\n© 2026 SimpliMed GmbH\n\nwww.simplimed.de",
        );
    }

    fn cmd_close(&self) { unsafe { let _ = PostMessageW(self.window.hwnd(), WM_CLOSE, WPARAM(0), LPARAM(0)); } }

    fn cmd_print(&mut self) {
        let (doc, current_page, rotation, file_path, total_pages) = {
            let state = self.state.lock();
            if let Some(ref doc) = state.document {
                (doc.clone(), state.current_page, state.rotation, state.file_path.clone(), state.total_pages)
            } else {
                return;
            }
        };

        // GDI print functions - manually linked since windows 0.48 doesn't expose them
        #[link(name = "gdi32")]
        extern "system" {
            fn StartDocW(hdc: HDC, lpdi: *const DOCINFOW) -> i32;
            fn EndDoc(hdc: HDC) -> i32;
            fn StartPage(hdc: HDC) -> i32;
            fn EndPage(hdc: HDC) -> i32;
        }

        #[repr(C)]
        struct DOCINFOW {
            cb_size: i32,
            lpsz_doc_name: PCWSTR,
            lpsz_output: PCWSTR,
            lpsz_datatype: PCWSTR,
            fw_type: u32,
        }

        unsafe {
            // Prepare PRINTDLGW structure
            let mut pd: PRINTDLGW = std::mem::zeroed();
            pd.lStructSize = std::mem::size_of::<PRINTDLGW>() as u32;
            pd.hwndOwner = self.window.hwnd();
            // Enable page numbers, disable selection (we don't support text selection)
            pd.Flags = PD_RETURNDC | PD_NOSELECTION | PD_USEDEVMODECOPIESANDCOLLATE;
            pd.nCopies = 1;
            pd.nMinPage = 1;
            pd.nMaxPage = total_pages as u16;
            pd.nFromPage = (current_page + 1) as u16;
            pd.nToPage = (current_page + 1) as u16;

            // Show print dialog
            if !PrintDlgW(&mut pd).as_bool() {
                // User cancelled or error
                return;
            }

            // Get the printer DC
            let hdc = pd.hDC;
            if hdc.is_invalid() {
                self.show_error("Drucker-Gerätekontext konnte nicht abgerufen werden");
                return;
            }

            // Determine page range
            let (start_page, end_page) = if (pd.Flags & PD_PAGENUMS) == PD_PAGENUMS {
                // User selected range (1-based to 0-based)
                let from = pd.nFromPage.max(1).min(total_pages as u16) as usize;
                let to = pd.nToPage.max(from as u16).min(total_pages as u16) as usize;
                (from - 1, to - 1)
            } else {
                // All pages
                (0, total_pages - 1)
            };

            // Show wait cursor as printing might take time
            let _wait = WaitCursorGuard::new();

            // Prepare document name
            let doc_name = file_path
                .as_ref()
                .and_then(|p| std::path::Path::new(p).file_name())
                .and_then(|n| n.to_str())
                .unwrap_or("SimpliView Document");
            let doc_name_wide: Vec<u16> = doc_name.encode_utf16().chain(std::iter::once(0)).collect();

            // Start document
            let doc_info = DOCINFOW {
                cb_size: std::mem::size_of::<DOCINFOW>() as i32,
                lpsz_doc_name: PCWSTR(doc_name_wide.as_ptr()),
                lpsz_output: PCWSTR::null(),
                lpsz_datatype: PCWSTR::null(),
                fw_type: 0,
            };

            if StartDocW(hdc, &doc_info) <= 0 {
                self.show_error("Druckauftrag konnte nicht gestartet werden");
                // Free memory allocated by PrintDlgW
                if !pd.hDevMode.is_invalid() {
                    let _ = GlobalFree(pd.hDevMode);
                }
                if !pd.hDevNames.is_invalid() {
                    let _ = GlobalFree(pd.hDevNames);
                }
                return;
            }

            // Get printable area dimensions (in pixels at printer resolution)
            let page_width = GetDeviceCaps(hdc, HORZRES);
            let page_height = GetDeviceCaps(hdc, VERTRES);

            let mut success = true;

            // Loop through pages
            for page_idx in start_page..=end_page {
                // Get bitmap data for printing
                let bitmap_data = match self.wic_loader.get_bitmap_for_clipboard(&doc, page_idx, rotation) {
                    Ok(data) => data,
                    Err(e) => {
                        self.show_error(&format!("Seite {} konnte nicht zum Drucken vorbereitet werden: {:?}", page_idx + 1, e));
                        success = false;
                        break;
                    }
                };

                // Start page
                if StartPage(hdc) <= 0 {
                    success = false;
                    break;
                }

                // Calculate scaling to fit image on page while preserving aspect ratio
                let img_width = bitmap_data.width as i32;
                let img_height = bitmap_data.height as i32;

                let scale_x = page_width as f64 / img_width as f64;
                let scale_y = page_height as f64 / img_height as f64;
                let scale = scale_x.min(scale_y);

                let dest_width = (img_width as f64 * scale) as i32;
                let dest_height = (img_height as f64 * scale) as i32;

                // Center on page
                let dest_x = (page_width - dest_width) / 2;
                let dest_y = (page_height - dest_height) / 2;

                // Prepare bitmap info header for StretchDIBits
                let bmi = BITMAPINFO {
                    bmiHeader: BITMAPINFOHEADER {
                        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                        biWidth: img_width,
                        biHeight: img_height, // Positive for bottom-up DIB
                        biPlanes: 1,
                        biBitCount: 32,
                        biCompression: BI_RGB.0 as u32,
                        biSizeImage: bitmap_data.data.len() as u32,
                        biXPelsPerMeter: 0,
                        biYPelsPerMeter: 0,
                        biClrUsed: 0,
                        biClrImportant: 0,
                    },
                    bmiColors: [RGBQUAD::default()],
                };

                // Set stretch mode for better quality
                SetStretchBltMode(hdc, HALFTONE);
                SetBrushOrgEx(hdc, 0, 0, None);

                // Draw the bitmap
                let result = StretchDIBits(
                    hdc,
                    dest_x,
                    dest_y,
                    dest_width,
                    dest_height,
                    0,
                    0,
                    img_width,
                    img_height,
                    Some(bitmap_data.data.as_ptr() as *const _),
                    &bmi,
                    DIB_RGB_COLORS,
                    SRCCOPY,
                );

                if result == 0 {
                    success = false;
                    EndPage(hdc);
                    break;
                }

                if EndPage(hdc) <= 0 {
                    success = false;
                    break;
                }
            }

            // End document
            if success {
                EndDoc(hdc);
            } else {
                // AbortDoc(hdc) would be better here if available, but EndDoc is safe enough
                EndDoc(hdc);
                self.show_error("Drucken fehlgeschlagen oder abgebrochen");
            }

            // Free memory allocated by PrintDlgW
            if !pd.hDevMode.is_invalid() {
                let _ = GlobalFree(pd.hDevMode);
            }
            if !pd.hDevNames.is_invalid() {
                let _ = GlobalFree(pd.hDevNames);
            }
        }
    }

    fn cmd_open(&mut self) {
        // Ignore if open is disabled (file was passed via command line)
        if self.open_disabled {
            return;
        }
        if let Some(path) = self.dialogs.open_file(self.window.hwnd()) {
            self.open_document(&path);
        }
    }

    fn cmd_export(&mut self) {
        let state = self.state.lock();
        if state.document.is_some() {
            let file_path = state.file_path.clone();
            drop(state);

            let (current_filename, extension) = if let Some(ref p) = file_path {
                let path = std::path::Path::new(p);
                let filename = path.file_name().and_then(|n| n.to_str()).map(|s| s.to_string());
                let ext = path.extension().and_then(|e| e.to_str()).map(|s| s.to_string());
                (filename, ext)
            } else { (None, None) };

            if let Some(path) = self.dialogs.save_file(
                self.window.hwnd(),
                current_filename.as_deref(),
                extension.as_deref(),
            ) {
                self.export_document(&path);
            }
        }
    }

    fn cmd_rotate_left(&mut self) {
        {
            let mut state = self.state.lock();
            state.rotation = (state.rotation + 270) % 360;
            state.scroll_x = 0;
            state.scroll_y = 0;
        }
        if self.state.lock().fit_to_page {
            self.calculate_fit_zoom();
        }
        self.update_content_size();
        self.invalidate();
    }

    fn cmd_rotate_right(&mut self) {
        {
            let mut state = self.state.lock();
            state.rotation = (state.rotation + 90) % 360;
            state.scroll_x = 0;
            state.scroll_y = 0;
        }
        if self.state.lock().fit_to_page {
            self.calculate_fit_zoom();
        }
        self.update_content_size();
        self.invalidate();
    }

    fn cmd_prev_page(&mut self) {
        let state = self.state.lock();
        let current_page = state.current_page;
        let _total_pages = state.total_pages;
        let folder_files = state.folder_files.clone();
        let folder_index = state.folder_file_index;
        let folder_mode = state.folder_navigation_mode;
        let is_multipage = state.multi_page_view && state.total_pages > 1;
        drop(state);

        if folder_mode {
            if folder_files.len() > 1 && folder_index > 0 {
                let prev_file = folder_files[folder_index - 1].clone();
                self.open_document_with_mode(&prev_file, true);
            }
        } else if current_page > 0 {
            if is_multipage {
                // In multi-page mode, scroll to previous page
                self.scroll_to_page(current_page - 1);
            } else {
                // Single page mode - switch page
                {
                    let mut state = self.state.lock();
                    state.current_page -= 1;
                    state.scroll_x = 0;
                    state.scroll_y = 0;
                }
                self.update_page_display_and_repaint();
            }
        }
    }

    fn cmd_next_page(&mut self) {
        let state = self.state.lock();
        let current_page = state.current_page;
        let total_pages = state.total_pages;
        let folder_files = state.folder_files.clone();
        let folder_index = state.folder_file_index;
        let folder_mode = state.folder_navigation_mode;
        let is_multipage = state.multi_page_view && state.total_pages > 1;
        drop(state);

        if folder_mode {
            if folder_files.len() > 1 && folder_index < folder_files.len() - 1 {
                let next_file = folder_files[folder_index + 1].clone();
                self.open_document_with_mode(&next_file, true);
            }
        } else if current_page < total_pages - 1 {
            if is_multipage {
                // In multi-page mode, scroll to next page
                self.scroll_to_page(current_page + 1);
            } else {
                // Single page mode - switch page
                {
                    let mut state = self.state.lock();
                    state.current_page += 1;
                    state.scroll_x = 0;
                    state.scroll_y = 0;
                }
                self.update_page_display_and_repaint();
            }
        }
    }

    fn cmd_first_page(&mut self) {
        let state = self.state.lock();
        let is_multipage = state.multi_page_view && state.total_pages > 1;
        let current_page = state.current_page;
        drop(state);

        if current_page == 0 { return; }

        if is_multipage {
            self.scroll_to_page(0);
        } else {
            {
                let mut state = self.state.lock();
                state.current_page = 0;
                state.scroll_x = 0;
                state.scroll_y = 0;
            }
            self.update_page_display_and_repaint();
        }
    }

    fn cmd_last_page(&mut self) {
        let state = self.state.lock();
        let last = state.total_pages - 1;
        let is_multipage = state.multi_page_view && state.total_pages > 1;
        let current_page = state.current_page;
        drop(state);

        if current_page == last { return; }

        if is_multipage {
            self.scroll_to_page(last);
        } else {
            {
                let mut state = self.state.lock();
                state.current_page = last;
                state.scroll_x = 0;
                state.scroll_y = 0;
            }
            self.update_page_display_and_repaint();
        }
    }

    /// Scroll to bring a specific page into view (for multi-page mode)
    fn scroll_to_page(&mut self, page: usize) {
        let mut state = self.state.lock();

        if let Some(ref layout) = state.page_layout {
            if page < layout.page_tops.len() {
                // Scroll to the top of the requested page
                let target_y = layout.page_tops[page];

                // Clamp to valid range
                let (_, render_h) = self.renderer.size();
                let viewport_height = render_h as i32;
                let max_y = (state.content_height - viewport_height).max(0);
                let new_y = target_y.clamp(0, max_y);

                state.scroll_y = new_y;
                state.current_page = page;
                let total = state.total_pages;
                let path = state.file_path.clone();
                drop(state);

                self.scroll_manager.set_pos(SB_VERT, new_y);
                self.update_page_display(page, total, path.as_deref());
                self.invalidate();

                return;
            }
        }

        // Fallback: just update current page
        state.current_page = page;
        drop(state);
        self.update_page_display_and_repaint();
    }

    fn update_page_display_and_repaint(&mut self) {
        let state = self.state.lock();
        let page = state.current_page;
        let total = state.total_pages;
        let path = state.file_path.clone();
        drop(state);
        self.update_page_display(page, total, path.as_deref());
        self.update_content_size();
        self.invalidate();
    }

    // Zoom levels - finer increments for smoother Ctrl+Wheel zooming
    const ZOOM_LEVELS: &'static [f32] = &[
        0.10, 0.125, 0.15, 0.175, 0.20, 0.25, 0.33, 0.40, 0.50, 0.60, 0.67, 0.75, 0.85,
        1.00, 1.10, 1.25, 1.40, 1.50, 1.75, 2.00, 2.50, 3.00, 4.00, 5.00, 6.00, 8.00, 10.00
    ];

    fn cmd_zoom_in(&mut self) {
        let mut state = self.state.lock();
        let current = state.zoom;
        let min_zoom = current * 1.10;
        let new_zoom = Self::ZOOM_LEVELS.iter().find(|&&z| z >= min_zoom).copied().unwrap_or(10.0);
        state.zoom = new_zoom;
        state.fit_to_page = false;
        drop(state);
        self.statusbar.set_zoom(new_zoom);
        self.update_content_size();
        self.invalidate();
    }

    fn cmd_zoom_out(&mut self) {
        let mut state = self.state.lock();
        let current = state.zoom;
        let max_zoom = current / 1.10;
        let new_zoom = Self::ZOOM_LEVELS.iter().rev().find(|&&z| z <= max_zoom).copied().unwrap_or(0.10);
        state.zoom = new_zoom;
        state.fit_to_page = false;
        drop(state);
        self.statusbar.set_zoom(new_zoom);
        self.update_content_size();
        self.invalidate();
    }

    fn cmd_zoom_reset(&mut self) {
        {
            let mut state = self.state.lock();
            state.zoom = 1.0;
            state.fit_to_page = false;
        }
        self.statusbar.set_zoom(1.0);
        self.update_content_size();
        self.invalidate();
    }

    fn cmd_fit_to_page(&mut self) {
        let is_multipage = {
            let state = self.state.lock();
            state.multi_page_view && state.total_pages > 1
        };

        if is_multipage {
            // Multi-page documents: reset to 100% zoom instead of fit
            self.cmd_zoom_reset();
        } else {
            // Single-page: original fit-to-page behavior
            {
                let mut state = self.state.lock();
                state.fit_to_page = true;
                state.scroll_x = 0;
                state.scroll_y = 0;
            }
            self.calculate_fit_zoom();
            self.update_content_size();
            self.invalidate();
        }
    }

    fn cmd_fit_to_height(&mut self) {
        let is_multipage = {
            let state = self.state.lock();
            state.multi_page_view && state.total_pages > 1
        };

        if is_multipage {
            // Multi-page documents: reset to 100% zoom instead of fit
            self.cmd_zoom_reset();
            return;
        }

        // Single-page: original fit-to-height behavior
        let state = self.state.lock();
        if let Some(ref doc) = state.document {
            let (doc_width, doc_height) = doc.dimensions();
            let (_, h) = if state.rotation == 90 || state.rotation == 270 {
                (doc_height, doc_width)
            } else {
                (doc_width, doc_height)
            };
            drop(state);

            let (_, render_h) = self.renderer.size();
            if render_h > 0 && h > 0.0 {
                let zoom = (render_h as f32 / h).clamp(0.1, 10.0);
                {
                    let mut state = self.state.lock();
                    state.zoom = zoom;
                    state.fit_to_page = false;
                    state.scroll_x = 0;
                }
                self.statusbar.set_zoom(zoom);
                self.update_content_size();
                self.invalidate();
            }
        }
    }

    fn cmd_fit_to_width(&mut self) {
        let is_multipage = {
            let state = self.state.lock();
            state.multi_page_view && state.total_pages > 1
        };

        if is_multipage {
            // Multi-page documents: reset to 100% zoom instead of fit
            self.cmd_zoom_reset();
            return;
        }

        // Single-page: original fit-to-width behavior
        let state = self.state.lock();
        if let Some(ref doc) = state.document {
            let (doc_width, doc_height) = doc.dimensions();
            let (w, _) = if state.rotation == 90 || state.rotation == 270 {
                (doc_height, doc_width)
            } else {
                (doc_width, doc_height)
            };
            drop(state);

            let (render_w, _) = self.renderer.size();
            if render_w > 0 && w > 0.0 {
                let zoom = (render_w as f32 / w).clamp(0.1, 10.0);
                {
                    let mut state = self.state.lock();
                    state.zoom = zoom;
                    state.fit_to_page = false;
                    state.scroll_y = 0;
                }
                self.statusbar.set_zoom(zoom);
                self.update_content_size();
                self.invalidate();
            }
        }
    }

    fn calculate_fit_zoom(&mut self) {
        let state = self.state.lock();
        if let Some(ref doc) = state.document {
            let use_multipage = state.multi_page_view && state.total_pages > 1;

            if use_multipage {
                // For multi-page view, fit the widest page to viewport width
                // This allows vertical scrolling through the document
                let (render_w, _render_h) = self.renderer.size();

                // Find the widest page (accounting for rotation)
                let mut max_width: f32 = 0.0;
                for i in 0..doc.page_count() {
                    let (pw, ph) = doc.page_dimensions(i);
                    let w = if state.rotation == 90 || state.rotation == 270 { ph } else { pw };
                    max_width = max_width.max(w);
                }
                drop(state);

                if render_w > 0 && max_width > 0.0 {
                    let zoom = (render_w as f32 / max_width).clamp(0.1, 10.0);

                    let mut state = self.state.lock();
                    state.zoom = zoom;
                    drop(state);
                    self.statusbar.set_zoom(zoom);
                }
            } else {
                // Single page view - fit both width and height
                let (doc_width, doc_height) = doc.dimensions();
                let (w, h) = if state.rotation == 90 || state.rotation == 270 {
                    (doc_height, doc_width)
                } else {
                    (doc_width, doc_height)
                };
                drop(state);

                let (render_w, render_h) = self.renderer.size();
                if render_w > 0 && render_h > 0 && w > 0.0 && h > 0.0 {
                    let zoom_w = render_w as f32 / w;
                    let zoom_h = render_h as f32 / h;
                    let zoom = zoom_w.min(zoom_h).clamp(0.1, 10.0);

                    let mut state = self.state.lock();
                    state.zoom = zoom;
                    drop(state);
                    self.statusbar.set_zoom(zoom);
                }
            }
        }
    }

    /// Determine which page is most visible in the current viewport (for status bar)
    fn get_most_visible_page(&self) -> usize {
        let state = self.state.lock();

        if !state.multi_page_view || state.total_pages <= 1 {
            return state.current_page;
        }

        if let (Some(ref _doc), Some(ref layout)) = (&state.document, &state.page_layout) {
            let (_, render_h) = self.renderer.size();
            let viewport_height = render_h as i32;
            let viewport_center = state.scroll_y + viewport_height / 2;

            // Find the page whose center is closest to viewport center
            let mut best_page = 0;
            let mut best_distance = i32::MAX;

            for (i, &top) in layout.page_tops.iter().enumerate() {
                let page_h = layout.page_sizes[i].1;
                let page_center = top + page_h / 2;
                let distance = (page_center - viewport_center).abs();

                if distance < best_distance {
                    best_distance = distance;
                    best_page = i;
                }
            }

            best_page
        } else {
            state.current_page
        }
    }

    fn cmd_copy_to_clipboard(&mut self) {
        let (doc, current_page, rotation) = {
            let state = self.state.lock();
            if let Some(ref doc) = state.document {
                (doc.clone(), state.current_page, state.rotation)
            } else { return; }
        };

        if let Ok(bitmap_data) = self.wic_loader.get_bitmap_for_clipboard(&doc, current_page, rotation) {
            unsafe {
                if OpenClipboard(self.window.hwnd()).as_bool() {
                    let _ = EmptyClipboard();
                    let header_size = std::mem::size_of::<BITMAPINFOHEADER>();
                    let total_size = header_size + bitmap_data.data.len();

                    if let Ok(hglobal) = GlobalAlloc(GMEM_MOVEABLE, total_size) {
                        let ptr = GlobalLock(hglobal);
                        if !ptr.is_null() {
                            let header = BITMAPINFOHEADER {
                                biSize: header_size as u32,
                                biWidth: bitmap_data.width as i32,
                                biHeight: bitmap_data.height as i32,
                                biPlanes: 1,
                                biBitCount: 32,
                                biCompression: BI_RGB.0 as u32,
                                biSizeImage: bitmap_data.data.len() as u32,
                                ..Default::default()
                            };
                            std::ptr::copy_nonoverlapping(&header as *const _ as *const u8, ptr as *mut u8, header_size);
                            std::ptr::copy_nonoverlapping(bitmap_data.data.as_ptr(), (ptr as *mut u8).add(header_size), bitmap_data.data.len());
                            let _ = GlobalUnlock(hglobal);
                            let _ = SetClipboardData(CF_DIB.0 as u32, HANDLE(hglobal.0));
                        }
                    }
                    let _ = CloseClipboard();
                }
            }
        }
    }

    // --- File Loading Helpers ---

    fn get_file_size(path: &str) -> u64 {
        unsafe {
            let path_wide: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
            let mut file_data = WIN32_FILE_ATTRIBUTE_DATA::default();
            if GetFileAttributesExW(PCWSTR(path_wide.as_ptr()), GetFileExInfoStandard, &mut file_data as *mut _ as *mut _).as_bool() {
                ((file_data.nFileSizeHigh as u64) << 32) | (file_data.nFileSizeLow as u64)
            } else { 0 }
        }
    }

    fn scan_folder_files(path: &str) -> (Vec<String>, usize) {
        let path_obj = std::path::Path::new(path);
        let folder = match path_obj.parent() { Some(f) => f, None => return (vec![path.to_string()], 0) };
        let folder_str = match folder.to_str() { Some(s) => s, None => return (vec![path.to_string()], 0) };
        let supported_extensions = ["pdf", "jpg", "jpeg", "png", "bmp", "tif", "tiff", "webp"];
        let mut files: Vec<String> = Vec::new();

        unsafe {
            let search_pattern = format!("{}\\*", folder_str);
            let pattern_wide: Vec<u16> = search_pattern.encode_utf16().chain(std::iter::once(0)).collect();
            let mut find_data = WIN32_FIND_DATAW::default();
            let handle = match FindFirstFileW(PCWSTR(pattern_wide.as_ptr()), &mut find_data) {
                Ok(h) => h, Err(_) => return (vec![path.to_string()], 0),
            };

            loop {
                if (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY.0) == 0 {
                    let filename_len = find_data.cFileName.iter().position(|&c| c == 0).unwrap_or(find_data.cFileName.len());
                    let filename = String::from_utf16_lossy(&find_data.cFileName[..filename_len]);
                    if let Some(dot_pos) = filename.rfind('.') {
                        let ext = &filename[dot_pos + 1..].to_lowercase();
                        if supported_extensions.contains(&ext.as_str()) {
                            files.push(format!("{}\\{}", folder_str, filename));
                        }
                    }
                }
                if !FindNextFileW(handle, &mut find_data).as_bool() { break; }
            }
            let _ = FindClose(handle);
        }

        files.sort_by_key(|a| a.to_lowercase());
        let current_index = files.iter().position(|f| f.eq_ignore_ascii_case(path)).unwrap_or(0);
        (files, current_index)
    }

    fn open_document(&mut self, path: &str) {
        let skip_scan = self.opened_from_cmdline;
        self.open_document_internal(path, false, skip_scan);
        self.opened_from_cmdline = false;
    }

    fn open_document_with_mode(&mut self, path: &str, keep_folder_mode: bool) {
        self.open_document_internal(path, keep_folder_mode, false);
    }

    fn open_document_internal(&mut self, path: &str, keep_folder_mode: bool, skip_folder_scan: bool) {
        let _wait_cursor = WaitCursorGuard::new();

        // Show filename in statusbar immediately before loading
        let filename = std::path::Path::new(path).file_name().and_then(|n| n.to_str()).unwrap_or("Datei");
        self.statusbar.set_loading_file(filename);
        // Force immediate repaint of statusbar
        unsafe {
            UpdateWindow(self.window.hwnd());
        }

        let ext = std::path::Path::new(path).extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
        let result = match ext.as_str() {
            "pdf" => self.load_pdf(path),
            "jpg" | "jpeg" | "png" | "bmp" | "tif" | "tiff" | "webp" | "ico" | "icon" => self.load_image(path),
            _ => Err(Error::from_win32()),
        };

        match result {
            Ok(doc) => {
                let total_pages = doc.page_count();
                let (width, height) = doc.dimensions();
                let file_size = Self::get_file_size(path);
                
                let (folder_files, folder_index) = if skip_folder_scan {
                    (vec![path.to_string()], 0)
                } else {
                    Self::scan_folder_files(path)
                };

                let nav_mode = if keep_folder_mode { true } else if skip_folder_scan { false } else { total_pages == 1 };

                // Multi-page documents use 100% zoom, single-page uses fit-to-page
                let is_multipage = total_pages > 1;
                let initial_zoom = 1.0; // Will be recalculated for single-page

                {
                    let mut state = self.state.lock();
                    state.document = Some(doc.clone());
                    state.current_page = 0;
                    state.total_pages = total_pages;
                    state.rotation = 0;
                    state.file_path = Some(path.to_string());
                    state.fit_to_page = !is_multipage; // Fit only for single-page documents
                    state.zoom = initial_zoom;
                    state.folder_files = folder_files;
                    state.folder_file_index = folder_index;
                    state.folder_navigation_mode = nav_mode;
                    state.scroll_x = 0;
                    state.scroll_y = 0;
                }

                let filename = std::path::Path::new(path).file_name().and_then(|n| n.to_str()).unwrap_or("SimpliView");
                self.window.set_title("SimpliView");

                let dim_str = if doc.doc_type() == crate::document::DocumentType::Pdf {
                   format!("{:.0}x{:.0} mm", width * 25.4 / 72.0, height * 25.4 / 72.0)
                } else {
                   format!("{}x{} px", width as u32, height as u32)
                };

                self.statusbar.set_file_info(filename, &dim_str, file_size, 0, total_pages);
                self.top_toolbar.set_document_loaded(true);
                self.top_toolbar.set_navigation_enabled(total_pages > 1);
                self.statusbar.set_document_loaded(true);
                self.context_menu.set_document_loaded(true);

                // Initial layout calculation - only fit for single-page documents
                if !is_multipage {
                    self.calculate_fit_zoom();
                } else {
                    self.statusbar.set_zoom(1.0);
                }
                self.update_content_size();
                self.invalidate();
            }
            Err(e) => {
                // Don't show error for user cancellation (e.g., cancelled password dialog)
                const ERROR_CANCELLED: u32 = 0x800704C7;
                if e.code().0 as u32 != ERROR_CANCELLED {
                    self.show_error(&format!("Datei konnte nicht geöffnet werden: {:?}", e));
                }
            }
        }
    }

    /// Loads a PDF document, handling password-protected files with user prompts.
    ///
    /// Flow:
    /// 1. Try loading without password
    /// 2. If password required, prompt user (up to MAX_PASSWORD_ATTEMPTS times)
    /// 3. User can cancel at any time to abort loading gracefully
    ///
    /// Returns ERROR_CANCELLED (0x800704C7) when user cancels to distinguish from real errors.
    fn load_pdf(&mut self, path: &str) -> Result<Document> {
        const MAX_PASSWORD_ATTEMPTS: u32 = 3;
        // ERROR_CANCELLED - used to signal user cancellation (no error message should be shown)
        const ERROR_CANCELLED: i32 = 0x800704C7u32 as i32;

        // First attempt: try without password
        match self.pdf_loader.load(path, None) {
            Ok(doc) => return Ok(doc),
            Err(e) => {
                // Check if this is a password-protected PDF
                let is_password_err = self.pdf_loader.needs_password()
                    || (e.code().0 as u32 == 0x8007052B);  // ERROR_WRONG_PASSWORD

                if !is_password_err {
                    // Not a password error - propagate the original error
                    return Err(e);
                }
            }
        }

        // PDF requires password - prompt user with retry limit
        let mut attempts = 0u32;
        loop {
            // Show password dialog
            let password = match self.prompt_password() {
                Some(pwd) => pwd,
                None => {
                    // User cancelled - return special cancellation error (handled silently)
                    return Err(Error::from(windows::core::HRESULT(ERROR_CANCELLED)));
                }
            };

            attempts += 1;

            // Try loading with provided password
            match self.pdf_loader.load(path, Some(&password)) {
                Ok(doc) => return Ok(doc),  // Success!
                Err(_) => {
                    // Wrong password
                    if attempts >= MAX_PASSWORD_ATTEMPTS {
                        // Max attempts reached - show final error and return cancellation
                        // (error already shown, so use cancellation code to prevent duplicate message)
                        self.show_error("Maximale Anzahl an Kennwort-Versuchen überschritten. Dokument kann nicht geöffnet werden.");
                        return Err(Error::from(windows::core::HRESULT(ERROR_CANCELLED)));
                    }

                    // Ask if user wants to retry
                    if !self.retry_password() {
                        // User chose not to retry - return cancellation
                        return Err(Error::from(windows::core::HRESULT(ERROR_CANCELLED)));
                    }
                    // Loop continues for another attempt
                }
            }
        }
    }

    fn load_image(&mut self, path: &str) -> Result<Document> {
        self.wic_loader.load(path)
    }

    fn export_document(&self, path: &str) {
        let (doc, current_page, source_path) = {
            let state = self.state.lock();
            if let Some(ref doc) = state.document {
                (Some(doc.clone()), state.current_page, state.file_path.clone())
            } else { (None, 0, None) }
        };

        if let Some(doc) = doc {
            // Check if user chose PDF export
            if path.to_lowercase().ends_with(".pdf") {
                if doc.doc_type() == crate::document::DocumentType::Pdf {
                    if let Some(src) = source_path {
                        if let Err(e) = std::fs::copy(&src, path) {
                            self.show_error(&format!("PDF-Export fehlgeschlagen: {}", e));
                        }
                        return;
                    }
                } else {
                    self.show_error("Bild kann nicht als PDF exportiert werden. Bitte wählen Sie ein Bildformat.");
                    return;
                }
            }

            if let Err(e) = self.wic_loader.save(&doc, path, current_page) {
                self.show_error(&format!("Export fehlgeschlagen: {:?}", e));
            }
        }
    }

    fn prompt_password(&self) -> Option<String> { crate::dialogs::password_dialog(self.window.hwnd()) }
    fn retry_password(&self) -> bool { crate::dialogs::retry_password_dialog(self.window.hwnd()) }
    fn show_error(&self, message: &str) { crate::dialogs::show_error(self.window.hwnd(), message); }
    
    fn update_page_display(&mut self, page: usize, total: usize, path: Option<&str>) {
        let state = self.state.lock();
        if let Some(ref doc) = state.document {
            let (width, height) = doc.page_dimensions(page);
            let file_size = path.map(Self::get_file_size).unwrap_or(0);
            let filename = path.and_then(|p| std::path::Path::new(p).file_name()).and_then(|n| n.to_str()).unwrap_or("");
            let dim_str = if doc.doc_type() == crate::document::DocumentType::Pdf {
                format!("{:.0}x{:.0} mm", width * 25.4 / 72.0, height * 25.4 / 72.0)
            } else {
                format!("{}x{} px", width as u32, height as u32)
            };
            drop(state);
            self.statusbar.set_file_info(filename, &dim_str, file_size, page, total);
        }
    }
}
