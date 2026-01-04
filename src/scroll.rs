//! Native Windows scrollbar management module
//!
//! Provides a wrapper around Win32 scrollbar APIs for managing native scrollbars
//! on windows with WS_HSCROLL and WS_VSCROLL styles.

use windows::Win32::{
    Foundation::*,
    UI::{
        Controls::{SetScrollInfo, ShowScrollBar},
        WindowsAndMessaging::{
            GetScrollInfo, SCROLLINFO, SIF_ALL, SIF_POS, SIF_TRACKPOS,
            SCROLLBAR_CONSTANTS, SB_BOTH, SB_HORZ, SB_VERT,
            SB_BOTTOM, SB_ENDSCROLL, SB_LEFT, SB_LINEDOWN, SB_LINELEFT,
            SB_LINERIGHT, SB_LINEUP, SB_PAGEDOWN, SB_PAGELEFT, SB_PAGERIGHT, SB_PAGEUP,
            SB_RIGHT, SB_THUMBPOSITION, SB_THUMBTRACK, SB_TOP,
        },
    },
};

/// Pixels to scroll per "line" (arrow click or single wheel notch line)
pub const LINE_SCROLL_PIXELS: i32 = 40;

/// Scroll action from WM_HSCROLL/WM_VSCROLL
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ScrollAction {
    LineUp,       // SB_LINEUP / SB_LINELEFT
    LineDown,     // SB_LINEDOWN / SB_LINERIGHT
    PageUp,       // SB_PAGEUP / SB_PAGELEFT
    PageDown,     // SB_PAGEDOWN / SB_PAGERIGHT
    ThumbTrack,   // SB_THUMBTRACK (dragging)
    ThumbPosition,// SB_THUMBPOSITION (drag released)
    Top,          // SB_TOP / SB_LEFT
    Bottom,       // SB_BOTTOM / SB_RIGHT
    EndScroll,    // SB_ENDSCROLL
}

impl ScrollAction {
    /// Parse scroll action from WPARAM's low word
    pub fn from_scroll_code(code: u16) -> Option<Self> {
        let code = code as i32;
        match code {
            x if x == SB_LINEUP.0 || x == SB_LINELEFT.0 => Some(ScrollAction::LineUp),
            x if x == SB_LINEDOWN.0 || x == SB_LINERIGHT.0 => Some(ScrollAction::LineDown),
            x if x == SB_PAGEUP.0 || x == SB_PAGELEFT.0 => Some(ScrollAction::PageUp),
            x if x == SB_PAGEDOWN.0 || x == SB_PAGERIGHT.0 => Some(ScrollAction::PageDown),
            x if x == SB_THUMBTRACK.0 => Some(ScrollAction::ThumbTrack),
            x if x == SB_THUMBPOSITION.0 => Some(ScrollAction::ThumbPosition),
            x if x == SB_TOP.0 || x == SB_LEFT.0 => Some(ScrollAction::Top),
            x if x == SB_BOTTOM.0 || x == SB_RIGHT.0 => Some(ScrollAction::Bottom),
            x if x == SB_ENDSCROLL.0 => Some(ScrollAction::EndScroll),
            _ => None,
        }
    }
}

/// Manages scrollbars for a window
pub struct ScrollManager {
    hwnd: HWND,
}

impl ScrollManager {
    pub fn new(hwnd: HWND) -> Self {
        Self { hwnd }
    }

    /// Update scrollbar range and position.
    /// Call this when viewport size, content size, or scroll position changes.
    pub fn update_scrollbar(
        &self,
        bar: SCROLLBAR_CONSTANTS,
        viewport_size: i32,
        content_size: i32,
        scroll_pos: i32,
    ) {
        // Only show scrollbar if content is strictly larger than viewport
        let needs_scroll = content_size > viewport_size;

        if needs_scroll {
            let si = SCROLLINFO {
                cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
                fMask: SIF_ALL,
                nMin: 0,
                // nMax is the maximum reachable value for the scrolling range.
                // The scrollbar thumb size is determined by nPage.
                // The maximum scroll position is (nMax - nPage + 1).
                // So if we want max pos to be (content - viewport), we need:
                // max_pos = nMax - viewport + 1 => content - viewport
                // nMax = content - 1
                nMax: content_size - 1, 
                nPage: viewport_size as u32,
                nPos: scroll_pos,
                nTrackPos: 0,
            };

            unsafe {
                let _ = SetScrollInfo(self.hwnd, bar, &si, true);
                let _ = ShowScrollBar(self.hwnd, bar, true);
            }
        } else {
            // Hide scrollbar when content fits
            unsafe {
                let _ = ShowScrollBar(self.hwnd, bar, false);
            }
        }
    }

    /// Update both scrollbars at once
    pub fn update_both(
        &self,
        viewport_width: i32,
        viewport_height: i32,
        content_width: i32,
        content_height: i32,
        scroll_x: i32,
        scroll_y: i32,
    ) {
        self.update_scrollbar(SB_HORZ, viewport_width, content_width, scroll_x);
        self.update_scrollbar(SB_VERT, viewport_height, content_height, scroll_y);
    }

    /// Hide both scrollbars (used when fit-to-page is active)
    pub fn hide_both(&self) {
        unsafe {
            let _ = ShowScrollBar(self.hwnd, SB_BOTH, false);
        }
    }

    /// Get the current track position (used during SB_THUMBTRACK for smooth dragging)
    pub fn get_track_pos(&self, bar: SCROLLBAR_CONSTANTS) -> i32 {
        let mut si = SCROLLINFO {
            cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
            fMask: SIF_TRACKPOS,
            ..Default::default()
        };
        unsafe {
            let _ = GetScrollInfo(self.hwnd, bar, &mut si);
        }
        si.nTrackPos
    }

    /// Set just the scroll position (without changing range)
    pub fn set_pos(&self, bar: SCROLLBAR_CONSTANTS, pos: i32) {
        let si = SCROLLINFO {
            cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
            fMask: SIF_POS,
            nPos: pos,
            ..Default::default()
        };
        unsafe {
            let _ = SetScrollInfo(self.hwnd, bar, &si, true);
        }
    }

    /// Calculate new scroll position based on scroll action
    pub fn calculate_new_pos(
        action: ScrollAction,
        current_pos: i32,
        page_size: i32,
        content_size: i32,
        viewport_size: i32,
        track_pos: i32,
    ) -> i32 {
        let max_scroll = (content_size - viewport_size).max(0);

        match action {
            ScrollAction::LineUp => (current_pos - LINE_SCROLL_PIXELS).max(0),
            ScrollAction::LineDown => (current_pos + LINE_SCROLL_PIXELS).min(max_scroll),
            ScrollAction::PageUp => (current_pos - page_size).max(0),
            ScrollAction::PageDown => (current_pos + page_size).min(max_scroll),
            ScrollAction::Top => 0,
            ScrollAction::Bottom => max_scroll,
            ScrollAction::ThumbTrack | ScrollAction::ThumbPosition => track_pos.clamp(0, max_scroll),
            ScrollAction::EndScroll => current_pos, // No change
        }
    }
}

/// Clamp scroll position to valid range
pub fn clamp_scroll(scroll: i32, viewport_size: i32, content_size: i32) -> i32 {
    let max_scroll = (content_size - viewport_size).max(0);
    scroll.clamp(0, max_scroll)
}