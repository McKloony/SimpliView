use crate::document::{Document, PageLayout};
use std::cell::RefCell;
use windows::{
    core::*,
    Foundation::Numerics::Matrix3x2,
    Win32::{
        Foundation::*,
        Graphics::{
            Direct2D::{Common::*, *},
            Dxgi::Common::*,
        },
        UI::WindowsAndMessaging::GetClientRect,
    },
};

thread_local! {
    static D2D_FACTORY: RefCell<Option<ID2D1Factory1>> = const { RefCell::new(None) };
}

fn get_d2d_factory() -> Result<ID2D1Factory1> {
    D2D_FACTORY.with(|cell| {
        let mut opt = cell.borrow_mut();
        if opt.is_none() {
            let factory: ID2D1Factory1 = unsafe {
                D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)?
            };
            *opt = Some(factory);
        }
        Ok(opt.as_ref().unwrap().clone())
    })
}

pub struct D2DRenderer {
    hwnd: HWND,
    render_target: Option<ID2D1HwndRenderTarget>,
    width: u32,
    height: u32,
}

impl D2DRenderer {
    pub fn new(hwnd: HWND) -> Result<Self> {
        let mut renderer = Self {
            hwnd,
            render_target: None,
            width: 0,
            height: 0,
        };
        renderer.create_render_target()?;
        Ok(renderer)
    }

    fn create_render_target(&mut self) -> Result<()> {
        unsafe {
            let factory = get_d2d_factory()?;

            // Get client rect
            let mut rect = RECT::default();
            let _ = GetClientRect(self.hwnd, &mut rect);

            self.width = (rect.right - rect.left) as u32;
            self.height = (rect.bottom - rect.top) as u32;

            let size = D2D_SIZE_U {
                width: self.width.max(1),
                height: self.height.max(1),
            };

            let props = D2D1_RENDER_TARGET_PROPERTIES {
                r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
                pixelFormat: D2D1_PIXEL_FORMAT {
                    format: DXGI_FORMAT_B8G8R8A8_UNORM,
                    alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
                },
                dpiX: 0.0,
                dpiY: 0.0,
                usage: D2D1_RENDER_TARGET_USAGE_NONE,
                minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
            };

            let hwnd_props = D2D1_HWND_RENDER_TARGET_PROPERTIES {
                hwnd: self.hwnd,
                pixelSize: size,
                presentOptions: D2D1_PRESENT_OPTIONS_NONE,
            };

            self.render_target = Some(factory.CreateHwndRenderTarget(&props, &hwnd_props)?);
        }
        Ok(())
    }

    pub fn resize(&mut self, width: u32, height: u32) -> Result<()> {
        self.width = width;
        self.height = height;

        if let Some(ref rt) = self.render_target {
            unsafe {
                let size = D2D_SIZE_U { width, height };
                rt.Resize(&size)?;
            }
        }
        Ok(())
    }

    pub fn size(&self) -> (u32, u32) {
        (self.width, self.height)
    }

    pub fn begin_draw(&mut self) -> Result<()> {
        if self.render_target.is_none() {
            self.create_render_target()?;
        }

        if let Some(ref rt) = self.render_target {
            unsafe {
                rt.BeginDraw();
                // No need for clip rect as we render to the full child window
            }
        }
        Ok(())
    }

    pub fn end_draw(&mut self) -> Result<()> {
        if let Some(ref rt) = self.render_target {
            unsafe {
                match rt.EndDraw(None, None) {
                    Ok(_) => Ok(()),
                    Err(e) if e.code() == D2DERR_RECREATE_TARGET => {
                        // Need to recreate render target
                        self.render_target = None;
                        self.create_render_target()
                    }
                    Err(e) => Err(e),
                }
            }
        } else {
            Ok(())
        }
    }

    pub fn clear(&self, color: D2D1_COLOR_F) {
        if let Some(ref rt) = self.render_target {
            unsafe {
                rt.Clear(Some(&color));
            }
        }
    }

    pub fn draw_document(
        &self,
        doc: &Document,
        zoom: f32,
        rotation: i32,
        page: usize,
        scroll_x: i32,
        scroll_y: i32,
    ) -> Result<()> {
        let rt = match &self.render_target {
            Some(rt) => rt,
            None => return Ok(()),
        };

        // Get the bitmap for the current page
        let bitmap = doc.get_page_bitmap(rt, page)?;

        unsafe {
            let bitmap_size = bitmap.GetSize();
            let unrotated_w = bitmap_size.width * zoom;
            let unrotated_h = bitmap_size.height * zoom;

            // Determine dimensions of the bounding box after rotation
            let (layout_w, layout_h) = if rotation == 90 || rotation == 270 {
                (unrotated_h, unrotated_w)
            } else {
                (unrotated_w, unrotated_h)
            };

            let viewport_width = self.width as f32;
            let viewport_height = self.height as f32;

            // Calculate top-left of the bounding box relative to viewport
            // If content fits, center it. Otherwise, use negative scroll offset.
            let bbox_left = if layout_w <= viewport_width {
                (viewport_width - layout_w) / 2.0
            } else {
                -(scroll_x as f32)
            };

            let bbox_top = if layout_h <= viewport_height {
                (viewport_height - layout_h) / 2.0
            } else {
                -(scroll_y as f32)
            };

            // The center of rotation is the center of the visual bounding box
            let center_x = bbox_left + layout_w / 2.0;
            let center_y = bbox_top + layout_h / 2.0;

            // Calculate rotation transform around this center
            let angle = rotation as f32;
            let rotation_transform = make_rotation_matrix(angle, center_x, center_y);

            // The destination rectangle is the unrotated image centered at the same point
            let dest_rect = D2D_RECT_F {
                left: center_x - unrotated_w / 2.0,
                top: center_y - unrotated_h / 2.0,
                right: center_x + unrotated_w / 2.0,
                bottom: center_y + unrotated_h / 2.0,
            };

            // Apply transform
            rt.SetTransform(&rotation_transform);

            // Draw bitmap
            rt.DrawBitmap(
                &bitmap,
                Some(&dest_rect),
                1.0,
                D2D1_BITMAP_INTERPOLATION_MODE_LINEAR,
                None,
            );

            // Reset transform
            rt.SetTransform(&make_identity_matrix());
        }

        Ok(())
    }

    #[allow(dead_code)]
    pub fn render_target(&self) -> Option<&ID2D1HwndRenderTarget> {
        self.render_target.as_ref()
    }

    /// Draw all visible pages of a multi-page document
    ///
    /// Pages are stacked vertically with gaps between them.
    /// Only pages intersecting the viewport are rendered.
    pub fn draw_document_multipage(
        &self,
        doc: &Document,
        layout: &PageLayout,
        zoom: f32,
        rotation: i32,
        scroll_x: i32,
        scroll_y: i32,
    ) -> Result<()> {
        let rt = match &self.render_target {
            Some(rt) => rt,
            None => return Ok(()),
        };

        let viewport_width = self.width as f32;
        let viewport_height = self.height as i32;

        // Find which pages are visible
        let (first_page, last_page) = doc.find_visible_pages(layout, scroll_y, viewport_height);

        // Draw each visible page
        for page_idx in first_page..last_page {
            let bitmap = doc.get_page_bitmap(rt, page_idx)?;

            unsafe {
                let bitmap_size = bitmap.GetSize();
                let unrotated_w = bitmap_size.width * zoom;
                let unrotated_h = bitmap_size.height * zoom;

                // Get page position from layout
                let page_top = layout.page_tops[page_idx];
                let (page_w, page_h) = layout.page_sizes[page_idx];

                // Calculate Y position relative to viewport
                let draw_y = (page_top - scroll_y) as f32;

                // Calculate X position - center page horizontally
                // If content fits in viewport, center the entire document
                // If content is wider, use scroll offset
                let draw_x = if layout.max_width <= viewport_width as i32 {
                    // Center the widest page, then center this page relative to that
                    let doc_center_x = viewport_width / 2.0;
                    doc_center_x - (page_w as f32 / 2.0)
                } else {
                    // Scrolling horizontally - all pages align left at -scroll_x
                    // Then offset by half the difference between max width and this page's width
                    let page_offset = (layout.max_width - page_w) / 2;
                    -(scroll_x as f32) + page_offset as f32
                };

                // The center of rotation is the center of the page's visual bounding box
                let center_x = draw_x + page_w as f32 / 2.0;
                let center_y = draw_y + page_h as f32 / 2.0;

                // Calculate rotation transform around this center
                let angle = rotation as f32;
                let rotation_transform = make_rotation_matrix(angle, center_x, center_y);

                // The destination rectangle is the unrotated image centered at the same point
                let dest_rect = D2D_RECT_F {
                    left: center_x - unrotated_w / 2.0,
                    top: center_y - unrotated_h / 2.0,
                    right: center_x + unrotated_w / 2.0,
                    bottom: center_y + unrotated_h / 2.0,
                };

                // Apply transform
                rt.SetTransform(&rotation_transform);

                // Draw bitmap
                rt.DrawBitmap(
                    &bitmap,
                    Some(&dest_rect),
                    1.0,
                    D2D1_BITMAP_INTERPOLATION_MODE_LINEAR,
                    None,
                );

                // Reset transform for next page
                rt.SetTransform(&make_identity_matrix());
            }
        }

        // Evict distant pages from cache to limit memory
        if first_page < last_page {
            let center_page = (first_page + last_page) / 2;
            doc.evict_distant_pages(center_page);
        }

        Ok(())
    }

    /// Draw a 1px horizontal separator line at the bottom of the viewport
    pub fn draw_bottom_separator(&self, color: D2D1_COLOR_F) {
        if let Some(ref rt) = self.render_target {
            unsafe {
                if let Ok(brush) = rt.CreateSolidColorBrush(&color, None) {
                    let y = self.height as f32 - 0.5; // Center of the bottom pixel row
                    rt.DrawLine(
                        D2D_POINT_2F { x: 0.0, y },
                        D2D_POINT_2F { x: self.width as f32, y },
                        &brush,
                        1.0,
                        None,
                    );
                }
            }
        }
    }
}

// Matrix helper functions
fn make_identity_matrix() -> Matrix3x2 {
    Matrix3x2 {
        M11: 1.0,
        M12: 0.0,
        M21: 0.0,
        M22: 1.0,
        M31: 0.0,
        M32: 0.0,
    }
}

fn make_rotation_matrix(angle_degrees: f32, center_x: f32, center_y: f32) -> Matrix3x2 {
    let angle_radians = angle_degrees * std::f32::consts::PI / 180.0;
    let cos = angle_radians.cos();
    let sin = angle_radians.sin();

    Matrix3x2 {
        M11: cos,
        M12: sin,
        M21: -sin,
        M22: cos,
        M31: center_x - center_x * cos + center_y * sin,
        M32: center_y - center_x * sin - center_y * cos,
    }
}