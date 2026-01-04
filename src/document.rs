use parking_lot::Mutex;
use std::collections::HashMap;
use windows::{
    core::*,
    Win32::Graphics::{
        Direct2D::{Common::*, *},
        Imaging::*,
    },
};

/// Vertical gap between pages in pixels (at zoom 1.0)
pub const PAGE_GAP: i32 = 20;

/// Maximum number of page bitmaps to keep in cache
pub const MAX_CACHED_PAGES: usize = 20;

/// Pre-computed layout information for multi-page rendering
#[derive(Clone, Debug)]
pub struct PageLayout {
    /// Y-position of each page's top edge (in scaled pixels)
    pub page_tops: Vec<i32>,
    /// Total document height including all pages and gaps
    pub total_height: i32,
    /// Width of the widest page (for horizontal centering)
    pub max_width: i32,
    /// Individual page dimensions (width, height) after rotation, scaled
    pub page_sizes: Vec<(i32, i32)>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum DocumentType {
    Image,
    Pdf,
}

pub struct Document {
    doc_type: DocumentType,
    pages: Vec<PageData>,
    bitmap_cache: Mutex<HashMap<usize, ID2D1Bitmap>>,
}

pub struct PageData {
    pub width: f32,
    pub height: f32,
    pub wic_bitmap: Option<IWICBitmap>,
    pub pixel_data: Option<Vec<u8>>,
    pub stride: u32,
}

impl Document {
    pub fn new_image(wic_bitmap: IWICBitmap, width: u32, height: u32) -> Self {
        Self {
            doc_type: DocumentType::Image,
            pages: vec![PageData {
                width: width as f32,
                height: height as f32,
                wic_bitmap: Some(wic_bitmap),
                pixel_data: None,
                stride: 0,
            }],
            bitmap_cache: Mutex::new(HashMap::new()),
        }
    }

    pub fn new_pdf(pages: Vec<PageData>) -> Self {
        Self {
            doc_type: DocumentType::Pdf,
            pages,
            bitmap_cache: Mutex::new(HashMap::new()),
        }
    }

    pub fn doc_type(&self) -> DocumentType {
        self.doc_type
    }

    pub fn page_count(&self) -> usize {
        self.pages.len()
    }

    pub fn dimensions(&self) -> (f32, f32) {
        if let Some(page) = self.pages.first() {
            (page.width, page.height)
        } else {
            (0.0, 0.0)
        }
    }

    pub fn page_dimensions(&self, page: usize) -> (f32, f32) {
        if let Some(p) = self.pages.get(page) {
            (p.width, p.height)
        } else {
            (0.0, 0.0)
        }
    }

    pub fn get_page_bitmap(&self, rt: &ID2D1HwndRenderTarget, page: usize) -> Result<ID2D1Bitmap> {
        // Check cache first
        {
            let cache = self.bitmap_cache.lock();
            if let Some(bitmap) = cache.get(&page) {
                return Ok(bitmap.clone());
            }
        }

        // Create bitmap from page data
        let page_data = self.pages.get(page).ok_or_else(Error::from_win32)?;

        let bitmap = if let Some(ref wic_bitmap) = page_data.wic_bitmap {
            // Create D2D bitmap from WIC bitmap
            unsafe {
                let props = D2D1_BITMAP_PROPERTIES {
                    pixelFormat: D2D1_PIXEL_FORMAT {
                        format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
                        alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
                    },
                    dpiX: 96.0,
                    dpiY: 96.0,
                };

                rt.CreateBitmapFromWicBitmap(wic_bitmap, Some(&props))?
            }
        } else if let Some(ref pixel_data) = page_data.pixel_data {
            // Create D2D bitmap from raw pixel data
            unsafe {
                let size = windows::Win32::Graphics::Direct2D::Common::D2D_SIZE_U {
                    width: page_data.width as u32,
                    height: page_data.height as u32,
                };

                let props = D2D1_BITMAP_PROPERTIES {
                    pixelFormat: D2D1_PIXEL_FORMAT {
                        format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
                        alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
                    },
                    dpiX: 96.0,
                    dpiY: 96.0,
                };

                rt.CreateBitmap(size, Some(pixel_data.as_ptr() as *const _), page_data.stride, &props)?
            }
        } else {
            return Err(Error::from_win32());
        };

        // Cache the bitmap
        self.bitmap_cache.lock().insert(page, bitmap.clone());

        Ok(bitmap)
    }

    #[allow(dead_code)]
    pub fn clear_cache(&self) {
        self.bitmap_cache.lock().clear();
    }

    pub fn get_wic_bitmap(&self, page: usize) -> Option<&IWICBitmap> {
        self.pages.get(page).and_then(|p| p.wic_bitmap.as_ref())
    }

    pub fn get_pixel_data(&self, page: usize) -> Option<(&[u8], u32, u32, u32)> {
        self.pages.get(page).and_then(|p| {
            p.pixel_data.as_ref().map(|data| {
                (data.as_slice(), p.width as u32, p.height as u32, p.stride)
            })
        })
    }

    /// Compute layout for multi-page vertical stacking
    ///
    /// Returns pre-computed Y positions for each page top, total height,
    /// and maximum width for horizontal centering.
    pub fn compute_layout(&self, zoom: f32, rotation: i32) -> PageLayout {
        let mut page_tops = Vec::with_capacity(self.pages.len());
        let mut page_sizes = Vec::with_capacity(self.pages.len());
        let mut current_y: i32 = 0;
        let mut max_width: i32 = 0;
        let scaled_gap = (PAGE_GAP as f32 * zoom) as i32;

        for (i, page) in self.pages.iter().enumerate() {
            // Determine dimensions based on rotation
            let (w, h) = if rotation == 90 || rotation == 270 {
                (page.height, page.width)
            } else {
                (page.width, page.height)
            };

            // Scale by zoom
            let scaled_w = (w * zoom) as i32;
            let scaled_h = (h * zoom) as i32;

            page_tops.push(current_y);
            page_sizes.push((scaled_w, scaled_h));

            max_width = max_width.max(scaled_w);
            current_y += scaled_h;

            // Add gap after each page except the last
            if i < self.pages.len() - 1 {
                current_y += scaled_gap;
            }
        }

        PageLayout {
            page_tops,
            total_height: current_y,
            max_width,
            page_sizes,
        }
    }

    /// Find which pages are visible in the current viewport
    ///
    /// Returns (first_visible_page, last_visible_page_exclusive)
    pub fn find_visible_pages(
        &self,
        layout: &PageLayout,
        scroll_y: i32,
        viewport_height: i32,
    ) -> (usize, usize) {
        if self.pages.is_empty() {
            return (0, 0);
        }

        let viewport_bottom = scroll_y + viewport_height;

        // Find first page overlapping viewport
        let mut first_visible = 0;
        for (i, &top) in layout.page_tops.iter().enumerate() {
            let page_bottom = top + layout.page_sizes[i].1;
            if page_bottom > scroll_y {
                first_visible = i;
                break;
            }
            first_visible = i; // Keep updating in case we reach the end
        }

        // Linear scan forward for last visible page
        let mut last_visible = first_visible;
        for i in first_visible..self.pages.len() {
            if layout.page_tops[i] >= viewport_bottom {
                break;
            }
            last_visible = i + 1;
        }

        (first_visible, last_visible.min(self.pages.len()))
    }

    /// Evict distant pages from cache to limit memory usage
    pub fn evict_distant_pages(&self, center_page: usize) {
        let mut cache = self.bitmap_cache.lock();
        if cache.len() <= MAX_CACHED_PAGES {
            return;
        }

        // Collect page indices sorted by distance from center
        let mut pages: Vec<usize> = cache.keys().copied().collect();
        pages.sort_by_key(|&p| (p as i32 - center_page as i32).abs());

        // Remove pages beyond the limit
        for &page in pages.iter().skip(MAX_CACHED_PAGES) {
            cache.remove(&page);
        }
    }
}

impl Clone for Document {
    fn clone(&self) -> Self {
        // Note: WIC bitmaps and cache are not cloned
        Self {
            doc_type: self.doc_type,
            pages: self.pages.iter().map(|p| PageData {
                width: p.width,
                height: p.height,
                wic_bitmap: p.wic_bitmap.clone(),
                pixel_data: p.pixel_data.clone(),
                stride: p.stride,
            }).collect(),
            bitmap_cache: Mutex::new(HashMap::new()),
        }
    }
}
