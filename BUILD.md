# CrossView - Build Instructions

## Overview

CrossView is a high-performance, minimalist PDF/Image viewer for Windows 11 and Windows Server 2022. It's built entirely in Rust using native Win32 APIs and Direct2D for rendering.

## Requirements

- **Rust**: 1.70 or later (stable toolchain)
- **Target**: `x86_64-pc-windows-msvc`
- **Visual Studio Build Tools 2019** or later with:
  - MSVC v142 or later (C++ build tools)
  - Windows 10/11 SDK

## Build Instructions

### Quick Build (Debug)

```bash
cargo build
```

### Release Build (Optimized)

```bash
cargo build --release
```

The output will be at `target/release/CrossView.exe`.

### Static CRT Build (Recommended for Distribution)

For a fully static executable with no external CRT dependencies:

```bash
set RUSTFLAGS=-C target-feature=+crt-static
cargo build --release
```

Or add to `.cargo/config.toml`:

```toml
[target.x86_64-pc-windows-msvc]
rustflags = ["-C", "target-feature=+crt-static"]
```

## Running CrossView

### Basic Usage

```bash
CrossView.exe
```

### Open a File on Startup

```bash
CrossView.exe "C:\path\to\document.pdf"
```

### With Restricted Path (Sandboxed File Dialogs)

```bash
CrossView.exe "C:\path\to\document.pdf" "C:\allowed\folder"
```

When a restricted path is provided:
- Open/Export dialogs default to and prefer this folder
- Note: Complete sandboxing is limited by Windows dialog behavior; users may still navigate away but the default location is enforced

## Command-Line Arguments

| Argument | Description |
|----------|-------------|
| `[file]` | Optional: File to open on startup |
| `[path]` | Optional: Restricted base path for file dialogs |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open file |
| `Ctrl+S` | Export current page |
| `Ctrl++` | Zoom in |
| `Ctrl+-` | Zoom out |
| `Ctrl+0` | Fit to page |
| `Left/PageUp` | Previous page |
| `Right/PageDown` | Next page |
| `Home` | First page |
| `End` | Last page |
| `Ctrl+Wheel` | Zoom in/out |
| `Right-click` | Context menu |

## Supported File Types

- **PDF**: PDF documents (including password-protected)
- **Images**: JPG, JPEG, PNG, BMP, TIF, TIFF, WEBP

## Features

- **Fast startup**: Minimal dependencies, native Win32 APIs
- **Multiple instances**: No single-instance mutex
- **Per-monitor DPI**: Full DPI awareness (PerMonitorV2)
- **Multi-monitor support**: Opens on monitor containing cursor
- **Dark/Light theme**: Auto-detects Windows theme, manual toggle available
- **PDF password support**: Prompts for password when needed
- **Zoom**: Fit-to-page, zoom in/out, mouse wheel zoom
- **Rotation**: 90-degree rotation in either direction
- **Page navigation**: Full PDF page navigation
- **Export**: Export current page to PNG, JPG, BMP, or TIFF

## File Association Registration

CrossView can register itself as a handler for supported file types:

```rust
// In your code:
use crossview::registration;

// Register (per-user, no admin required)
registration::register_file_associations()?;

// Unregister
registration::unregister_file_associations()?;
```

### Important Notes on Windows File Associations

1. **Cannot Force Default App**: Windows 10/11 does not allow applications to programmatically set themselves as the default handler. Users must manually set defaults.

2. **User Steps to Set Default**:
   - Right-click a file → "Open with" → "Choose another app"
   - Select "CrossView" and check "Always use this app"
   - Or: Settings → Apps → Default apps → Choose defaults by file type

3. **What Registration Does**:
   - Registers CrossView as a capable handler
   - Adds entries to "Open with" context menus
   - Creates ProgID in per-user registry (no admin required)

## Architecture

```
CrossView/
├── Cargo.toml          # Project configuration
├── build.rs            # Resource compilation
├── BUILD.md            # This file
├── photo_portrait.ico  # Application icon
├── skn16g.png         # Toolbar icon strip
├── resources/
│   ├── crossview.rc    # Windows resources
│   └── crossview.manifest  # Application manifest
└── src/
    ├── main.rs         # Entry point
    ├── app.rs          # Application logic
    ├── window.rs       # Win32 window
    ├── toolbar.rs      # Toolbar (Common Controls)
    ├── statusbar.rs    # Status bar
    ├── d2d.rs          # Direct2D rendering
    ├── document.rs     # Document abstraction
    ├── wic.rs          # WIC image loading
    ├── pdf.rs          # WinRT PDF loading
    ├── dialogs.rs      # File dialogs
    ├── menu.rs         # Context menu
    ├── theme.rs        # Dark/light theme
    └── registration.rs # File associations
```

## Technical Details

### Rendering Pipeline

1. **Images**: Loaded via Windows Imaging Component (WIC) → converted to BGRA → Direct2D bitmap
2. **PDFs**: Loaded via WinRT Windows.Data.Pdf → rendered to PNG stream → decoded via WIC → Direct2D bitmap
3. **Display**: Direct2D HwndRenderTarget with hardware acceleration

### Theme Support

- Detects Windows theme via registry (`AppsUseLightTheme`)
- Uses `DwmSetWindowAttribute` for dark title bar
- Toolbar/statusbar use `SetWindowTheme` with `DarkMode_Explorer`

### DPI Awareness

- PerMonitorV2 DPI awareness via manifest
- Handles `WM_DPICHANGED` for runtime DPI changes
- Initial window sizing uses monitor work area

## Windows Limitations

1. **File Dialog Sandboxing**: Windows' `IFileOpenDialog` cannot be fully sandboxed to a single folder. The restricted path feature sets the default folder but cannot prevent navigation.

2. **Default App Registration**: Windows 10/11 requires user consent to change default apps. The registration module adds CrossView to "Open with" but cannot force it as default.

3. **PDF Rendering**: Uses WinRT `Windows.Data.Pdf` which may have rendering differences from Adobe Reader for complex PDFs.

## Troubleshooting

### Build Errors

1. **Missing Windows SDK**: Install Windows 10/11 SDK via Visual Studio Installer
2. **Link errors**: Ensure MSVC build tools are installed
3. **Resource errors**: Check that `photo_portrait.ico` and `skn16g.png` exist

### Runtime Issues

1. **Black window**: GPU driver issue; try updating graphics drivers
2. **PDF won't open**: May be encrypted; CrossView will prompt for password
3. **Icons missing**: Ensure `skn16g.png` is in the executable directory

## License

MIT License
