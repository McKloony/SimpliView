# SimpliView

High-performance PDF and Image viewer for Windows.

## Features

- Fast PDF rendering using Windows PDF API
- Image viewing support (PNG, JPG, BMP, etc.)
- Zoom and rotation controls
- Keyboard navigation
- Lightweight and native Windows application

## Requirements

- Windows 10 or later
- Rust toolchain (for building from source)

## Building

```bash
cargo build --release
```

The executable will be created at `target/release/SimpliView.exe`.

## Usage

```bash
SimpliView.exe [file_path]
```

Or simply double-click on the executable and use the File menu to open documents.

## License

MIT License
