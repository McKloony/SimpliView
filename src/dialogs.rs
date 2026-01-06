use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::Com::*,
        System::LibraryLoader::*,
        UI::{
            Input::KeyboardAndMouse::SetFocus,
            Shell::*,
            Shell::Common::*,
            WindowsAndMessaging::*,
        },
    },
};

// File type filter
const FILE_TYPES: &[COMDLG_FILTERSPEC] = &[
    COMDLG_FILTERSPEC {
        pszName: w!("All Supported Files"),
        pszSpec: w!("*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.webp"),
    },
    COMDLG_FILTERSPEC {
        pszName: w!("PDF Documents"),
        pszSpec: w!("*.pdf"),
    },
    COMDLG_FILTERSPEC {
        pszName: w!("Images"),
        pszSpec: w!("*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.webp"),
    },
    COMDLG_FILTERSPEC {
        pszName: w!("All Files"),
        pszSpec: w!("*.*"),
    },
];

const SAVE_TYPES: &[COMDLG_FILTERSPEC] = &[
    COMDLG_FILTERSPEC { pszName: w!("PNG Image (*.png)"), pszSpec: w!("*.png") },
    COMDLG_FILTERSPEC { pszName: w!("JPEG Image (*.jpg;*.jpeg)"), pszSpec: w!("*.jpg;*.jpeg") },
    COMDLG_FILTERSPEC { pszName: w!("BMP Image (*.bmp)"), pszSpec: w!("*.bmp") },
    COMDLG_FILTERSPEC { pszName: w!("TIFF Image (*.tif;*.tiff)"), pszSpec: w!("*.tif;*.tiff") },
    COMDLG_FILTERSPEC { pszName: w!("WebP Image (*.webp)"), pszSpec: w!("*.webp") },
    COMDLG_FILTERSPEC { pszName: w!("PDF Document (*.pdf)"), pszSpec: w!("*.pdf") },
];

fn get_save_type_index(ext: &str) -> (u32, PCWSTR) {
    match ext.to_lowercase().as_str() {
        "jpg" | "jpeg" => (2, w!("jpg")),
        "bmp" => (3, w!("bmp")),
        "tif" | "tiff" => (4, w!("tif")),
        "webp" => (5, w!("webp")),
        "pdf" => (6, w!("pdf")),
        _ => (1, w!("png")), // PNG is default (index 1)
    }
}

pub struct FileDialogs {
    pub restricted_path: Option<String>,
}

impl FileDialogs {
    pub fn new(restricted_path: Option<String>) -> Self {
        Self { restricted_path }
    }

    pub fn open_file(&self, parent: HWND) -> Option<String> {
        unsafe {
            // Create file open dialog
            let dialog: IFileOpenDialog =
                CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER).ok()?;

            // Set file types
            dialog.SetFileTypes(FILE_TYPES).ok()?;
            dialog.SetFileTypeIndex(1).ok()?;

            // Set options
            let options = dialog.GetOptions().ok()?;
            dialog.SetOptions(options | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST).ok()?;

            // Show dialog
            if dialog.Show(parent).is_err() {
                return None;
            }

            // Get result
            let result = dialog.GetResult().ok()?;
            let path = result.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
            let path_str = path.to_string().ok()?;
            CoTaskMemFree(Some(path.0 as *const _));

            Some(path_str)
        }
    }

    pub fn save_file(&self, parent: HWND, default_filename: Option<&str>, original_extension: Option<&str>) -> Option<String> {
        loop {
            unsafe {
                // Create file save dialog
                let dialog: IFileSaveDialog = match CoCreateInstance(&FileSaveDialog, None, CLSCTX_INPROC_SERVER) {
                    Ok(d) => d,
                    Err(_) => return None,
                };

                // Get default filter index and extension based on original file extension
                let ext = original_extension.unwrap_or("png");
                let (index, default_ext) = get_save_type_index(ext);

                // Set file types - allow all supported save formats
                if dialog.SetFileTypes(SAVE_TYPES).is_err() { return None; }
                if dialog.SetFileTypeIndex(index).is_err() { return None; }
                if dialog.SetDefaultExtension(default_ext).is_err() { return None; }

                // Set default filename if provided
                if let Some(filename) = default_filename {
                    let name_wide: Vec<u16> = filename.encode_utf16().chain(std::iter::once(0)).collect();
                    let _ = dialog.SetFileName(PCWSTR(name_wide.as_ptr()));
                }

                // Set options
                if let Ok(options) = dialog.GetOptions() {
                    let _ = dialog.SetOptions(options | FOS_FORCEFILESYSTEM | FOS_OVERWRITEPROMPT);
                }

                // Apply restricted path if specified
                if let Some(ref path) = self.restricted_path {
                    self.apply_folder_restriction_save(&dialog, path);
                }

                // Show dialog
                if dialog.Show(parent).is_err() {
                    return None;
                }

                // Get result
                let result = match dialog.GetResult() {
                    Ok(r) => r,
                    Err(_) => return None,
                };
                
                let path_item = match result.GetDisplayName(SIGDN_FILESYSPATH) {
                    Ok(p) => p,
                    Err(_) => return None,
                };
                
                let path_str = match path_item.to_string() {
                    Ok(s) => s,
                    Err(_) => {
                        CoTaskMemFree(Some(path_item.0 as *const _));
                        return None;
                    }
                };
                CoTaskMemFree(Some(path_item.0 as *const _));

                // Validate restriction
                if let Some(ref restricted) = self.restricted_path {
                    // Normalize paths for comparison (lowercase, handle separators)
                    let p_lower = path_str.to_lowercase();
                    let r_lower = restricted.to_lowercase();
                    
                    // Simple check: does the selected path start with the restricted path?
                    // We also ensure the restricted path ends with a separator for correct prefix matching
                    // unless it's just a drive root like "C:\"
                    let r_clean = r_lower.trim_end_matches('\\');
                    
                    if !p_lower.starts_with(r_clean) {
                        show_error(parent, &format!("Speichern ist nur im Ordner '{}' erlaubt.", restricted));
                        continue; // Re-open dialog
                    }
                }

                return Some(path_str);
            }
        }
    }

    fn apply_folder_restriction<T: windows::core::Interface + windows::core::ComInterface>(&self, dialog: &T, path: &str) {
        unsafe {
            let path_wide: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();

            // Get shell item for the folder
            if let Ok(folder) = SHCreateItemFromParsingName::<_, _, IShellItem>(PCWSTR(path_wide.as_ptr()), None) {
                if let Ok(file_dialog) = dialog.cast::<IFileDialog>() {
                    let _ = file_dialog.SetDefaultFolder(&folder);
                    let _ = file_dialog.SetFolder(&folder);

                    if let Ok(options) = file_dialog.GetOptions() {
                        let _ = file_dialog.SetOptions(options | FOS_FORCEFILESYSTEM);
                    }
                }
            }
        }
    }

    fn apply_folder_restriction_save<T: windows::core::Interface + windows::core::ComInterface>(&self, dialog: &T, path: &str) {
        self.apply_folder_restriction(dialog, path);
    }
}

// Resource IDs (Must match .rc file)
const IDD_PASSWORD_DIALOG: isize = 200;
const IDC_PASSWORD_EDIT: i32 = 201;

struct PasswordData {
    password: Option<String>,
}

// Password dialog for encrypted PDFs using standard Resource Dialog
pub fn password_dialog(parent: HWND) -> Option<String> {
    unsafe {
        let instance = GetModuleHandleW(None).unwrap_or_default();
        let mut data = PasswordData { password: None };

        let result = DialogBoxParamW(
            instance,
            PCWSTR(IDD_PASSWORD_DIALOG as *const u16),
            parent,
            Some(password_dialog_proc),
            LPARAM(&mut data as *mut _ as isize),
        );

        if result == IDOK.0 as isize {
            data.password
        } else {
            None
        }
    }
}

/// Dialog procedure for the password input dialog.
///
/// Keyboard handling:
/// - Enter: Submits password (DEFPUSHBUTTON handles this)
/// - ESC: Cancels dialog (WM_CLOSE handler)
/// - Tab: Navigates between controls (handled by dialog manager)
///
/// The password is NOT logged or persisted - it's passed directly to the PDF loader
/// and then dropped when the string goes out of scope.
extern "system" fn password_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> isize {
    unsafe {
        match msg {
            WM_INITDIALOG => {
                // Store the pointer to PasswordData in window's user data
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, lparam.0);

                // Center the dialog relative to parent window
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    let mut rc_parent = RECT::default();
                    let mut rc_dlg = RECT::default();
                    GetWindowRect(parent, &mut rc_parent);
                    GetWindowRect(hwnd, &mut rc_dlg);

                    let dlg_w = rc_dlg.right - rc_dlg.left;
                    let dlg_h = rc_dlg.bottom - rc_dlg.top;
                    let parent_w = rc_parent.right - rc_parent.left;
                    let parent_h = rc_parent.bottom - rc_parent.top;

                    let x = rc_parent.left + (parent_w - dlg_w) / 2;
                    let y = rc_parent.top + (parent_h - dlg_h) / 2;

                    SetWindowPos(hwnd, None, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                }

                // Set focus to the password input field
                let edit = GetDlgItem(hwnd, IDC_PASSWORD_EDIT);
                SetFocus(edit);

                // Return 0 (FALSE) to indicate we set focus manually
                0
            }
            WM_COMMAND => {
                let id = (wparam.0 & 0xFFFF) as i32;
                match id {
                    1 => {
                        // IDOK - User pressed OK or Enter
                        let data_ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut PasswordData;
                        if !data_ptr.is_null() {
                            let edit = GetDlgItem(hwnd, IDC_PASSWORD_EDIT);
                            let len = GetWindowTextLengthW(edit);
                            // Allow empty passwords (some PDFs may have empty password)
                            let mut buffer = vec![0u16; (len + 1).max(1) as usize];
                            GetWindowTextW(edit, &mut buffer);
                            // Convert UTF-16 to Rust String (supports Unicode passwords)
                            (*data_ptr).password = Some(String::from_utf16_lossy(&buffer[..len as usize]));
                        }
                        EndDialog(hwnd, IDOK.0 as isize);
                        1
                    }
                    2 => {
                        // IDCANCEL - User pressed Cancel or ESC
                        EndDialog(hwnd, IDCANCEL.0 as isize);
                        1
                    }
                    _ => 0
                }
            }
            WM_CLOSE => {
                // Handle window close button (X) - treat as cancel
                EndDialog(hwnd, IDCANCEL.0 as isize);
                1
            }
            _ => 0,
        }
    }
}

pub fn retry_password_dialog(parent: HWND) -> bool {
    unsafe {
        let result = MessageBoxW(
            parent,
            w!("Falsches Kennwort. Erneut versuchen?"),
            w!("Kennwortfehler"),
            MB_YESNO | MB_ICONWARNING,
        );
        result == IDYES
    }
}

pub fn show_error(parent: HWND, message: &str) {
    let message_wide: Vec<u16> = message.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        MessageBoxW(
            parent,
            PCWSTR(message_wide.as_ptr()),
            w!("Fehler"),
            MB_OK | MB_ICONERROR,
        );
    }
}

pub fn show_info(parent: HWND, title: &str, message: &str) {
    let title_wide: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
    let message_wide: Vec<u16> = message.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let instance = GetModuleHandleW(None).unwrap_or_default();
        let params = MSGBOXPARAMSW {
            cbSize: std::mem::size_of::<MSGBOXPARAMSW>() as u32,
            hwndOwner: parent,
            hInstance: instance,
            lpszText: PCWSTR(message_wide.as_ptr()),
            lpszCaption: PCWSTR(title_wide.as_ptr()),
            dwStyle: MB_OK | MB_USERICON,
            #[allow(clippy::manual_dangling_ptr)] // MAKEINTRESOURCEW(1) - intentional
            lpszIcon: PCWSTR(1 as *const u16),
            dwContextHelpId: 0,
            lpfnMsgBoxCallback: None,
            dwLanguageId: 0,
        };
        MessageBoxIndirectW(&params);
    }
}