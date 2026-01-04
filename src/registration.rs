//! File Association Registration for Windows 10/11
//!
//! This module provides comprehensive per-user file association registration for SimpliView.
//!
//! ## Windows Default Programs Registration
//!
//! For Windows 10/11, applications must register in multiple locations:
//! 1. **ProgID** - Defines the handler with shell\open\command. We use specific ProgIDs
//!    for different file types (PDF vs Images) for better Windows integration.
//! 2. **Capabilities** - Describes the application and its supported types.
//! 3. **RegisteredApplications** - Makes the app selectable in Windows Settings.
//! 4. **OpenWithProgids** - Adds the app to "Open with" context menu.
//!
//! ## Important Notes:
//!
//! - Windows 10/11 protects UserChoice with a hash; only Settings UI can set defaults.
//! - This registration makes SimpliView appear in "Open with" and Default Apps.
//! - User must manually choose SimpliView as default via Settings or "Open with" dialog.
//! - Per-user registration (HKCU) - no admin elevation required.

use windows::{
    core::*,
    Win32::{
        Foundation::WIN32_ERROR,
        System::Registry::*,
        UI::Shell::*,
    },
};

const APP_NAME: &str = "SimpliView";
// Changed to just "SimpliView" to ensure consistent branding in Open With/Default Apps
const APP_DESCRIPTION: &str = "SimpliView";
const APP_COMPANY: &str = "SimpliView";

// Legacy ProgID to clean up
const LEGACY_PROG_ID: &str = "SimpliView.Document.1";

struct FileTypeInfo {
    extension: &'static str,
    prog_id: &'static str,
    description: &'static str,
    perceived_type: &'static str,
    content_type: &'static str,
}

const FILE_TYPES: &[FileTypeInfo] = &[
    FileTypeInfo {
        extension: ".pdf",
        prog_id: "SimpliView.AssocFile.PDF",
        description: "SimpliView PDF Document",
        perceived_type: "Document",
        content_type: "application/pdf",
    },
    FileTypeInfo {
        extension: ".jpg",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/jpeg",
    },
    FileTypeInfo {
        extension: ".jpeg",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/jpeg",
    },
    FileTypeInfo {
        extension: ".png",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/png",
    },
    FileTypeInfo {
        extension: ".bmp",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/bmp",
    },
    FileTypeInfo {
        extension: ".tif",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/tiff",
    },
    FileTypeInfo {
        extension: ".tiff",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/tiff",
    },
    FileTypeInfo {
        extension: ".webp",
        prog_id: "SimpliView.AssocFile.Image",
        description: "SimpliView Image",
        perceived_type: "Image",
        content_type: "image/webp",
    },
];

/// Helper to check if registry operation succeeded
fn reg_ok(result: WIN32_ERROR) -> bool {
    result.0 == 0
}

/// Create a null-terminated UTF-16 string
fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Set a REG_SZ string value
unsafe fn set_string_value(hkey: HKEY, name: PCWSTR, value: &str) -> bool {
    let value_wide = to_wide(value);
    reg_ok(RegSetValueExW(
        hkey,
        name,
        0,
        REG_SZ,
        Some(std::slice::from_raw_parts(
            value_wide.as_ptr() as *const u8,
            value_wide.len() * 2,
        )),
    ))
}

/// Create a registry key and return the handle
unsafe fn create_key(parent: HKEY, subkey: &str) -> Option<HKEY> {
    let subkey_wide = to_wide(subkey);
    let mut hkey = HKEY::default();
    let mut disposition = REG_CREATE_KEY_DISPOSITION::default();

    if reg_ok(RegCreateKeyExW(
        parent,
        PCWSTR(subkey_wide.as_ptr()),
        0,
        None,
        REG_OPTION_NON_VOLATILE,
        KEY_WRITE,
        None,
        &mut hkey,
        Some(&mut disposition),
    )) {
        Some(hkey)
    } else {
        None
    }
}

/// Open an existing registry key
unsafe fn open_key(parent: HKEY, subkey: &str, access: REG_SAM_FLAGS) -> Option<HKEY> {
    let subkey_wide = to_wide(subkey);
    let mut hkey = HKEY::default();

    if reg_ok(RegOpenKeyExW(
        parent,
        PCWSTR(subkey_wide.as_ptr()),
        0,
        access,
        &mut hkey,
    )) {
        Some(hkey)
    } else {
        None
    }
}

/// Delete a registry tree
unsafe fn delete_tree(parent: HKEY, subkey: &str) {
    let subkey_wide = to_wide(subkey);
    let _ = RegDeleteTreeW(parent, PCWSTR(subkey_wide.as_ptr()));
}

/// Delete a specific value from a key
unsafe fn delete_value(hkey: HKEY, name: &str) {
    let name_wide = to_wide(name);
    let _ = RegDeleteValueW(hkey, PCWSTR(name_wide.as_ptr()));
}

/// Register SimpliView as a file handler (per-user registration)
pub fn register_file_associations() -> Result<()> {
    let exe_path = std::env::current_exe()
        .map_err(|_| Error::from_win32())?
        .to_string_lossy()
        .to_string();

    unsafe {
        // 0. Clean up legacy registration
        cleanup_legacy_registration();

        // 1. Register ProgIDs for each type
        register_prog_ids(&exe_path)?;

        // 2. Register Application Capabilities
        register_capabilities(&exe_path)?;

        // 3. Register in RegisteredApplications
        register_in_registered_applications()?;

        // 4. Register OpenWithProgids for each extension
        register_extension_mappings()?;

        // 5. Register in Applications key
        register_application(&exe_path)?;

        // 6. Notify shell of changes
        notify_shell_of_changes();
    }

    Ok(())
}

/// Unregister SimpliView file associations
pub fn unregister_file_associations() -> Result<()> {
    unsafe {
        // Remove all ProgIDs
        let prog_ids = ["SimpliView.AssocFile.PDF", "SimpliView.AssocFile.Image", LEGACY_PROG_ID];
        for pid in prog_ids {
            delete_tree(HKEY_CURRENT_USER, &format!("Software\\Classes\\{}", pid));
        }

        // Remove Capabilities
        delete_tree(HKEY_CURRENT_USER, "Software\\SimpliView");

        // Remove from RegisteredApplications
        if let Some(hkey) = open_key(
            HKEY_CURRENT_USER,
            "Software\\RegisteredApplications",
            KEY_SET_VALUE,
        ) {
            delete_value(hkey, APP_NAME);
            let _ = RegCloseKey(hkey);
        }

        // Remove extension OpenWithProgids entries
        for ft in FILE_TYPES {
            let ext_path = format!("Software\\Classes\\{}\\OpenWithProgids", ft.extension);
            if let Some(hkey) = open_key(HKEY_CURRENT_USER, &ext_path, KEY_SET_VALUE) {
                delete_value(hkey, ft.prog_id);
                delete_value(hkey, LEGACY_PROG_ID); // Clean up legacy too
                let _ = RegCloseKey(hkey);
            }
        }

        // Remove from Applications
        delete_tree(
            HKEY_CURRENT_USER,
            &format!("Software\\Classes\\Applications\\{}.exe", APP_NAME),
        );

        notify_shell_of_changes();
    }

    Ok(())
}

unsafe fn cleanup_legacy_registration() {
    delete_tree(HKEY_CURRENT_USER, &format!("Software\\Classes\\{}", LEGACY_PROG_ID));
}

/// Register the ProgIDs
unsafe fn register_prog_ids(exe_path: &str) -> Result<()> {
    // We only need to register unique ProgIDs
    let unique_prog_ids: std::collections::HashSet<&str> = FILE_TYPES.iter().map(|ft| ft.prog_id).collect();

    for prog_id in unique_prog_ids {
        // Find one file type info that uses this ProgID to get details
        let info = FILE_TYPES.iter().find(|ft| ft.prog_id == prog_id).unwrap();

        let prog_id_path = format!("Software\\Classes\\{}", prog_id);

        // Create ProgID key
        let hkey = create_key(HKEY_CURRENT_USER, &prog_id_path)
            .ok_or_else(Error::from_win32)?;

        // Set default value (friendly name for the file type)
        set_string_value(hkey, PCWSTR::null(), info.description);
        
        // Set FriendlyTypeName
        let friendly_type_name = to_wide("FriendlyTypeName");
        set_string_value(hkey, PCWSTR(friendly_type_name.as_ptr()), info.description);

        let _ = RegCloseKey(hkey);

        // Set PerceivedType
        if let Some(hkey) = open_key(HKEY_CURRENT_USER, &prog_id_path, KEY_WRITE) {
             let perceived_type = to_wide("PerceivedType");
             set_string_value(hkey, PCWSTR(perceived_type.as_ptr()), info.perceived_type);
             let _ = RegCloseKey(hkey);
        }

        // Create DefaultIcon subkey
        if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\DefaultIcon", prog_id_path)) {
            // Use index 0 for now. In a real app we might want specific icons for PDF vs Image
            let icon_value = format!("{},0", exe_path);
            set_string_value(hkey, PCWSTR::null(), &icon_value);
            let _ = RegCloseKey(hkey);
        }

        // Create shell\open subkey with FriendlyAppName
        if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\shell\\open", prog_id_path)) {
            let friendly_app_name = to_wide("FriendlyAppName");
            set_string_value(hkey, PCWSTR(friendly_app_name.as_ptr()), APP_NAME);
            let _ = RegCloseKey(hkey);
        }

        // Create shell\open\command subkey
        if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\shell\\open\\command", prog_id_path)) {
            let command_value = format!("\"{}\" \"%1\"", exe_path);
            set_string_value(hkey, PCWSTR::null(), &command_value);
            let _ = RegCloseKey(hkey);
        }
    }

    Ok(())
}

/// Register application capabilities
unsafe fn register_capabilities(exe_path: &str) -> Result<()> {
    // Create Capabilities key
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, "Software\\SimpliView\\Capabilities") {
        let app_name = to_wide("ApplicationName");
        set_string_value(hkey, PCWSTR(app_name.as_ptr()), APP_NAME);

        let app_desc = to_wide("ApplicationDescription");
        set_string_value(hkey, PCWSTR(app_desc.as_ptr()), APP_DESCRIPTION);
        
        let app_comp = to_wide("ApplicationCompany");
        set_string_value(hkey, PCWSTR(app_comp.as_ptr()), APP_COMPANY);

        // ApplicationIcon
        let app_icon = to_wide("ApplicationIcon");
        set_string_value(hkey, PCWSTR(app_icon.as_ptr()), &format!("{},0", exe_path));

        let _ = RegCloseKey(hkey);
    }

    // Create FileAssociations subkey
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, "Software\\SimpliView\\Capabilities\\FileAssociations") {
        for ft in FILE_TYPES {
            let ext_wide = to_wide(ft.extension);
            set_string_value(hkey, PCWSTR(ext_wide.as_ptr()), ft.prog_id);
        }
        let _ = RegCloseKey(hkey);
    }

    Ok(())
}

/// Register in RegisteredApplications
unsafe fn register_in_registered_applications() -> Result<()> {
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, "Software\\RegisteredApplications") {
        let app_name_wide = to_wide(APP_NAME);
        set_string_value(
            hkey,
            PCWSTR(app_name_wide.as_ptr()),
            "Software\\SimpliView\\Capabilities",
        );
        let _ = RegCloseKey(hkey);
    }
    Ok(())
}

/// Register extension mappings (OpenWithProgids)
unsafe fn register_extension_mappings() -> Result<()> {
    for ft in FILE_TYPES {
        let ext_path = format!("Software\\Classes\\{}\\OpenWithProgids", ft.extension);

        if let Some(hkey) = create_key(HKEY_CURRENT_USER, &ext_path) {
            // Add ProgID as value
            let prog_id_wide = to_wide(ft.prog_id);
            let empty: Vec<u16> = vec![0];
            let _ = RegSetValueExW(
                hkey,
                PCWSTR(prog_id_wide.as_ptr()),
                0,
                REG_SZ,
                Some(std::slice::from_raw_parts(
                    empty.as_ptr() as *const u8,
                    2,
                )),
            );
            let _ = RegCloseKey(hkey);
        }
    }
    Ok(())
}

/// Register in Applications key
unsafe fn register_application(exe_path: &str) -> Result<()> {
    let app_path = format!("Software\\Classes\\Applications\\{}.exe", APP_NAME);

    // Create application key
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, &app_path) {
        let friendly_app = to_wide("FriendlyAppName");
        set_string_value(hkey, PCWSTR(friendly_app.as_ptr()), APP_NAME);

        // Also set ApplicationCompany and Description here for good measure
        let app_desc = to_wide("ApplicationDescription");
        set_string_value(hkey, PCWSTR(app_desc.as_ptr()), APP_DESCRIPTION);

        let _ = RegCloseKey(hkey);
    }

    // Create DefaultIcon subkey
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\DefaultIcon", app_path)) {
        set_string_value(hkey, PCWSTR::null(), &format!("{},0", exe_path));
        let _ = RegCloseKey(hkey);
    }

    // Create shell\open\command subkey
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\shell\\open\\command", app_path)) {
        let command_value = format!("\"{}\" \"%1\"", exe_path);
        set_string_value(hkey, PCWSTR::null(), &command_value);
        let _ = RegCloseKey(hkey);
    }

    // Create SupportedTypes subkey
    if let Some(hkey) = create_key(HKEY_CURRENT_USER, &format!("{}\\SupportedTypes", app_path)) {
        for ft in FILE_TYPES {
            let ext_wide = to_wide(ft.extension);
            let empty: Vec<u16> = vec![0];
            let _ = RegSetValueExW(
                hkey,
                PCWSTR(ext_wide.as_ptr()),
                0,
                REG_SZ,
                Some(std::slice::from_raw_parts(
                    empty.as_ptr() as *const u8,
                    2,
                )),
            );
        }
        let _ = RegCloseKey(hkey);
    }

    Ok(())
}

/// Notify Windows Shell of file association changes
fn notify_shell_of_changes() {
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }
}

/// Check if SimpliView is registered
#[allow(dead_code)]
pub fn is_registered() -> bool {
    // Basic check for one of our ProgIDs
    unsafe {
        open_key(
            HKEY_CURRENT_USER,
            "Software\\Classes\\SimpliView.AssocFile.PDF",
            KEY_READ,
        ).map(|hkey| {
            let _ = RegCloseKey(hkey);
            true
        }).unwrap_or(false)
    }
}

/// Get registration status for diagnostics
pub fn get_registration_status() -> Vec<(String, bool)> {
    let mut results = Vec::new();

    unsafe {
        // Check Capabilities
        let caps_ok = open_key(
            HKEY_CURRENT_USER,
            "Software\\SimpliView\\Capabilities",
            KEY_READ,
        ).map(|hkey| {
            let _ = RegCloseKey(hkey);
            true
        }).unwrap_or(false);
        results.push(("Capabilities".to_string(), caps_ok));

        // Check each file type
        for ft in FILE_TYPES {
            // ProgID
             let prog_id_ok = open_key(
                HKEY_CURRENT_USER,
                &format!("Software\\Classes\\{}", ft.prog_id),
                KEY_READ,
            ).map(|hkey| {
                let _ = RegCloseKey(hkey);
                true
            }).unwrap_or(false);
            results.push((format!("ProgID {}", ft.prog_id), prog_id_ok));

            // OpenWithProgids
            let ext_ok = open_key(
                HKEY_CURRENT_USER,
                &format!("Software\\Classes\\{}\\OpenWithProgids", ft.extension),
                KEY_READ,
            ).map(|hkey| {
                let prog_id_wide = to_wide(ft.prog_id);
                let mut data_size = 0u32;
                let result = RegQueryValueExW(
                    hkey,
                    PCWSTR(prog_id_wide.as_ptr()),
                    None,
                    None,
                    None,
                    Some(&mut data_size),
                );
                let _ = RegCloseKey(hkey);
                reg_ok(result)
            }).unwrap_or(false);
            results.push((format!("OpenWithProgids for {}", ft.extension), ext_ok));
        }
    }

    results
}
