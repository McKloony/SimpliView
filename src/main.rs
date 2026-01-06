#![windows_subsystem = "windows"]

mod app;
mod d2d;
mod dialogs;
mod document;
mod icons;
mod menu;
mod pdf;
mod registration;
mod scroll;
mod statusbar;
mod theme;
mod toolbar;
mod utils;
mod view_window;
mod wic;
mod window;

use app::App;
use std::env;
use windows::{
    core::*,
    Win32::System::Com::*,
};

fn main() -> Result<()> {
    // Initialize COM for WIC and WinRT
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)?;
    }

    // Parse command-line arguments
    let args: Vec<String> = env::args().collect();
    
    // Handle registration commands
    if args.contains(&String::from("--register")) {
        match registration::register_file_associations() {
            Ok(_) => show_message("Erfolg", "Dateiverknüpfungen wurden erfolgreich registriert."),
            Err(e) => show_message("Fehler", &format!("Dateiverknüpfungen konnten nicht registriert werden: {:?}", e)),
        }
        unsafe { CoUninitialize(); }
        return Ok(());
    }

    if args.contains(&String::from("--unregister")) {
        match registration::unregister_file_associations() {
            Ok(_) => show_message("Erfolg", "Dateiverknüpfungen wurden erfolgreich entfernt."),
            Err(e) => show_message("Fehler", &format!("Dateiverknüpfungen konnten nicht entfernt werden: {:?}", e)),
        }
        unsafe { CoUninitialize(); }
        return Ok(());
    }

    if args.contains(&String::from("--diagnose")) {
        let status = registration::get_registration_status();
        let mut report = String::from("SimpliView Registration Status:\n\n");
        let mut all_ok = true;
        for (name, ok) in &status {
            let symbol = if *ok { "✓" } else { "✗" };
            report.push_str(&format!("{} {}\n", symbol, name));
            if !*ok {
                all_ok = false;
            }
        }
        report.push_str(&format!("\n{}", if all_ok {
            "All registrations OK. Use Windows Settings to set SimpliView as default."
        } else {
            "Some registrations missing. Run --register first."
        }));
        show_message("SimpliView Diagnostics", &report);
        unsafe { CoUninitialize(); }
        return Ok(());
    }

    let mut file_to_open = None;
    let mut restricted_path = None;
    
    // Parse arguments
    let mut i = 1;
    while i < args.len() {
        let arg = &args[i];
        if arg == "--restricted" {
             if i + 1 < args.len() {
                 restricted_path = Some(args[i+1].clone());
                 i += 1;
             }
        } else if !arg.starts_with("--") {
            if file_to_open.is_none() {
                file_to_open = Some(arg.clone());
            } else if restricted_path.is_none() {
                restricted_path = Some(arg.clone());
            }
        }
        i += 1;
    }

    // Validate restricted path if provided
    if let Some(ref path) = restricted_path {
        // Check for common CLI quoting errors (quote inside string)
        if path.contains('\"') {
            show_message("Fehler beim Starten", "Der angegebene Pfad enthält ungültige Zeichen (Anführungszeichen).\nMöglicherweise wurde ein abschließender Backslash vor dem Anführungszeichen verwendet (z.B. \"C:\\Pfad\\\").\nBitte verwenden Sie \"C:\\Pfad\" oder \"C:\\Pfad\\\\\".");
            return Ok(());
        }

        let p = std::path::Path::new(path);
        if !p.exists() {
            show_message("Fehler beim Starten", &format!("Der eingeschränkte Speicherpfad existiert nicht:\n{}", path));
            return Ok(());
        }
        if !p.is_dir() {
            show_message("Fehler beim Starten", &format!("Der eingeschränkte Speicherpfad ist kein Verzeichnis:\n{}", path));
            return Ok(());
        }
    }

    // Create and run the application
    let mut app = App::new(file_to_open, restricted_path)?;
    let result = app.run();

    // Cleanup COM
    unsafe {
        CoUninitialize();
    }

    result
}

fn show_message(title: &str, message: &str) {
    // Print to stdout for CLI usage
    println!("{}: {}", title, message);

    use windows::Win32::UI::WindowsAndMessaging::{MessageBoxW, MB_OK, MB_ICONINFORMATION};
    
    let title_wide: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
    let message_wide: Vec<u16> = message.encode_utf16().chain(std::iter::once(0)).collect();
    
    unsafe {
        MessageBoxW(
            None,
            PCWSTR(message_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            MB_OK | MB_ICONINFORMATION,
        );
    }
}
