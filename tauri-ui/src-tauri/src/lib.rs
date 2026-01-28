use std::path::PathBuf;
use std::process::Command;
use tauri::Manager;

/// Find the engine directory by looking for the engine module.
fn find_engine_dir() -> Result<PathBuf, String> {
    // Try multiple strategies to find the engine directory

    // Strategy 1: Current directory's parent (works in dev mode from tauri-ui)
    if let Ok(cwd) = std::env::current_dir() {
        if let Some(parent) = cwd.parent() {
            let engine_path = parent.join("engine");
            if engine_path.exists() {
                return Ok(parent.to_path_buf());
            }
        }
        // Maybe we're already in the project root
        let engine_path = cwd.join("engine");
        if engine_path.exists() {
            return Ok(cwd);
        }
    }

    // Strategy 2: Look relative to executable
    if let Ok(exe_path) = std::env::current_exe() {
        // In dev: target/debug/draftmate -> go up to find project root
        let mut dir = exe_path.parent();
        for _ in 0..5 {
            if let Some(d) = dir {
                let engine_path = d.join("engine");
                if engine_path.exists() {
                    return Ok(d.to_path_buf());
                }
                dir = d.parent();
            }
        }
    }

    // Strategy 3: Hardcoded fallback for development
    let dev_path = PathBuf::from("/Users/arinaggarwal/Documents/IB Prep Materials/Draftmate v3");
    if dev_path.join("engine").exists() {
        return Ok(dev_path);
    }

    Err("Could not find engine directory. Make sure the 'engine' folder exists.".to_string())
}

/// Run the Python engine CLI and return the JSON output.
/// This is the bridge between Tauri frontend and Python backend.
#[tauri::command]
fn run_engine(args: Vec<String>) -> Result<String, String> {
    let engine_dir = find_engine_dir()?;

    let output = Command::new("python3")
        .arg("-m")
        .arg("engine")
        .args(&args)
        .current_dir(&engine_dir)
        .output()
        .map_err(|e| format!("Failed to execute Python engine: {}", e))?;

    if output.status.success() {
        String::from_utf8(output.stdout)
            .map_err(|e| format!("Failed to parse stdout: {}", e))
    } else {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        Err(format!(
            "Engine command failed: {}{}",
            stderr,
            if stdout.is_empty() { String::new() } else { format!("\nOutput: {}", stdout) }
        ))
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![run_engine])
        .setup(|app| {
            #[cfg(debug_assertions)]
            {
                let window = app.get_webview_window("main").unwrap();
                window.open_devtools();
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
