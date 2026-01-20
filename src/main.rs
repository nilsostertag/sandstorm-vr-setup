use std::{
    env, fs,
    path::{Path, PathBuf},
    process::{Command, Child},
    os::windows::ffi::OsStrExt,
    error::Error,
};
use widestring::U16CString;

use windows::{
    core::{PCWSTR, Interface},
    Win32::{
        System::Com::{CoCreateInstance, CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, CLSCTX_INPROC_SERVER, IPersistFile},
        UI::Shell::{IShellLinkW, ShellLink},
    },
};
use sysinfo::{Pid, RefreshKind, System};
use serde::Deserialize;

// Struktur für die JSON
#[derive(Deserialize)]
struct Config {
    sandstorm_dir: String,
    uevr_path: String,
}


fn main() {
    println!("[*] start");
    println!("[*] reading config");
    let config = load_config().expect("Failed to load configuration file.");

    let sandstorm_dir: &str = &config.sandstorm_dir;
    let uevr_path: &str = &config.uevr_path;

    let mut _uevr: Child = match Command::new("cmd")
        .args(["/C", "start", "", uevr_path])
        .spawn()
    {
        Ok(c) => c,
        Err(e) => {
            eprintln!("[!] failed to start UEVR_Injector: {e}");
            rollback();
            return;
        }
    };

    if let Err(e) = env::set_current_dir(sandstorm_dir) {
        eprintln!("[!] failed to enter workingdir: {e}");
        return;
    }
    println!("[+] enter working directory");

    // rename + copy
    if Path::new("InsurgencyEAC.exe").exists() {
        if let Err(e) = fs::rename("InsurgencyEAC.exe", "InsurgencyEACg.exe") {
            eprintln!("[!] rename failed: {e}");
            rollback();
            return;
        }
        println!("[+] InsurgencyEAC.exe -> InsurgencyEACg.exe");
    }

    if let Err(e) = fs::copy("Insurgency.exe", "InsurgencyEAC.exe") {
        eprintln!("[!] copy failed: {e}");
        rollback();
        return;
    }
    println!("[+] Insurgency.exe -> InsurgencyEAC.exe");

    // Shortcut erstellen & Desktop
    let shortcut = match create_and_copy_shortcut() {
        Ok(path) => {
            println!("[+] Shortcut created: {:?}", path);
            path
        }
        Err(e) => {
            eprintln!("[!] failed to create/copy shortcut: {e:?}");
            return;
        }
    };

    // Spiel starten über Shortcut
    let mut _child: Child = match Command::new("cmd")
        .args(["/C", "start", "", &shortcut.to_string_lossy()])
        .spawn()
    {
        Ok(c) => c,
        Err(e) => {
            eprintln!("[!] failed to start game via shortcut: {e}");
            rollback();
            return;
        }
    };

    let target_name = "InsurgencyClient-Win64-Shipping.exe";
    let mut sys = System::new_with_specifics(RefreshKind::everything());
    let mut game_pid: Option<u32> = None;

    std::thread::sleep(std::time::Duration::from_secs(10));

    for _ in 0..30 {
        println!("[+] Looking for target process `{target_name}`");
        print!(".");
        sys.refresh_specifics(RefreshKind::everything());
        // Search for the target process
        if let Some((pid, _process)) = sys.processes()
            .iter()
            .find(|(_, process)| process.name().eq_ignore_ascii_case(target_name))
        {
            game_pid = Some(pid.as_u32()); // `.as_u32()` gives u32
            println!("[+] Found target process `{target_name}` (pid={})", pid.as_u32());
            break;
        }
        std::thread::sleep(std::time::Duration::from_secs(1));
    }
    
    if let Some(pid) = game_pid {
        println!("[*] Attaching to process pid={pid}");
    } else {
        eprintln!("[!] Failed to find `{target_name}` after 30s");
        rollback();
        return;
    }
    // Watchdog-Schleife: überwacht das Insurgency-Process
    if let Some(pid) = game_pid {
        let pid = Pid::from_u32(pid);

        println!("[*] Monitoring Insurgency process...");

        loop {
        // Prozesse aktualisieren
        sys.refresh_specifics(RefreshKind::everything());

        // Prüfen, ob der Prozess noch existiert
        if sys.process(pid).is_none() {
            println!("[!] Insurgency has exited or lost connection, performing rollback...");
            rollback();
            println!("[*] Exiting script.");
            return;
        }
        std::thread::sleep(std::time::Duration::from_secs(1));
        }
    }
}

fn load_config() -> Result<Config, Box<dyn Error>> {
    let exe_path = std::env::current_exe()?;
    let exe_dir = exe_path
        .parent()
        .ok_or("Failed to get exe directory")?;
    let config_path = exe_dir.join("sandstorm_vr_setup.json");

    let content = fs::read_to_string(config_path)?;
    let config: Config = serde_json::from_str(&content)?;

    Ok(config)
}

fn rollback() {
    println!("[*] rollback");
    let _ = fs::rename("InsurgencyEAC.exe", "DELETE.exe");
    std::thread::sleep(std::time::Duration::from_secs(1));
    let _ = fs::rename("InsurgencyEACg.exe", "InsurgencyEAC.exe");
    let _ = fs::remove_file("DELETE.exe");
    let __ = fs::remove_file("InsurgencyVR.exe.lnk");
    
    std::thread::sleep(std::time::Duration::from_secs(1));
    println!("    removed InsurgencyEAC.exe");
    println!("    restored InsurgencyEAC.exe from archive");
    println!("[*] rollback finished");
    std::thread::sleep(std::time::Duration::from_secs(1));
}

pub fn create_and_copy_shortcut() -> windows::core::Result<PathBuf> {
    // COM initialisieren
    unsafe {
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(hr.into());
        }
    }

    // Shortcut über CoCreateInstance erzeugen
    let shell_link: IShellLinkW = unsafe { CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)? };

    // Zielprogramm
    let exe_path = env::current_dir()?.join("Insurgency.exe");
    let exe_utf16: Vec<u16> = exe_path.as_os_str().encode_wide().chain(Some(0)).collect();
    unsafe { shell_link.SetPath(windows::core::PCWSTR(exe_utf16.as_ptr()))? };

    // Argumente
    let args_utf16: Vec<u16> = "-eac-nop-loaded".encode_utf16().chain(Some(0)).collect();
    unsafe { shell_link.SetArguments(windows::core::PCWSTR(args_utf16.as_ptr()))? };

    // Arbeitsverzeichnis
    let cwd = env::current_dir()?;
    let cwd_utf16: Vec<u16> = cwd.as_os_str().encode_wide().chain(Some(0)).collect();
    unsafe { shell_link.SetWorkingDirectory(windows::core::PCWSTR(cwd_utf16.as_ptr()))? };

    // Shortcut speichern im Working Directory
    let workingdir_shortcut = PathBuf::from("InsurgencyVR.exe.lnk");
    let persist_file: IPersistFile = shell_link.cast()?;
    let shortcut_utf16 = U16CString::from_os_str(&workingdir_shortcut).unwrap();
    let shortcut_pcwstr = PCWSTR(shortcut_utf16.as_ptr());
    unsafe { persist_file.Save(shortcut_pcwstr, true)? };

    // COM freigeben
    unsafe { CoUninitialize() };

    // Pfad im Working Directory zurückgeben
    Ok(workingdir_shortcut)
}
