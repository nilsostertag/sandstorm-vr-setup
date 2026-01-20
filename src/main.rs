use std::{
    env, fs,
    path::{Path, PathBuf},
    process::{Command, Child, exit},
    time::{Instant, Duration},
    thread,
    os::windows::ffi::OsStrExt,
};
use widestring::U16CString;

use windows::{
    core::{PCWSTR, Interface},
    Win32::{
        System::Com::{CoCreateInstance, CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, CLSCTX_INPROC_SERVER, IPersistFile},
        UI::Shell::{IShellLinkW, ShellLink},
    },
};
use sysinfo::{Pid, ProcessRefreshKind, RefreshKind, System};

fn main() {
    println!("[*] start");

    let mut renamed = false;
    let mut copied = false;

    if let Err(e) = env::set_current_dir("E:\\SteamLibrary\\steamapps\\common\\sandstorm") {
        eprintln!("[!] failed to enter workingdir: {e}");
        return;
    }
    println!("[+] cd workingdir");

    // rename + copy
    if Path::new("InsurgencyEAC.exe").exists() {
        if let Err(e) = fs::rename("InsurgencyEAC.exe", "InsurgencyEACg.exe") {
            eprintln!("[!] rename failed: {e}");
            rollback(renamed, copied);
            return;
        }
        renamed = true;
        println!("[+] InsurgencyEAC.exe -> InsurgencyEACg.exe");
    }

    if let Err(e) = fs::copy("Insurgency.exe", "InsurgencyEAC.exe") {
        eprintln!("[!] copy failed: {e}");
        rollback(renamed, copied);
        return;
    }
    copied = true;
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

    //Prüfen, wie weit der code fehlerfrei läuft
    //exit(0);

    // Spiel starten über Shortcut
    let mut _child: Child = match Command::new("cmd")
        .args(["/C", "start", "", &shortcut.to_string_lossy()])
        .spawn()
    {
        Ok(c) => c,
        Err(e) => {
            eprintln!("[!] failed to start game via shortcut: {e}");
            rollback(renamed, copied);
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
        rollback(renamed, copied);
        return;
    }
    // Watchdog-Schleife: überwacht das Insurgency-Process
    if let Some(pid) = game_pid {
        let pid = Pid::from_u32(pid); // sysinfo benötigt Pid-Typ

        println!("[*] Monitoring Insurgency process...");

        loop {
        // Prozesse aktualisieren
        sys.refresh_specifics(RefreshKind::everything());

        // Prüfen, ob der Prozess noch existiert
        if sys.process(pid).is_none() {
            println!("[!] Insurgency has exited or lost connection, performing rollback...");
            rollback(renamed, copied);
            println!("[*] Exiting script.");
            return;
        }
        std::thread::sleep(std::time::Duration::from_secs(1));
        }
    }
}

fn rollback(renamed: bool, copied: bool) {
    println!("[*] rollback");
    let _ = fs::rename("InsurgencyEAC.exe", "DELETE.exe");
    let _ = fs::rename("InsurgencyEACg.exe", "InsurgencyEAC.exe");
    let _ = fs::remove_file("DELETE.exe");
    let __ = fs::remove_file("InsurgencyVR.exe.lnk");

    println!("    removed InsurgencyEAC.exe");
    println!("    restored InsurgencyEAC.exe from archive");
    println!("[*] rollback finished");
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
