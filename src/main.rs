#![windows_subsystem = "windows"]

use std::cell::RefCell;
use std::fmt::Debug;
use std::fs::{self, OpenOptions};
use std::io::{self, Read};
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::mpsc;
use std::thread;
use std::time::Duration;

use anyhow::Result;
use chrono::{Datelike, NaiveDate};
use lazy_static::lazy_static;
use nwd::NwgUi;
use nwg::NativeUi;
use regex::Regex;
use winapi::um::{
    combaseapi::CoTaskMemFree,
    knownfolders::FOLDERID_LocalAppDataLow,
    shellapi::ShellExecuteW,
    shlobj::{SHGetKnownFolderPath, KF_FLAG_DEFAULT},
    winuser::SW_SHOWNORMAL,
};

const VENDOR_NAME: &str = "KOBA789";
const APP_NAME: &str = "VRCLogUnrotate";

struct CrashNotifier(Option<nwg::NoticeSender>);
impl CrashNotifier {
    fn disable(&mut self) {
        self.0 = None;
    }
}
impl Drop for CrashNotifier {
    fn drop(&mut self) {
        if let Some(sender) = self.0 {
            sender.notice();
        }
    }
}

#[derive(Default, NwgUi)]
pub struct SystemTray {
    #[nwg_control]
    #[nwg_events(OnInit: [SystemTray::init])]
    window: nwg::MessageWindow,

    #[nwg_resource]
    embed: nwg::EmbedResource,

    #[nwg_resource(source_embed: Some(&data.embed), source_embed_str: Some("MAINICON"))]
    icon: nwg::Icon,

    #[nwg_control(icon: Some(&data.icon), tip: Some("VRCLogUnrotate"))]
    #[nwg_events(MousePressLeftUp: [SystemTray::show_menu], OnContextMenu: [SystemTray::show_menu])]
    tray: nwg::TrayNotification,

    #[nwg_control(parent: window, popup: true)]
    tray_menu: nwg::Menu,

    #[nwg_control(parent: tray_menu, text: "ログのフォルダを開く")]
    #[nwg_events(OnMenuItemSelected: [SystemTray::open_collection])]
    tray_item_open_collection: nwg::MenuItem,

    #[nwg_control(parent: tray_menu)]
    tray_item_sep1: nwg::MenuSeparator,

    #[nwg_control(parent: tray_menu, text: "終了")]
    #[nwg_events(OnMenuItemSelected: [SystemTray::exit])]
    tray_item_exit: nwg::MenuItem,

    #[nwg_control]
    #[nwg_events( OnNotice: [SystemTray::on_error] )]
    error_notice: nwg::Notice,

    #[nwg_control]
    #[nwg_events( OnNotice: [SystemTray::on_crash] )]
    crash_notice: nwg::Notice,

    error_mpsc: RefCell<Option<mpsc::Receiver<anyhow::Error>>>,
    collection_path: RefCell<Option<PathBuf>>,
}

impl SystemTray {
    fn init(&self) {
        let crash_notifier = CrashNotifier(Some(self.crash_notice.sender()));
        let unrotate = Unrotate::new().unwrap();
        *self.collection_path.borrow_mut() = Some(unrotate.collection.collection_path.clone());
        let (tx, rx) = mpsc::channel();
        *self.error_mpsc.borrow_mut() = Some(rx);
        let error_notifier = self.error_notice.sender();
        thread::spawn(move || {
            let mut crash_notifier = crash_notifier;
            loop {
                if let Err(e) = unrotate.step() {
                    if tx.send(e).is_err() {
                        break;
                    }
                    error_notifier.notice();
                }
                thread::sleep(Duration::from_secs(60));
            }
            crash_notifier.disable();
        });
    }

    fn show_menu(&self) {
        let (x, y) = nwg::GlobalCursor::position();
        self.tray_menu.popup(x, y);
    }

    fn open_collection(&self) {
        if let Some(ref collection_path) = *self.collection_path.borrow() {
            open_explore(&collection_path);
        }
    }

    fn exit(&self) {
        nwg::stop_thread_dispatch();
    }

    fn on_error(&self) {
        if let Some(ref rx) = *self.error_mpsc.borrow() {
            for e in rx.try_iter() {
                let flags = nwg::TrayNotificationFlags::WARNING_ICON
                    | nwg::TrayNotificationFlags::LARGE_ICON;
                self.tray.show(
                    &format!("{}", e),
                    Some("VRCLogUnrotateの動作中にエラーが発生しました"),
                    Some(flags),
                    None,
                );
            }
        }
    }

    fn on_crash(&self) {
        let flags = nwg::TrayNotificationFlags::ERROR_ICON | nwg::TrayNotificationFlags::LARGE_ICON;
        self.tray.show(
            "回復できないエラーが発生したためVRCLogUnrotateを終了します",
            Some("VRCLogUnrotateがクラッシュしました"),
            Some(flags),
            None,
        );
    }
}

fn open_explore(path: &Path) {
    use std::ffi::OsString;
    use std::os::windows::ffi::OsStrExt;
    use std::{iter, ptr};
    #[allow(non_snake_case)]
    let lpOperation: Vec<_> = OsString::from("explore".to_string())
        .encode_wide()
        .chain(iter::once(0))
        .collect();
    #[allow(non_snake_case)]
    let lpFile: Vec<_> = path
        .as_os_str()
        .encode_wide()
        .chain(iter::once(0))
        .collect();
    unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            lpOperation.as_ptr(),
            lpFile.as_ptr(),
            ptr::null(),
            ptr::null(),
            SW_SHOWNORMAL,
        );
    }
}

fn get_appdata_locallow() -> Option<PathBuf> {
    use std::ffi::OsString;
    use std::os::windows::prelude::OsStringExt;
    use std::{ffi, ptr, slice};

    let mut path = None;
    unsafe {
        let mut ppszpath = ptr::null_mut();
        let ret = SHGetKnownFolderPath(
            &FOLDERID_LocalAppDataLow,
            KF_FLAG_DEFAULT,
            ptr::null_mut(),
            &mut ppszpath,
        );
        if ret == 0 {
            let len = (0..).take_while(|&i| *ppszpath.offset(i) != 0).count();
            let wide_slice = slice::from_raw_parts(ppszpath, len);
            let osstring = OsString::from_wide(wide_slice);
            path = Some(PathBuf::from(osstring));
        }
        CoTaskMemFree(ppszpath as *mut ffi::c_void);
    }
    path
}

struct LocalLowVRChat {
    vrchat_path: PathBuf,
}

impl LocalLowVRChat {
    fn new(vrchat_path: PathBuf) -> Self {
        Self { vrchat_path }
    }

    fn from_locallow_path(locallow_path: &Path) -> Self {
        let vrchat_path = locallow_path.join("VRChat").join("VRChat");
        Self::new(vrchat_path)
    }

    fn list_logfile_paths(&self) -> Result<Vec<PathBuf>> {
        lazy_static! {
            static ref RE: Regex = Regex::new("^output_log_\\d{2}-\\d{2}-\\d{2}\\.txt$").unwrap();
        }
        self.vrchat_path
            .read_dir()?
            .filter_map(|dir_entry| {
                dir_entry
                    .map_err(anyhow::Error::new)
                    .and_then(|dir_entry| {
                        dir_entry.file_type().map_err(Into::into).map(|file_type| {
                            if file_type.is_file() {
                                dir_entry.file_name().to_str().and_then(|file_name| {
                                    RE.is_match(file_name).then(|| dir_entry.path())
                                })
                            } else {
                                None
                            }
                        })
                    })
                    .transpose()
            })
            .collect()
    }
}

#[derive(Debug)]
struct VRCLogfile {
    path: PathBuf,
    date: NaiveDate,
}

impl VRCLogfile {
    fn new(path: PathBuf) -> io::Result<Option<Self>> {
        lazy_static! {
            static ref RE: regex::bytes::Regex = regex::bytes::Regex::new("(?m)^(?P<yyyy>\\d{4})\\.(?P<MM>\\d{2})\\.(?P<dd>\\d{2}) (?:\\d{2}):(?:\\d{2}):(?:\\d{2}) ").unwrap();
        }
        let mut file = OpenOptions::new()
            .create(false)
            .write(false)
            .append(false)
            .read(true)
            .open(&path)?;
        let mut head_buf = vec![0u8; 30];
        file.read_exact(&mut head_buf)?;
        let captures = if let Some(captures) = RE.captures(&head_buf) {
            captures
        } else {
            return Ok(None);
        };
        fn parse<T>(bytes: &[u8]) -> T
        where
            T: FromStr,
            T::Err: Debug,
        {
            std::str::from_utf8(bytes)
                .expect("digits")
                .parse()
                .expect("digits")
        }
        let year: i32 = parse(captures.name("yyyy").unwrap().as_bytes());
        let month: u32 = parse(captures.name("MM").unwrap().as_bytes());
        let day: u32 = parse(captures.name("dd").unwrap().as_bytes());
        let date = NaiveDate::from_ymd(year, month, day);
        Ok(Some(Self { path, date }))
    }
}

struct UnrotateCollection {
    collection_path: PathBuf,
}

impl UnrotateCollection {
    fn new(collection_path: PathBuf) -> Self {
        Self { collection_path }
    }

    fn with_locallow_path(locallow_path: &Path) -> Self {
        let collection_path = locallow_path.join(VENDOR_NAME).join(APP_NAME).join("Logs");
        Self::new(collection_path)
    }

    fn partition_folder_path(&self, date: NaiveDate) -> PathBuf {
        self.collection_path
            .join(format!("{:04}-{:02}", date.year(), date.month()))
            .join(format!("{:02}", date.day()))
    }

    fn create_link(&self, logfile: &VRCLogfile) -> io::Result<()> {
        let partition_folder_path = self.partition_folder_path(logfile.date);
        fs::create_dir_all(&partition_folder_path)?;
        let new_link_path = partition_folder_path.join(logfile.path.file_name().unwrap());
        match fs::hard_link(&logfile.path, &new_link_path) {
            Ok(_) => Ok(()),
            Err(e) => match e.kind() {
                io::ErrorKind::AlreadyExists => Ok(()),
                _ => Err(e),
            },
        }
    }
}

struct Unrotate {
    vrchat: LocalLowVRChat,
    collection: UnrotateCollection,
}

impl Unrotate {
    fn step(&self) -> Result<()> {
        for path in self.vrchat.list_logfile_paths()? {
            if let Some(logfile) = VRCLogfile::new(path)? {
                self.collection.create_link(&logfile)?;
            }
        }
        Ok(())
    }

    fn new() -> Result<Self> {
        let locallow = get_appdata_locallow()
            .ok_or_else(|| anyhow::anyhow!("failed to get LocalAppDataLow path"))?;
        let vrchat = LocalLowVRChat::from_locallow_path(&locallow);
        let collection = UnrotateCollection::with_locallow_path(&locallow);
        Ok(Self { vrchat, collection })
    }
}

fn main() {
    nwg::init().expect("Failed to init Native Windows GUI");
    let _ui = SystemTray::build_ui(Default::default()).expect("Failed to build UI");
    nwg::dispatch_thread_events();
}
