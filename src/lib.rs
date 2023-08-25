#[macro_use]
extern crate serde_derive;
mod comglue;
mod server;
use anyhow::{bail, Result};
use com::{
    production::Class,
    sys::{CLASS_E_CLASSNOTAVAILABLE, CLSID, HRESULT, IID, NOERROR, SELFREG_E_CLASS},
};
use comglue::glue::NetidxRTD;
use comglue::interface::CLSID;
use netidx::subscriber::Value;
use std::{
    ffi::{c_char, c_void, CStr},
    mem, ptr,
};

// sadly this doesn't register the class name, just the ID, so we must do all the
// registration ourselves because excel requires the name to be mapped to the id
//com::inproc_dll_module![(CLSID, NetidxRTD),];

static mut _HMODULE: *mut c_void = ptr::null_mut();

mod writer;
use writer::ExcelNetidxWriter;

lazy_static::lazy_static! {
    static ref NETIDXWRITER: anyhow::Result<ExcelNetidxWriter> = ExcelNetidxWriter::new();
    static ref EXCEL_BEGIN_TIME: chrono::NaiveDateTime = chrono::NaiveDate::from_ymd_opt(1899,12,31).expect("this should never happen").and_hms_opt(0, 0, 0).expect("this should never happen");
}

// interface of writing value to netidx container
#[no_mangle]
pub extern "C" fn write_value_string(
    path: *const c_char,
    value: *const c_char,
) -> writer::SendResult {
    match unsafe { CStr::from_ptr(value) }.to_str() {
        Err(_) => writer::SendResult::ExcelErrorNA,
        Ok(value) => write_value(path, Value::String(value.to_string().into())),
    }
}

#[no_mangle]
pub extern "C" fn write_value_i64(path: *const c_char, value: i64) -> writer::SendResult {
    write_value(path, value.into())
}

#[no_mangle]
pub extern "C" fn write_value_f64(path: *const c_char, value: f64) -> writer::SendResult {
    write_value(path, value.into())
}

#[no_mangle]
pub extern "C" fn write_value_bool(
    path: *const c_char,
    value: bool,
) -> writer::SendResult {
    write_value(path, value.into())
}

#[no_mangle]
pub extern "C" fn write_value_timestamp(
    path: *const c_char,
    mut value: f64,
) -> writer::SendResult {
    if value > 59.0 {
        // Excel time starts at 1900/01/01 and assumes Feb 1900 has 29 days by mistake
        value -= 1.0;
    }
    if value < 0.0 {
        return writer::SendResult::ExcelErrorNA;
    }
    let date: chrono::NaiveDateTime =
        *EXCEL_BEGIN_TIME + chrono::Duration::days(value as i64);
    let milliseconds = (value.fract() * 86400.0 * 1000.0) as i64; // convert to milliseconds *24.0 * 60.0 * 60.0 * 1000
    write_value(
        path,
        Value::DateTime(chrono::DateTime::<chrono::Utc>::from_local(
            date + chrono::Duration::milliseconds(milliseconds),
            chrono::Utc,
        )),
    )
}

#[no_mangle]
pub extern "C" fn write_value_error(
    path: *const c_char,
    value: *const c_char,
) -> writer::SendResult {
    match unsafe { CStr::from_ptr(value) }.to_str() {
        Err(_) => writer::SendResult::ExcelErrorNA,
        Ok(value) => write_value(path, Value::Error(value.into())),
    }
}

pub fn write_value(path: *const c_char, value: Value) -> writer::SendResult {
    match unsafe { CStr::from_ptr(path) }.to_str() {
        Err(_) => writer::SendResult::ExcelErrorNA,
        Ok(path) => match NETIDXWRITER.as_ref() {
            Ok(writer) => writer.send(path, value),
            Err(_) => writer::SendResult::ExcelErrorNull,
        },
    }
}

#[test]
fn test_convert_time() {
    let date: chrono::NaiveDateTime = *EXCEL_BEGIN_TIME + chrono::Duration::days(45133);
    let seconds = (0.69032 * 86400.0) as i64; // convert to seconds *24.0 * 60.0 * 60.0
    let datetime = chrono::DateTime::<chrono::Utc>::from_local(
        date + chrono::Duration::seconds(seconds),
        chrono::Utc,
    );
    assert_eq!(&datetime.to_string(), "2023-07-27 16:34:03 UTC");
}

#[no_mangle]
unsafe extern "system" fn DllMain(
    hinstance: *mut c_void,
    fdw_reason: u32,
    _reserved: *mut c_void,
) -> i32 {
    const DLL_PROCESS_ATTACH: u32 = 1;
    if fdw_reason == DLL_PROCESS_ATTACH {
        _HMODULE = hinstance;
    }
    1
}

#[no_mangle]
unsafe extern "system" fn DllGetClassObject(
    class_id: *const CLSID,
    iid: *const IID,
    result: *mut *mut c_void,
) -> HRESULT {
    assert!(
        !class_id.is_null(),
        "class id passed to DllGetClassObject should never be null"
    );

    let class_id = &*class_id;
    if class_id == &CLSID {
        let instance = <NetidxRTD as Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else {
        CLASS_E_CLASSNOTAVAILABLE
    }
}

use winreg::{enums::*, RegKey};

extern "system" {
    fn GetModuleFileNameA(hModule: *mut c_void, lpFilename: *mut i8, nSize: u32) -> u32;
}

unsafe fn get_dll_file_path(hmodule: *mut c_void) -> String {
    const MAX_FILE_PATH_LENGTH: usize = 260;

    let mut path = [0u8; MAX_FILE_PATH_LENGTH];

    let len = GetModuleFileNameA(
        hmodule,
        path.as_mut_ptr() as *mut _,
        MAX_FILE_PATH_LENGTH as _,
    );

    String::from_utf8(path[..len as usize].to_vec()).unwrap()
}

fn clsid(id: CLSID) -> String {
    format!("{{{}}}", id)
}

fn register_clsid(root: &RegKey, clsid: &String) -> Result<()> {
    let (by_id, _) = root.create_subkey(&format!("CLSID\\{}", &clsid))?;
    let (by_id_inproc, _) = by_id.create_subkey("InprocServer32")?;
    by_id.set_value(&"", &"NetidxRTD")?;
    by_id_inproc.set_value("", &unsafe { get_dll_file_path(_HMODULE) })?;
    Ok(())
}

fn dll_register_server() -> Result<()> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let (by_name, _) = hkcr.create_subkey("NetidxRTD\\CLSID")?;
    let clsid = clsid(CLSID);
    by_name.set_value("", &clsid)?;
    if mem::size_of::<usize>() == 8 {
        register_clsid(&hkcr, &clsid)?;
    } else if mem::size_of::<usize>() == 4 {
        let wow = hkcr.open_subkey("WOW6432Node")?;
        register_clsid(&wow, &clsid)?;
    } else {
        bail!("can't figure out the word size")
    }
    Ok(())
}

#[no_mangle]
extern "system" fn DllRegisterServer() -> HRESULT {
    match dll_register_server() {
        Err(_) => SELFREG_E_CLASS,
        Ok(()) => NOERROR,
    }
}

fn dll_unregister_server() -> Result<()> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let clsid = clsid(CLSID);
    hkcr.delete_subkey_all("NetidxRTD")?;
    assert!(clsid.len() > 0);
    hkcr.delete_subkey_all(&format!("CLSID\\{}", clsid))?;
    hkcr.delete_subkey_all(&format!("WOW6432Node\\CLSID\\{}", clsid))?;
    Ok(())
}

#[no_mangle]
extern "system" fn DllUnregisterServer() -> HRESULT {
    match dll_unregister_server() {
        Err(_) => SELFREG_E_CLASS,
        Ok(()) => NOERROR,
    }
}
