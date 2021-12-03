use crate::{
    comglue::interface::{
        IDispatch, IRTDServer, IRTDUpdateEvent, IID_IDISPATCH, IID_IRTDSERVER,
        IID_IRTDUPDATE_EVENT,
    },
    server::{Server, TopicId},
};
use com::{
    interfaces::IUnknown,
    sys::{HRESULT, IID, NOERROR},
};
use log::{debug, LevelFilter};
use oaidl::{SafeArrayExt, VariantExt, VtNull};
use once_cell::sync::Lazy;
use simplelog;
use std::{
    ffi::OsString,
    fs::File,
    marker::{Send, Sync},
    os::windows::ffi::{OsStrExt, OsStringExt},
    ptr,
    cell::RefCell,
};
use winapi::{
    shared::{
        guiddef::GUID,
        minwindef::{UINT, WORD},
        wtypes,
        wtypesbase::LPOLESTR,
    },
    um::{
        self,
        oaidl::{ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT},
        oleauto::DISPATCH_METHOD,
        winbase::lstrlenW,
        winnt::LCID,
    },
};

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Debug, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

fn maybe_init_logger() {
    *LOGGER
}

unsafe fn string_from_wstr<'a>(s: *const u16) -> OsString {
    OsString::from_wide(std::slice::from_raw_parts(s, lstrlenW(s) as usize))
}

fn str_to_wstr(s: &str) -> Vec<u16> {
    let mut v = OsString::from(s).encode_wide().collect::<Vec<_>>();
    v.push(0);
    v
}

// Excel hands us an IRTDUpdateEvent class that we need to use to tell it that we have data,
// however it doesn't give us an IRTDUpdateEvent COM interface, it gives us an IDispatch COM
// interface, so we need to use that to call the methods of IRTDUpdateEvent through IDispatch.

struct DispIds {
    update_notify_id: DISPID,
    heartbeat_interval_id: DISPID,
    disconnect_id: DISPID,
}

pub(crate) struct IRTDUpdateEventWrap {
    ptr: *mut um::oaidl::IDispatch,
    iid: GUID,
    dispids: Option<DispIds>,
}

// CR estokes: verify that this is ok. Somehow ...
unsafe impl Send for IRTDUpdateEventWrap {}
unsafe impl Sync for IRTDUpdateEventWrap {}

impl IRTDUpdateEventWrap {
    fn get_dispids(&mut self) {
        let mut update_notify = str_to_wstr("UpdateNotify");
        let mut heartbeat_interval = str_to_wstr("HeartbeatInterval");
        let mut disconnect = str_to_wstr("Disconnect");
        let mut names = [
            update_notify.as_mut_ptr(),
            heartbeat_interval.as_mut_ptr(),
            disconnect.as_mut_ptr(),
        ];
        let mut dispids: [DISPID; 3] = [0x0, 0x0, 0x0];
        let res = unsafe {
            (*self.ptr).GetIDsOfNames(&self.iid, names.as_mut_ptr(), 3, 0, dispids.as_mut_ptr())
        };
        if res != NOERROR {
            panic!("IRTDUpdateEventWrap: could not get names {}", res);
        }
        self.dispids = Some(DispIds {
            update_notify_id: dispids[0],
            heartbeat_interval_id: dispids[1],
            disconnect_id: dispids[2],
        });
    }

    fn new(ptr: *mut um::oaidl::IDispatch) -> Self {
        assert!(!ptr.is_null());
        let iid = GUID {
            Data1: IID_IRTDUPDATE_EVENT.data1,
            Data2: IID_IRTDUPDATE_EVENT.data2,
            Data3: IID_IRTDUPDATE_EVENT.data3,
            Data4: IID_IRTDUPDATE_EVENT.data4,
        };
        IRTDUpdateEventWrap {
            ptr,
            iid,
            dispids: None,
        }
    }

    pub(crate) fn update_notify(&mut self) {
        if self.dispids.is_none() {
            self.get_dispids();
        }
        let update_notify_id = self.dispids.as_ref().unwrap().update_notify_id;
        let mut args = [];
        let mut named_args = [];
        let mut params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: named_args.as_mut_ptr(),
            cArgs: 0,
            cNamedArgs: 0,
        };
        let mut _result =
            VariantExt::into_variant(VtNull).expect("couldn't create result variant");
        let mut _arg_err = 0;
        let res = unsafe {
            (*self.ptr).Invoke(
                update_notify_id,
                &self.iid,
                0,
                DISPATCH_METHOD,
                &mut params,
                _result.as_ptr(),
                ptr::null_mut(),
                &mut _arg_err,
            )
        };
        if res != NOERROR {
            panic!("IRTDUpdateEvent: update_notify failed {}", res);
        }
    }
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {
        server: RefCell<Option<Server>>,
    }

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut UINT) -> HRESULT {
            maybe_init_logger();
            debug!("get_type_info_count(info: {})", unsafe { *info });
            if !info.is_null() {
                unsafe { *info = 0; } // no we don't support type info
            }
            NOERROR
        }

        fn get_type_info(&self, _lcid: LCID, _type_info: *mut *mut ITypeInfo) -> HRESULT { NOERROR }

        pub fn get_ids_of_names(
            &self,
            riid: *const IID,
            names: *const LPOLESTR,
            names_len: UINT,
            lcid: LCID,
            ids: *mut DISPID
        ) -> HRESULT {
            maybe_init_logger();
            debug!("get_ids_of_names(riid: {:?}, names: {:?}, names_len: {}, lcid: {}, ids: {:?})", riid, names, names_len, lcid, ids);
            if !ids.is_null() && !names.is_null() {
                for i in 0..names_len {
                    let name = unsafe { string_from_wstr(*names.offset(i as isize)) };
                    let name = name.to_string_lossy();
                    debug!("name: {}", name);
                    match &*name {
                        "ServerStart" => unsafe { *ids.offset(i as isize) = 0; },
                        "ServerTerminate" => unsafe { *ids.offset(i as isize) = 1; }
                        "ConnectData" => unsafe { *ids.offset(i as isize) = 2; }
                        "RefreshData" => unsafe { *ids.offset(i as isize) = 3; }
                        "DisconnectData" => unsafe { *ids.offset(i as isize) = 4; }
                        "Heartbeat" => unsafe { *ids.offset(i as isize) = 5; }
                        _ => debug!("unknown method: {}", name)
                    }
                }
            }
            NOERROR
        }

        fn invoke(
            &self,
            id: DISPID,
            iid: *const IID,
            lcid: LCID,
            flags: WORD,
            params: *mut DISPPARAMS,
            result: *mut VARIANT,
            exception: *mut EXCEPINFO,
            arg_error: *mut UINT
        ) -> HRESULT {
            maybe_init_logger();
            debug!(
                "invoke(id: {}, iid: {:?}, lcid: {}, flags: {}, params: {:?}, result: {:?}, exception: {:?}, arg_error: {:?})",
                id, iid, lcid, flags, params, result, exception, arg_error
            );
            assert!(!params.is_null());
            match id {
                0 => {
                    debug!("ServerStart called");
                    if self.server.borrow().is_none() {
                        debug!("starting netidx rtd server");
                        let updates = unsafe { (*(*params).rgvarg).n1.n2().n3.pdispVal() };
                        let updates = IRTDUpdateEventWrap::new(*updates);
                        *self.server.borrow_mut() = Some(Server::new(updates).expect("failed to start server"));
                        debug!("server started");
                        unsafe { *(*result).n1.n2_mut().n3.lVal_mut() = 1; }
                    }
                },
                1 => {
                    debug!("ServerTerminate called")
                },
                2 => {
                    debug!("ConnectData called")
                },
                3 => {
                    debug!("RefreshData called")
                },
                4 => {
                    debug!("DisconnectData called")
                },
                5 => {
                    debug!("Heartbeat called")
                },
                _ => {
                    debug!("unknown method {} called", id)
                },
            }
            NOERROR
        }
    }

    impl IRTDServer for NetidxRTD {
        fn server_start(&self, _cb: *const IRTDUpdateEvent, _res: *mut i32) -> HRESULT {
            NOERROR
        }

        fn connect_data(&self, _topic_id: i32, _topic: *const SAFEARRAY, _get_new_values: *mut VARIANT, _res: *mut VARIANT) -> HRESULT {
            NOERROR
        }

        fn refresh_data(&self, _topic_count: *mut i32, _data: *mut SAFEARRAY) -> HRESULT {
            NOERROR
        }

        fn disconnect_data(&self, _topic_id: i32) -> HRESULT {
            NOERROR
        }

        fn heartbeat(&self, _res: *mut i32) -> HRESULT {
            NOERROR
        }

        fn server_terminate(&self) -> HRESULT {
            NOERROR
        }
    }
}
