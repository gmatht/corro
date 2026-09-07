#![allow(non_snake_case, non_upper_case_globals, non_camel_case_types, dead_code, clippy::all)]

// (Win95 patch, vendored) GetTimeZoneInformationForYear is a Vista+ API;
// statically importing it (windows_link::link!) would make the Win95 loader
// reject the whole exe. Instead resolve it dynamically and fall back to
// GetTimeZoneInformation (Win95-era), which returns the CURRENT year's rules:
// on Win95/98/ME DST-rule changes across years are ignored and the dynamic
// timezone argument (pdtzi, also Vista+) is unsupported. Modern Windows keeps
// exact behavior. The function keeps the same signature and BOOL semantics
// (0 = failure) as the API so the caller is unchanged.
#[allow(non_snake_case)]
pub unsafe fn GetTimeZoneInformationForYear(
    wyear: u16,
    pdtzi: *const DYNAMIC_TIME_ZONE_INFORMATION,
    ptzi: *mut TIME_ZONE_INFORMATION,
) -> BOOL {
    let _ = wyear;
    let _ = pdtzi;
    unsafe extern "system" {
        fn GetModuleHandleA(name: *const u8) -> isize;
        fn GetProcAddress(module: isize, name: *const u8) -> isize; // FARPROC
    }
    unsafe {
        let k32 = GetModuleHandleA(b"kernel32\0".as_ptr());
        if k32 != 0 {
            let p = GetProcAddress(k32, b"GetTimeZoneInformationForYear\0".as_ptr());
            if p != 0 {
                let f: unsafe extern "system" fn(
                    u16,
                    *const DYNAMIC_TIME_ZONE_INFORMATION,
                    *mut TIME_ZONE_INFORMATION,
                ) -> BOOL = std::mem::transmute(p);
                return f(wyear, pdtzi, ptzi);
            }
        }
        unsafe extern "system" {
            fn GetTimeZoneInformation(ptzi: *mut TIME_ZONE_INFORMATION) -> u32;
        }
        // TIME_ZONE_ID_UNKNOWN (0), STANDARD (1) and DAYLIGHT (2) all fill the
        // struct; only 0xFFFFFFFF (undefined error) is a failure.
        let r = GetTimeZoneInformation(ptzi);
        if r == u32::MAX {
            0
        } else {
            1
        }
    }
}
windows_link::link!("kernel32.dll" "system" fn SystemTimeToFileTime(lpsystemtime : *const SYSTEMTIME, lpfiletime : *mut FILETIME) -> BOOL);
windows_link::link!("kernel32.dll" "system" fn SystemTimeToTzSpecificLocalTime(lptimezoneinformation : *const TIME_ZONE_INFORMATION, lpuniversaltime : *const SYSTEMTIME, lplocaltime : *mut SYSTEMTIME) -> BOOL);
windows_link::link!("kernel32.dll" "system" fn TzSpecificLocalTimeToSystemTime(lptimezoneinformation : *const TIME_ZONE_INFORMATION, lplocaltime : *const SYSTEMTIME, lpuniversaltime : *mut SYSTEMTIME) -> BOOL);
pub type BOOL = i32;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DYNAMIC_TIME_ZONE_INFORMATION {
    pub Bias: i32,
    pub StandardName: [u16; 32],
    pub StandardDate: SYSTEMTIME,
    pub StandardBias: i32,
    pub DaylightName: [u16; 32],
    pub DaylightDate: SYSTEMTIME,
    pub DaylightBias: i32,
    pub TimeZoneKeyName: [u16; 128],
    pub DynamicDaylightTimeDisabled: bool,
}
impl Default for DYNAMIC_TIME_ZONE_INFORMATION {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct FILETIME {
    pub dwLowDateTime: u32,
    pub dwHighDateTime: u32,
}
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct SYSTEMTIME {
    pub wYear: u16,
    pub wMonth: u16,
    pub wDayOfWeek: u16,
    pub wDay: u16,
    pub wHour: u16,
    pub wMinute: u16,
    pub wSecond: u16,
    pub wMilliseconds: u16,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub struct TIME_ZONE_INFORMATION {
    pub Bias: i32,
    pub StandardName: [u16; 32],
    pub StandardDate: SYSTEMTIME,
    pub StandardBias: i32,
    pub DaylightName: [u16; 32],
    pub DaylightDate: SYSTEMTIME,
    pub DaylightBias: i32,
}
impl Default for TIME_ZONE_INFORMATION {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
