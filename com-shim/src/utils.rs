use windows::{
    Win32::System::{Com::*, Ole::*, Variant::*},
    core::*,
};

pub(crate) fn get_method_dispid<S>(disp: &IDispatch, name: S) -> Result<i32>
where
    S: AsRef<str>,
{
    unsafe {
        let riid = GUID::zeroed();
        let hstring = HSTRING::from(name.as_ref());
        let rgsznames = PCWSTR::from_raw(hstring.as_ptr());
        let cnames = 1;
        let lcid = 0x09; // en
        let mut dispidmember = 0;

        disp.GetIDsOfNames(&riid, &rgsznames, cnames, lcid, &mut dispidmember)?;
        Ok(dispidmember)
    }
}

pub(crate) fn assemble_dispparams_get(args: &mut Vec<VARIANT>) -> DISPPARAMS {
    args.reverse(); // https://stackoverflow.com/a/65255739
    DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        cArgs: args.len() as u32,
        ..Default::default()
    }
}

static PUT_NAMED_ARGS: [i32; 1] = [DISPID_PROPERTYPUT];

pub(crate) fn assemble_dispparams_put(args: &mut Vec<VARIANT>) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        cArgs: args.len() as u32,
        cNamedArgs: PUT_NAMED_ARGS.len() as u32,
        rgdispidNamedArgs: PUT_NAMED_ARGS.as_ptr().cast_mut(),
    }
}
