use std::mem::ManuallyDrop;

use windows::{
    Win32::{
        Foundation::VARIANT_BOOL,
        System::Com::{DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS},
        System::Variant::{
            VAR_CHANGE_FLAGS, VARIANT_0_0, VT_BOOL, VT_BSTR, VT_I2, VT_I4, VT_I8, VT_NULL, VT_UI1,
            VT_UI2, VT_UI4, VT_UI8, VariantChangeType, VariantClear,
        },
    },
    core::{self, BSTR},
};

pub use com_shim_macro::com_shim;

pub use windows::Win32::System::{Com::IDispatch, Variant::VARIANT};
pub use windows::core::{GUID, Result};

mod utils;

/// A component that has an IDispatch value. Every component needs this, and this trait guarantees that.
pub trait HasIDispatch<T = Self> {
    /// Get the IDispatch object for low-level access to this component.
    fn get_idispatch(&self) -> &IDispatch;
}

pub trait IDispatchExt {
    /// Call a function on this IDispatch
    fn call<S>(&self, name: S, args: Vec<VARIANT>) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Get the value of a variable on this IDispatch
    fn get<S>(&self, name: S) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Set a value of a variable on this IDispatch
    fn set<S>(&self, name: S, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>;
}

impl IDispatchExt for IDispatch {
    fn call<S>(&self, name: S, args: Vec<VARIANT>) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            log::debug!("Invoking method: {}", name.as_ref());
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_METHOD,
                &utils::assemble_dispparams_get(args),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn get<S>(&self, name: S) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYGET,
                &DISPPARAMS::default(),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn set<S>(&self, name: S, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYPUT,
                &utils::assemble_dispparams_put(vec![value]),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }
}

pub trait VariantExt {
    fn null() -> VARIANT;
    fn from_i16(n: i16) -> VARIANT;
    fn from_i32(n: i32) -> VARIANT;
    fn from_i64(n: i64) -> VARIANT;
    fn from_u8(n: u8) -> VARIANT;
    fn from_u16(n: u16) -> VARIANT;
    fn from_u32(n: u32) -> VARIANT;
    fn from_u64(n: u64) -> VARIANT;
    fn from_str<S: AsRef<str>>(s: S) -> VARIANT;
    fn from_bool(b: bool) -> VARIANT;
    fn to_i16(&self) -> core::Result<i16>;
    fn to_i32(&self) -> core::Result<i32>;
    fn to_i64(&self) -> core::Result<i64>;
    fn to_u8(&self) -> core::Result<u8>;
    fn to_u16(&self) -> core::Result<u16>;
    fn to_u32(&self) -> core::Result<u32>;
    fn to_u64(&self) -> core::Result<u64>;
    fn to_string(&self) -> core::Result<String>;
    fn to_bool(&self) -> core::Result<bool>;
    fn to_idispatch(&self) -> core::Result<&IDispatch>;
}

impl VariantExt for VARIANT {
    fn null() -> VARIANT {
        let mut variant = VARIANT::default();
        let v00 = VARIANT_0_0 {
            vt: VT_NULL,
            ..Default::default()
        };
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i16(n: i16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I2,
            ..Default::default()
        };
        v00.Anonymous.iVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i32(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I4,
            ..Default::default()
        };
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i64(n: i64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I8,
            ..Default::default()
        };
        v00.Anonymous.llVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u8(n: u8) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI1,
            ..Default::default()
        };
        v00.Anonymous.bVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u16(n: u16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI2,
            ..Default::default()
        };
        v00.Anonymous.uiVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u32(n: u32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI4,
            ..Default::default()
        };
        v00.Anonymous.ulVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u64(n: u64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI8,
            ..Default::default()
        };
        v00.Anonymous.ullVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_str<S: AsRef<str>>(s: S) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BSTR,
            ..Default::default()
        };
        let bstr = BSTR::from(s.as_ref());
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_bool(b: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BOOL,
            ..Default::default()
        };
        v00.Anonymous.boolVal = VARIANT_BOOL::from(b);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn to_i16(&self) -> core::Result<i16> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.iVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_i32(&self) -> core::Result<i32> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.lVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_i64(&self) -> core::Result<i64> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.llVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u8(&self) -> core::Result<u8> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI1)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.bVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u16(&self) -> core::Result<u16> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.uiVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u32(&self) -> core::Result<u32> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ulVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u64(&self) -> core::Result<u64> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ullVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_string(&self) -> core::Result<String> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BSTR)?;
            let v00 = &new.Anonymous.Anonymous;
            let str = v00.Anonymous.bstrVal.to_string();
            VariantClear(&mut new)?;
            Ok(str)
        }
    }
    fn to_bool(&self) -> core::Result<bool> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BOOL)?;
            let v00 = &new.Anonymous.Anonymous;
            let b = v00.Anonymous.boolVal.as_bool();
            VariantClear(&mut new)?;
            Ok(b)
        }
    }
    fn to_idispatch(&self) -> core::Result<&IDispatch> {
        unsafe {
            log::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let v00 = &self.Anonymous.Anonymous;
            let idisp = v00.Anonymous.pdispVal.as_ref().ok_or(core::Error::new(
                core::HRESULT(0x00123456),
                core::HSTRING::from("com-shim: Cannot read IDispatch"),
            ))?;
            Ok(idisp)
        }
    }
}
