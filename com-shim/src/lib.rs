#![warn(missing_docs)]
#![warn(clippy::pedantic)]
#![doc = include_str!("../README.md")]

use std::mem::ManuallyDrop;

use windows::{
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::{DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS},
            Variant::{
                VAR_CHANGE_FLAGS, VARIANT_0_0, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I2, VT_I4, VT_I8,
                VT_NULL, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VariantChangeType, VariantClear,
            },
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

/// Additional functions for working with an [`IDispatch`].
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
            tracing::debug!("Invoking method: {}", name.as_ref());
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

/// Extension functions for a [`VARIANT`].
pub trait VariantExt {
    /// Generate a null [`VARIANT`].
    fn null() -> VARIANT;
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
}

/// Indicates that this type is also a parent type and can be upcast to it.
pub trait IsA<T> {
    /// Upcast this value to it's parent type.
    fn upcast(&self) -> T;
}

/// Functions to convert to and from a type that can be stored in a [`VARIANT`].
pub trait VariantTypeExt<'a, T> {
    /// Convert from a [`VARIANT`] into a type, T.
    fn variant_into(&'a self) -> core::Result<T>;

    /// Convert from a type T into a [`VARIANT`].
    fn variant_from(value: T) -> VARIANT;
}

impl VariantTypeExt<'_, ()> for VARIANT {
    fn variant_from(_value: ()) -> VARIANT {
        VARIANT::null()
    }

    fn variant_into(&'_ self) -> core::Result<()> {
        Ok(())
    }
}

impl VariantTypeExt<'_, i16> for VARIANT {
    fn variant_from(value: i16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I2,
            ..Default::default()
        };
        v00.Anonymous.iVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<i16> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.iVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, i32> for VARIANT {
    fn variant_from(value: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I4,
            ..Default::default()
        };
        v00.Anonymous.lVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<i32> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.lVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, i64> for VARIANT {
    fn variant_from(value: i64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I8,
            ..Default::default()
        };
        v00.Anonymous.llVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<i64> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.llVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, u8> for VARIANT {
    fn variant_from(value: u8) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI1,
            ..Default::default()
        };
        v00.Anonymous.bVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<u8> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI1)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.bVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, u16> for VARIANT {
    fn variant_from(value: u16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI2,
            ..Default::default()
        };
        v00.Anonymous.uiVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<u16> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.uiVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, u32> for VARIANT {
    fn variant_from(value: u32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI4,
            ..Default::default()
        };
        v00.Anonymous.ulVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<u32> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ulVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, u64> for VARIANT {
    fn variant_from(value: u64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI8,
            ..Default::default()
        };
        v00.Anonymous.ullVal = value;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<u64> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ullVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
}

impl VariantTypeExt<'_, String> for VARIANT {
    fn variant_from(value: String) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BSTR,
            ..Default::default()
        };
        let bstr = BSTR::from(&value);
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<String> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BSTR)?;
            let v00 = &new.Anonymous.Anonymous;
            let str = v00.Anonymous.bstrVal.to_string();
            VariantClear(&mut new)?;
            Ok(str)
        }
    }
}

impl VariantTypeExt<'_, bool> for VARIANT {
    fn variant_from(value: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BOOL,
            ..Default::default()
        };
        v00.Anonymous.boolVal = VARIANT_BOOL::from(value);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&self) -> core::Result<bool> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BOOL)?;
            let v00 = &new.Anonymous.Anonymous;
            let b = v00.Anonymous.boolVal.as_bool();
            VariantClear(&mut new)?;
            Ok(b)
        }
    }
}

impl<'a> VariantTypeExt<'a, &'a IDispatch> for VARIANT {
    fn variant_from(value: &'a IDispatch) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_DISPATCH,
            ..Default::default()
        };
        v00.Anonymous.pdispVal = ManuallyDrop::new(Some(value.clone()));
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn variant_into(&'a self) -> core::Result<&'a IDispatch> {
        unsafe {
            tracing::debug!("Own type: {:?}", self.Anonymous.Anonymous.vt);
            let v00 = &self.Anonymous.Anonymous;
            let idisp = v00.Anonymous.pdispVal.as_ref().ok_or(core::Error::new(
                core::HRESULT(0x00123456),
                core::HSTRING::from("com-shim: Cannot read IDispatch"),
            ))?;
            Ok(idisp)
        }
    }
}
