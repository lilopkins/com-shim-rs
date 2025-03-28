#![deny(unsafe_code)]
#![warn(clippy::pedantic)]
#![warn(missing_docs)]
#![doc = include_str!("../README.md")]

use heck::ToSnakeCase;
use proc_macro::TokenStream;
use quote::{ToTokens, TokenStreamExt, quote};
use syn::{
    braced, ext::IdentExt, parenthesized, parse::Parse, parse_macro_input, punctuated::Punctuated, Attribute, Ident, Token
};

struct Class {
    attributes: Vec<Attribute>,
    ident: Ident,
    inherited: Vec<Ident>,
    functions_and_variables: Punctuated<FunctionOrVariable, Token![,]>,
}

impl Parse for Class {
    fn parse(input: syn::parse::ParseStream) -> syn::Result<Self> {
        let attributes = Attribute::parse_outer(input)?;
        let _: Token![struct] = input.parse()?;
        let ident: Ident = input.parse()?;
        let mut inherited: Vec<Ident> = vec![];
        if input.peek(Token![:]) {
            // Parse inheritance
            let _: Token![:] = input.parse()?;
            loop {
                inherited.push(input.parse()?);
                if input.peek(Token![+]) {
                    let _: Token![+] = input.parse()?;
                } else {
                    break;
                }
            }
        }
        let content;
        braced!(content in input);
        let functions_and_variables =
            content.parse_terminated(FunctionOrVariable::parse, Token![,])?;

        Ok(Self {
            attributes,
            ident,
            inherited,
            functions_and_variables,
        })
    }
}

enum FunctionOrVariable {
    Function(Function),
    Variable(Variable),
}

impl ToTokens for FunctionOrVariable {
    fn to_tokens(&self, tokens: &mut proc_macro2::TokenStream) {
        match self {
            Self::Function(f) => f.to_tokens(tokens),
            Self::Variable(v) => v.to_tokens(tokens),
        }
    }
}

impl Parse for FunctionOrVariable {
    fn parse(input: syn::parse::ParseStream) -> syn::Result<Self> {
        let attributes = Attribute::parse_outer(input)?;
        if input.peek(Token![fn]) {
            // Parse function next
            let _: Token![fn] = input.parse()?;
            let ident: Ident = input.parse()?;
            let parameters_raw;
            parenthesized!(parameters_raw in input);
            let parameters = parameters_raw.parse_terminated(Ident::parse, Token![,])?;
            let returns = if input.peek(Token![->]) {
                let _: Token![->] = input.parse()?;
                Some(input.parse::<Ident>()?)
            } else {
                None
            };
            Ok(FunctionOrVariable::Function(Function {
                attributes,
                ident,
                parameters,
                returns,
            }))
        } else if input.peek(Token![mut]) {
            // Parse read/write variable
            let _: Token![mut] = input.parse()?;
            let ident: Ident = input.parse()?;
            let _: Token![:] = input.parse()?;
            let type_: Ident = input.parse()?;
            Ok(FunctionOrVariable::Variable(Variable {
                attributes,
                mutable: true,
                ident,
                type_,
            }))
        } else {
            // Parse read-only variable
            let ident: Ident = input.parse()?;
            let _: Token![:] = input.parse()?;
            let type_: Ident = input.parse()?;
            Ok(FunctionOrVariable::Variable(Variable {
                attributes,
                mutable: false,
                ident,
                type_,
            }))
        }
    }
}

struct Variable {
    attributes: Vec<Attribute>,
    mutable: bool,
    ident: Ident,
    type_: Ident,
}

impl ToTokens for Variable {
    fn to_tokens(&self, tokens: &mut proc_macro2::TokenStream) {
        let Variable {
            attributes,
            mutable,
            ident,
            type_,
        } = self;
        let ident_str = ident.to_string();
        let ident_unraw_str = ident.unraw().to_string();

        let read_ident = Ident::new(&ident_str.to_snake_case(), ident.span());
        tokens.append_all(quote! {
            #(#attributes)*
            fn #read_ident(&self) -> ::com_shim::Result<#type_> {
                use ::com_shim::{IDispatchExt, VariantTypeExt};
                ::std::result::Result::Ok(self.get_idispatch().get(#ident_unraw_str)?.variant_into()?)
            }
        });

        if *mutable {
            let write_ident =
                Ident::new(&format!("set_{}", ident_str.to_snake_case()), ident.span());
            tokens.append_all(quote! {
                #(#attributes)*
                fn #write_ident(&self, value: #type_) -> ::com_shim::Result<()> {
                    use ::com_shim::{IDispatchExt, VariantTypeExt};
                    let _ = self.get_idispatch().set(#ident_unraw_str, ::com_shim::VARIANT::variant_from(value))?;
                    ::std::result::Result::Ok(())
                }
            });
        }
    }
}

struct Function {
    attributes: Vec<Attribute>,
    ident: Ident,
    parameters: Punctuated<Ident, Token![,]>,
    returns: Option<Ident>,
}

impl ToTokens for Function {
    fn to_tokens(&self, tokens: &mut proc_macro2::TokenStream) {
        let Function {
            attributes,
            ident,
            parameters,
            returns,
        } = self;
        let ident_str = ident.to_string();
        let ident_unraw_str = ident.unraw().to_string();
        let fn_ident = Ident::new(&ident_str.to_snake_case(), ident.span());
        let fn_parameters = parameters.iter().enumerate().map(|(idx, p)| {
            let ident = Ident::new(&format!("p{idx}"), p.span());
            quote!(#ident: #p)
        });
        let parameters = parameters.iter().enumerate().map(|(idx, p)| {
            let ident = Ident::new(&format!("p{idx}"), p.span());
            quote!(::com_shim::VARIANT::variant_from(#ident))
        });
        let (returns_type, return_statement) = if let Some(returns) = returns {
            (quote!(#returns), quote!(r.variant_into()?))
        } else {
            (quote!(()), quote!(()))
        };
        tokens.append_all(quote! {
            #(#attributes)*
            fn #fn_ident(&self, #(#fn_parameters),*) -> ::com_shim::Result<#returns_type> {
                use ::com_shim::{IDispatchExt, VariantTypeExt};
                let r = self.get_idispatch().call(#ident_unraw_str, vec![
                    #(#parameters),*
                ])?;
                ::std::result::Result::Ok(#return_statement)
            }
        });
    }
}

/// Generate a COM-compatible class structure.
#[proc_macro]
pub fn com_shim(stream: TokenStream) -> TokenStream {
    let Class {
        attributes,
        ident,
        inherited,
        functions_and_variables,
    } = parse_macro_input!(stream as Class);

    let functions_and_variables = functions_and_variables.into_iter();
    let self_impl = Ident::new(&format!("{ident}Ext"), ident.span());
    let inherited_casts = inherited.iter().map(|i| quote! {
        impl ::com_shim::IsA<#i> for #ident {
            fn upcast(&self) -> #i {
                #i::from(self.inner.clone())
            }
        }
    });
    let inherited_impls = inherited
        .iter()
        .map(|i| Ident::new(&format!("{i}Ext"), i.span()));
    quote! {
        #(#attributes)*
        pub struct #ident {
            inner: ::com_shim::IDispatch,
        }

        impl ::com_shim::HasIDispatch for #ident {
            fn get_idispatch(&self) -> &::com_shim::IDispatch {
                &self.inner
            }
        }

        pub trait #self_impl<T: ::com_shim::HasIDispatch = Self>: ::com_shim::HasIDispatch<T> {
            #(#functions_and_variables)*
        }

        impl #self_impl for #ident {}

        #(impl #inherited_impls for #ident {})*

        #(#inherited_casts)*

        impl ::std::convert::From<::com_shim::IDispatch> for #ident {
            fn from(value: ::com_shim::IDispatch) -> Self {
                Self { inner: value }
            }
        }

        impl ::com_shim::VariantTypeExt<'_, #ident> for ::com_shim::VARIANT {
            fn variant_from(value: #ident) -> ::com_shim::VARIANT {
                let idisp = &value.inner;
                ::com_shim::VARIANT::variant_from(idisp)
            }

            fn variant_into(&'_ self) -> ::com_shim::Result<#ident> {
                let idisp = <::com_shim::VARIANT as ::com_shim::VariantTypeExt<'_, &::com_shim::IDispatch>>
                    ::variant_into(self)?.clone();
                ::std::result::Result::Ok(#ident::from(idisp))
            }
        }
    }.into()
}
