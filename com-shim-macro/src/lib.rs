use std::fmt;

use proc_macro::{self, Delimiter, Group, Ident, Span, TokenStream, TokenTree};

#[proc_macro]
pub fn com_shim(stream: TokenStream) -> TokenStream {
    let mut result_stream = TokenStream::new();
    let mut stream_iter = stream.into_iter().peekable();

    // Class token
    let class_token = stream_iter.next().expect("Syntax error: expected `class`");
    match class_token {
        TokenTree::Ident(id) => {
            if id.to_string() != String::from("class") {
                panic!("Syntax error: expected `class`");
            }
            result_stream.extend(vec![TokenTree::Ident(Ident::new(
                "struct",
                Span::call_site(),
            ))]);
        }
        _ => panic!("Syntax error: expect `class` ident"),
    }

    // Name
    let name_token = stream_iter.next().expect("Syntax error: expected name");
    match &name_token {
        proc_macro::TokenTree::Ident(id) => {
            result_stream.extend(vec![TokenTree::Ident(id.clone())])
        }
        _ => panic!("Syntax error: expect name identifier"),
    }
    let name = name_token.to_string();

    // Push struct group
    result_stream.extend(
        "{ inner: ::com_shim::IDispatch }"
            .parse::<TokenStream>()
            .unwrap(),
    );

    // Inherit HasIDispatch trait
    result_stream.extend(format!("impl ::com_shim::HasIDispatch for {name} {{ fn get_idispatch(&self) -> &::com_shim::IDispatch {{ &self.inner }} }}").parse::<TokenStream>().unwrap());

    if stream_iter
        .peek()
        .expect("Syntax error: expected `:` or start of class")
        .to_string()
        == String::from(":")
    {
        // TODO Process parent class.
        loop {
            let _separator_token = stream_iter.next().unwrap();
            let parent_token = stream_iter
                .next()
                .expect("Syntax Error: expected identifier of parent class after `:`");
            match parent_token {
                TokenTree::Ident(id) => {
                    result_stream.extend(
                        format!("impl {}_Impl for {name} {{}}", id.to_string())
                            .parse::<TokenStream>()
                            .unwrap(),
                    );
                }
                _ => panic!("Syntax error: expected identifier for parent"),
            }

            // peek next
            let next = stream_iter
                .peek()
                .expect("Syntax error: expected `+` or start of class");
            match next {
                TokenTree::Group(_) => break,
                TokenTree::Punct(p) => {
                    if p.to_string() != String::from("+") {
                        panic!("Syntax Error: expected identifier of parent class after `:`");
                    }
                }
                _ => panic!("Syntax Error: expected identifier of parent class after `:`"),
            }
        }
    }

    result_stream.extend(
        format!("impl {name}_Impl for {name} {{}}")
            .parse::<TokenStream>()
            .unwrap(),
    );

    // create trait for this impl
    //  { ... }
    result_stream.extend(format!("pub trait {name}_Impl<T: ::com_shim::HasIDispatch = Self>: ::com_shim::HasIDispatch<T>").parse::<TokenStream>().unwrap());

    let mut trait_body_stream: Vec<TokenTree> = Vec::new();
    match stream_iter
        .next()
        .expect("Syntax error: expected `{ ... }`")
    {
        TokenTree::Group(group) => {
            if group.delimiter() != Delimiter::Brace {
                panic!("Syntax error: expected `{{ ... }}`");
            }
            // parse group members
            let items = divide_items(group.stream());
            for item in items {
                build_item_token_strem(item, &mut trait_body_stream);
            }
        }
        _ => panic!("Syntax error: expected `{{ ... }}`"),
    }
    result_stream.extend(vec![TokenTree::Group(Group::new(
        Delimiter::Brace,
        TokenStream::from_iter(trait_body_stream),
    ))]);

    if stream_iter.next().is_some() {
        panic!("Syntax error: expected end of shim definition.");
    }

    result_stream
}

#[derive(Debug)]
enum DivideItemState {
    ExpectFnVarNameOrMut, // fn or NameOfVariable or mut
    ExpectVarName,        // NameOfVariable
    ExpectVarSep,         // :
    ExpectVarType,        // i32
    ExpectSep,            // ,
    ExpectFnName,         // function_name
    ExpectFnParams,       // (...params...)
    ExpectFnArrow1OrSep,  // - or ,
    ExpectFnArrow2,       // >
    ExpectFnReturnType,   // String, i32
}

#[derive(Debug)]
enum ChildItem {
    Invalid,
    Variable {
        mutable: bool,
        name: Option<String>,
        typ: Option<ReturnType>,
    },
    Function {
        name: Option<String>,
        params: Option<Group>,
        return_typ: Option<ReturnType>,
    },
}

enum ParamType {
    String,
    I32,
    I64,
    Bool,
}
impl ParamType {
    fn transformer_to_variant(&self) -> &str {
        match self {
            Self::String => "from_str",
            Self::I32 => "from_i32",
            Self::I64 => "from_i64",
            Self::Bool => "from_bool",
        }
    }
}
impl fmt::Display for ParamType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{}",
            match self {
                Self::String => "String",
                Self::I32 => "i32",
                Self::I64 => "i64",
                Self::Bool => "bool",
            }
        )
    }
}
impl From<&str> for ParamType {
    fn from(value: &str) -> Self {
        match value {
            "String" => Self::String,
            "i32" => Self::I32,
            "i64" => Self::I64,
            "bool" => Self::Bool,
            _ => panic!("Parameter type error: one of the function parameters cannot be transformed by this library.")
        }
    }
}

#[derive(Debug, Clone)]
enum ReturnType {
    None,
    String,
    I32,
    I64,
    Bool,
}
impl ReturnType {
    fn transformer_to_variant(&self) -> &str {
        match self {
            Self::None => panic!("none cannot be made into a variant"),
            Self::String => "from_str",
            Self::I32 => "from_i32",
            Self::I64 => "from_i64",
            Self::Bool => "from_bool",
        }
    }
    fn transformer_from_variant(&self) -> &str {
        match self {
            Self::None => panic!("no transformer for none"),
            Self::String => "to_string",
            Self::I32 => "to_i32",
            Self::I64 => "to_i64",
            Self::Bool => "to_bool",
        }
    }
}
impl fmt::Display for ReturnType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{}",
            match self {
                Self::None => "()",
                Self::String => "String",
                Self::I32 => "i32",
                Self::I64 => "i64",
                Self::Bool => "bool",
            }
        )
    }
}
impl From<&str> for ReturnType {
    fn from(value: &str) -> Self {
        match value {
            "()" => Self::None,
            "String" => Self::String,
            "i32" => Self::I32,
            "i64" => Self::I64,
            "bool" => Self::Bool,
            _ => panic!("Parameter type error: one of the function parameters cannot be transformed by this library.")
        }
    }
}

fn divide_items(stream: TokenStream) -> Vec<ChildItem> {
    let mut output = Vec::new();
    let mut buffer = ChildItem::Invalid;
    let mut state = DivideItemState::ExpectFnVarNameOrMut;

    for item in stream {
        match state {
            DivideItemState::ExpectFnVarNameOrMut => match item {
                TokenTree::Ident(id) => {
                    if id.to_string() == "fn" {
                        buffer = ChildItem::Function {
                            name: None,
                            params: None,
                            return_typ: None,
                        };
                        state = DivideItemState::ExpectFnName;
                    } else if id.to_string() == "mut" {
                        buffer = ChildItem::Variable {
                            mutable: true,
                            name: None,
                            typ: None,
                        };
                        state = DivideItemState::ExpectVarName;
                    } else {
                        buffer = ChildItem::Variable {
                            mutable: false,
                            name: Some(id.to_string()),
                            typ: None,
                        };
                        state = DivideItemState::ExpectVarSep;
                    }
                }
                _ => panic!("Syntax error: expected variable name or `fn`"),
            },
            DivideItemState::ExpectVarName => match item {
                TokenTree::Ident(id) => {
                    match &mut buffer {
                        ChildItem::Variable {
                            mutable: _,
                            name,
                            typ: _,
                        } => *name = Some(id.to_string()),
                        _ => unreachable!(),
                    };
                    state = DivideItemState::ExpectVarSep;
                }
                _ => panic!("Syntax error: expected variable name"),
            },
            DivideItemState::ExpectVarSep => match item {
                TokenTree::Punct(p) => {
                    if p.as_char() != ':' {
                        panic!("Syntax error: expected `:`");
                    }
                    state = DivideItemState::ExpectVarType;
                }
                _ => panic!("Syntax error: expected `:`"),
            },
            DivideItemState::ExpectVarType => match item {
                TokenTree::Ident(id) => {
                    match &mut buffer {
                        ChildItem::Variable {
                            mutable: _,
                            name: _,
                            typ,
                        } => *typ = Some(id.to_string().as_str().into()),
                        _ => unreachable!(),
                    };
                    state = DivideItemState::ExpectSep;
                }
                _ => panic!("Syntax error: expected variable name"),
            },
            DivideItemState::ExpectSep => match item {
                TokenTree::Punct(p) => {
                    if p.as_char() != ',' {
                        panic!("Syntax error: expected `,`");
                    }
                    output.push(buffer);
                    buffer = ChildItem::Invalid;
                    state = DivideItemState::ExpectFnVarNameOrMut;
                }
                _ => panic!("Syntax error: expected `,`"),
            },
            DivideItemState::ExpectFnName => match item {
                TokenTree::Ident(id) => {
                    match &mut buffer {
                        ChildItem::Function {
                            name,
                            params: _,
                            return_typ: _,
                        } => *name = Some(id.to_string()),
                        _ => unreachable!(),
                    };
                    state = DivideItemState::ExpectFnParams;
                }
                _ => panic!("Syntax error: expected function name"),
            },
            DivideItemState::ExpectFnParams => match item {
                TokenTree::Group(group) => {
                    match &mut buffer {
                        ChildItem::Function {
                            name: _,
                            params,
                            return_typ: _,
                        } => *params = Some(group),
                        _ => unreachable!(),
                    };
                    state = DivideItemState::ExpectFnArrow1OrSep;
                }
                _ => panic!("Syntax error: expected function name"),
            },
            DivideItemState::ExpectFnArrow1OrSep => match item {
                TokenTree::Punct(p) => {
                    if p.as_char() == ',' {
                        state = DivideItemState::ExpectFnVarNameOrMut;
                    } else if p.as_char() == '-' {
                        state = DivideItemState::ExpectFnArrow2;
                    } else {
                        panic!("Syntax error: expected `,` or `->`");
                    }
                }
                _ => panic!("Syntax error: expected `,` or `->`"),
            },
            DivideItemState::ExpectFnArrow2 => match item {
                TokenTree::Punct(p) => {
                    if p.as_char() != '>' {
                        panic!("Syntax error: expected `->`");
                    }
                    state = DivideItemState::ExpectFnReturnType;
                }
                _ => panic!("Syntax error: expected `->`"),
            },
            DivideItemState::ExpectFnReturnType => match item {
                TokenTree::Ident(id) => {
                    match &mut buffer {
                        ChildItem::Function {
                            name: _,
                            params: _,
                            return_typ,
                        } => *return_typ = Some(ReturnType::from(id.to_string().as_str())),
                        _ => unreachable!(),
                    };
                    state = DivideItemState::ExpectSep;
                }
                _ => panic!("Syntax error: expected function return type"),
            },
        }
    }

    match state {
        DivideItemState::ExpectFnArrow2 => panic!("Syntax error: expected `->`"),
        DivideItemState::ExpectFnName => panic!("Syntax error: expected function name"),
        DivideItemState::ExpectFnParams => panic!("Expected function parameters"),
        DivideItemState::ExpectFnReturnType => panic!("Expected function return type"),
        DivideItemState::ExpectVarName => panic!("Syntax error: expected variable name"),
        DivideItemState::ExpectVarSep => panic!("Syntax error: expected `:`"),
        DivideItemState::ExpectVarType => panic!("Syntax error: expected variable type"),
        DivideItemState::ExpectFnArrow1OrSep => (),
        DivideItemState::ExpectFnVarNameOrMut => (),
        DivideItemState::ExpectSep => (),
    }

    match buffer {
        ChildItem::Invalid => (),
        ChildItem::Function {
            name: _,
            params: _,
            return_typ: _,
        } => output.push(buffer),
        ChildItem::Variable {
            mutable: _,
            name: _,
            typ: _,
        } => output.push(buffer),
    }

    output
}

fn parse_param_group(params: Group) -> Vec<ParamType> {
    let mut output = Vec::new();
    if params.delimiter() != Delimiter::Parenthesis {
        panic!("Syntax error: expected `(... params ...)`");
    }
    let mut is_separator = false;
    for item in params.stream() {
        match item {
            TokenTree::Ident(id) => {
                if is_separator {
                    panic!("Syntax error: expected `,`");
                } else {
                    output.push(id.to_string().as_str().into());
                    is_separator = true;
                }
            }
            TokenTree::Punct(p) => {
                if is_separator {
                    if p.as_char() != ',' {
                        panic!("Syntax error: expected `,`");
                    }
                    is_separator = false;
                } else {
                    panic!("Syntax error: expected parameter type");
                }
            }
            _ => {
                if is_separator {
                    panic!("Syntax error: expected `,`")
                } else {
                    panic!("Syntax error: expected parameter type")
                }
            }
        }
    }
    output
}

fn build_item_token_strem(item: ChildItem, trait_body_stream: &mut Vec<TokenTree>) {
    match item {
        ChildItem::Invalid => unreachable!(),
        ChildItem::Function {
            name,
            params,
            return_typ,
        } => {
            use heck::*;
            // fn $snake_name(&self) -> crate::Result<()>
            let mut fn_definition = String::new();
            let mut variant_args = String::new();
            let mut index = 0;
            for p in parse_param_group(params.unwrap()) {
                fn_definition.push_str(&format!("p{index}: {p},"));
                variant_args.push_str(&format!(
                    "::com_shim::VARIANT::{}(p{index}),",
                    p.transformer_to_variant()
                ));
                index += 1;
            }

            let return_typ = return_typ.unwrap_or(ReturnType::None);
            trait_body_stream.extend(
                format!(
                    "fn {}(&self,{fn_definition}) -> ::com_shim::Result<{return_typ}>",
                    name.as_ref().unwrap().to_snake_case()
                )
                .parse::<TokenStream>()
                .unwrap(),
            );

            let result_transformer = match return_typ {
                ReturnType::None => "()".to_owned(),
                a => format!("r.{}()?", a.transformer_from_variant()),
            };
            trait_body_stream.extend(
                format!(
                    r#"{{
                use ::com_shim::{{IDispatchExt, VariantExt}};
                let r = self.get_idispatch().call("{}", vec![{}])?;
                ::std::result::Result::Ok({result_transformer})
            }}"#,
                    name.unwrap(),
                    variant_args
                )
                .parse::<TokenStream>()
                .unwrap(),
            );
        }
        ChildItem::Variable { mutable, name, typ } => {
            // get function
            use heck::*;
            trait_body_stream.extend(
                format!(
                    r#"fn {}(&self) -> ::com_shim::Result<{}> {{
                    use ::com_shim::{{IDispatchExt, VariantExt}};
                    Ok(self.get_idispatch().get("{}")?.{}()?)
                }}"#,
                    name.clone().unwrap().to_snake_case(),
                    typ.clone().unwrap(),
                    name.clone().unwrap(),
                    typ.clone().unwrap().transformer_from_variant()
                )
                .parse::<TokenStream>()
                .unwrap(),
            );

            if mutable {
                // set function
                // fn $snake_name(&self) -> ::com_shim::Result<$kind> {
                //     use ::com_shim::IDispatchExt;
                //     Ok(self.get_idispatch().set($name)?.$transformer()?)
                // }
                trait_body_stream.extend(
                    format!(
                        r#"fn set_{}(&self, value: {}) -> ::com_shim::Result<()> {{
                        use ::com_shim::{{IDispatchExt, VariantExt}};
                        let _ = self.get_idispatch().set("{}", ::com_shim::VARIANT::{}(value))?;
                        Ok(())
                    }}"#,
                        name.as_ref().unwrap().to_snake_case(),
                        typ.clone().unwrap(),
                        name.unwrap(),
                        typ.unwrap().transformer_to_variant()
                    )
                    .parse::<TokenStream>()
                    .unwrap(),
                );
            }
        }
    }
}
