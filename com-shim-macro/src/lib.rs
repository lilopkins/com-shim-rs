#![deny(unsafe_code)]
#![warn(clippy::pedantic)]

use std::fmt;

use proc_macro::{self, Delimiter, Group, TokenStream, TokenTree};

macro_rules! debug_log {
    ($($arg:expr),*) => {
        #[cfg(feature = "debug")] eprintln!($($arg),*);
    };
}

#[proc_macro]
pub fn com_shim(stream: TokenStream) -> TokenStream {
    let mut result_stream = TokenStream::new();
    let mut stream_iter = stream.into_iter().peekable();

    // Class token
    let class_token = stream_iter.next().expect("Syntax error: expected `class`");
    match class_token {
        TokenTree::Ident(id) => {
            assert!(id.to_string() == *"class", "Syntax error: expected `class`");
            result_stream.extend("pub struct".parse::<TokenStream>().unwrap());
        }
        _ => panic!("Syntax error: expect `class` ident"),
    }

    // Name
    let name_token = stream_iter.next().expect("Syntax error: expected name");
    match &name_token {
        proc_macro::TokenTree::Ident(id) => {
            result_stream.extend(vec![TokenTree::Ident(id.clone())]);
        }
        _ => panic!("Syntax error: expect name identifier"),
    }
    let name = name_token.to_string();
    debug_log!("Generating struct for {}", name);

    // Push struct group
    result_stream.extend(
        "{ inner: ::com_shim::IDispatch }"
            .parse::<TokenStream>()
            .unwrap(),
    );

    // Inherit HasIDispatch trait
    debug_log!("Writing HasIDispatch entry for {}", name);
    result_stream.extend(format!("impl ::com_shim::HasIDispatch for {name} {{ fn get_idispatch(&self) -> &::com_shim::IDispatch {{ &self.inner }} }}").parse::<TokenStream>().unwrap());

    if stream_iter
        .peek()
        .expect("Syntax error: expected `:` or start of class")
        .to_string()
        == *":"
    {
        loop {
            let _separator_token = stream_iter.next().unwrap();
            let parent_token = stream_iter
                .next()
                .expect("Syntax Error: expected identifier of parent class after `:`");
            match parent_token {
                TokenTree::Ident(id) => {
                    debug_log!("Class {} has parent {}", name, id);
                    result_stream.extend(
                        format!("impl {id}_Impl for {name} {{}}")
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
                    assert!(
                        p.to_string() == *"+",
                        "Syntax Error: expected identifier of parent class after `:`"
                    );
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
    result_stream.extend(
        format!(
            "pub trait {name}_Impl<T: ::com_shim::HasIDispatch = Self>: ::com_shim::HasIDispatch<T>"
        )
        .parse::<TokenStream>()
        .unwrap(),
    );

    let mut trait_body_stream: Vec<TokenTree> = Vec::new();
    match stream_iter
        .next()
        .expect("Syntax error: expected `{ ... }`")
    {
        TokenTree::Group(group) => {
            assert!(
                !(group.delimiter() != Delimiter::Brace),
                "Syntax error: expected `{{ ... }}`"
            );
            // parse group members
            let items = divide_items(group.stream());
            debug_log!("Trait for {} has items:", name);
            for item in items {
                debug_log!("  {item}");
                build_item_token_strem(item, &mut trait_body_stream);
            }
        }
        _ => panic!("Syntax error: expected `{{ ... }}`"),
    }
    result_stream.extend(vec![TokenTree::Group(Group::new(
        Delimiter::Brace,
        TokenStream::from_iter(trait_body_stream),
    ))]);

    // create From<IDispatch> trait
    result_stream.extend(format!("impl ::std::convert::From<::com_shim::IDispatch> for {name} {{ fn from(value: ::com_shim::IDispatch) -> Self {{ Self {{ inner: value }} }} }}").parse::<TokenStream>().unwrap());

    assert!(
        stream_iter.next().is_none(),
        "Syntax error: expected end of shim definition."
    );

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

impl fmt::Display for ChildItem {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ChildItem::Invalid => panic!("cannot render invalid option!"),
            ChildItem::Variable { mutable, name, typ } => write!(
                f,
                "{}: {} (mutable: {mutable})",
                name.as_ref().unwrap(),
                typ.as_ref().unwrap()
            ),
            ChildItem::Function {
                name,
                params: _,
                return_typ,
            } => write!(f, "fn {} -> {return_typ:?}", name.as_ref().unwrap()),
        }
    }
}

enum ParamType {
    String,
    I16,
    I32,
    I64,
    U8,
    U16,
    U32,
    U64,
    Bool,
}
impl ParamType {
    fn transformer_to_variant(&self) -> &str {
        match self {
            Self::String => "from_str",
            Self::I16 => "from_i16",
            Self::I32 => "from_i32",
            Self::I64 => "from_i64",
            Self::Bool => "from_bool",
            Self::U8 => "from_u8",
            Self::U16 => "from_u16",
            Self::U32 => "from_u32",
            Self::U64 => "from_u64",
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
                Self::I16 => "i16",
                Self::I32 => "i32",
                Self::I64 => "i64",
                Self::U8 => "u8",
                Self::U16 => "u16",
                Self::U32 => "u32",
                Self::U64 => "u64",
                Self::Bool => "bool",
            }
        )
    }
}
impl From<&str> for ParamType {
    fn from(value: &str) -> Self {
        match value {
            "String" => Self::String,
            "i16" => Self::I16,
            "i32" => Self::I32,
            "i64" => Self::I64,
            "u8" => Self::U8,
            "u16" => Self::U16,
            "u32" => Self::U32,
            "u64" => Self::U64,
            "bool" => Self::Bool,
            _ => panic!(
                "Parameter type error: one of the function parameters cannot be transformed by this library."
            ),
        }
    }
}

#[derive(Debug, Clone)]
enum ReturnType {
    None,
    String,
    I16,
    I32,
    I64,
    U8,
    U16,
    U32,
    U64,
    Bool,
    VariantInto(String),
}
impl ReturnType {
    fn transformer_to_variant(&self) -> &str {
        match self {
            Self::None => panic!("none cannot be made into a variant"),
            Self::VariantInto(kind) => panic!("object {kind} cannot be made into a variant"),
            Self::String => "from_str",
            Self::I16 => "from_i16",
            Self::I32 => "from_i32",
            Self::I64 => "from_i64",
            Self::U8 => "from_u8",
            Self::U16 => "from_u16",
            Self::U32 => "from_u32",
            Self::U64 => "from_u64",
            Self::Bool => "from_bool",
        }
    }
    fn transformer_from_variant(&self) -> &str {
        match self {
            Self::None => panic!("no transformer for none"),
            Self::VariantInto(_) => panic!("no transformer for any object"),
            Self::String => "to_string",
            Self::I16 => "to_i16",
            Self::I32 => "to_i32",
            Self::I64 => "to_i64",
            Self::U8 => "to_u8",
            Self::U16 => "to_u16",
            Self::U32 => "to_u32",
            Self::U64 => "to_u64",
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
                Self::None => "()".to_owned(),
                Self::VariantInto(a) => a.to_string(),
                Self::String => "String".to_owned(),
                Self::I16 => "i16".to_owned(),
                Self::I32 => "i32".to_owned(),
                Self::I64 => "i64".to_owned(),
                Self::U8 => "u8".to_owned(),
                Self::U16 => "u16".to_owned(),
                Self::U32 => "u32".to_owned(),
                Self::U64 => "u64".to_owned(),
                Self::Bool => "bool".to_owned(),
            }
        )
    }
}
impl From<&str> for ReturnType {
    fn from(value: &str) -> Self {
        match value {
            "()" => Self::None,
            "String" => Self::String,
            "i16" => Self::I16,
            "i32" => Self::I32,
            "i64" => Self::I64,
            "u8" => Self::U8,
            "u16" => Self::U16,
            "u32" => Self::U32,
            "u64" => Self::U64,
            "bool" => Self::Bool,
            s => Self::VariantInto(s.to_owned()),
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
                    assert!(p.as_char() == ':', "Syntax error: expected `:`");
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
                    assert!(p.as_char() == ',', "Syntax error: expected `,`");
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
                        output.push(buffer);
                        buffer = ChildItem::Invalid;
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
                    assert!(p.as_char() == '>', "Syntax error: expected `->`");
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
        DivideItemState::ExpectFnArrow1OrSep
        | DivideItemState::ExpectFnVarNameOrMut
        | DivideItemState::ExpectSep => (),
    }

    match buffer {
        ChildItem::Invalid => (),
        ChildItem::Function {
            name: _,
            params: _,
            return_typ: _,
        }
        | ChildItem::Variable {
            mutable: _,
            name: _,
            typ: _,
        } => output.push(buffer),
    }

    output
}

fn parse_param_group(params: &Group) -> Vec<ParamType> {
    let mut output = Vec::new();
    assert!(
        !(params.delimiter() != Delimiter::Parenthesis),
        "Syntax error: expected `(... params ...)`"
    );
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
                    assert!(p.as_char() == ',', "Syntax error: expected `,`");
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
            use heck::ToSnakeCase;
            // fn $snake_name(&self) -> crate::Result<()>
            let mut fn_definition = String::new();
            let mut variant_args = String::new();
            for (index, p) in parse_param_group(params.as_ref().unwrap())
                .iter()
                .enumerate()
            {
                fn_definition.push_str(&format!("p{index}: {p},"));
                variant_args.push_str(&format!(
                    "::com_shim::VARIANT::{}(p{index}),",
                    p.transformer_to_variant()
                ));
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
                ReturnType::VariantInto(target) => {
                    format!("{target}::from(r.to_idispatch()?.clone())")
                }
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
            use heck::ToSnakeCase;
            let get_result = format!(r#"self.get_idispatch().get("{}")?"#, name.clone().unwrap());
            let last_line = match typ.clone().unwrap() {
                ReturnType::VariantInto(to) => {
                    format!("Ok({to}::from({get_result}.to_idispatch()?.clone()))")
                }
                typ => format!("Ok({get_result}.{}()?)", typ.transformer_from_variant()),
            };
            trait_body_stream.extend(
                format!(
                    r"fn {}(&self) -> ::com_shim::Result<{}> {{
                    use ::com_shim::{{IDispatchExt, VariantExt}};
                    {last_line}
                }}",
                    safe_name(name.clone().unwrap().to_snake_case()),
                    typ.clone().unwrap(),
                )
                .parse::<TokenStream>()
                .unwrap(),
            );

            if mutable {
                // set function
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

fn safe_name<S: AsRef<str>>(name: S) -> String {
    let name = name.as_ref();
    let keywords = vec![
        "as", "break", "const", "continue", "crate", "else", "enum", "extern", "false", "fn",
        "for", "if", "impl", "in", "let", "loop", "match", "mod", "move", "mut", "pub", "ref",
        "return", "self", "Self", "static", "struct", "super", "trait", "true", "type", "unsafe",
        "use", "where", "while", "async", "await", "dyn", "abstract", "become", "box", "do",
        "final", "macro", "override", "priv", "typeof", "unsized", "virtual", "yield", "try",
    ];
    if keywords.contains(&name) {
        format!("_{name}")
    } else {
        name.to_owned()
    }
}
