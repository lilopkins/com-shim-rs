# COM Shim

Easily write interfaces that can read from COM, without worrying about the underlying functionality (unless you want to!).

## Example

```rust
use com_shim::com_shim;

com_shim! {
    class GuiComponent {
        Text: String,
    }
}

com_shim! {
    class GuiVComponent {
        fn SetFocus(),
    }
}

com_shim! {
    class GuiTextField: GuiVComponent + GuiComponent {
        CaretPosition: i64,
        DisplayedText: String,
        mut Highlighted: bool,

        fn GetListProperty(String) -> GuiComponent,
    }
}

fn main() {
    // The following call now would trigger a COM call:
    // let a: GuiTextField;
    // a.get_list_property("property");
}
```

You can also see it implemented in the [`sap-scripting`](https://github.com/lilopkins/sap-scripting-rs.git) package.
