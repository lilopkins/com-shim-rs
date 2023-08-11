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

        fn GetListProperty(String) -> String,
    }
}

fn main() {
    // The following call now would trigger a COM call:
    // let a: GuiTextField;
    // a.get_list_property("property");
}
