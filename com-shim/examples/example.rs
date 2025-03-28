use com_shim::com_shim;

com_shim! {
    /// A generic GUI component
    struct GuiComponent {
        /// The component's text
        Text: String,
    }
}

com_shim! {
    /// A component that has visible behaviours
    struct GuiVComponent {
        /// Set the focus of this component
        fn SetFocus(),
    }
}

com_shim! {
    /// A text field component
    struct GuiTextField: GuiVComponent + GuiComponent {
        /// The caret position
        CaretPosition: i64,
        /// The displayed text
        DisplayedText: String,
        /// Whether the field is highlighted
        mut Highlighted: bool,

        /// Get a property from the component
        fn GetListProperty(String) -> GuiComponent,
    }
}

fn main() {
    // The following call now would trigger a COM call:
    // let a: GuiTextField;
    // a.get_list_property("property");
}
