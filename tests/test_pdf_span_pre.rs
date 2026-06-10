use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_span_pre() {
    let html = r#"<!DOCTYPE html><html><head><style>
p { margin: 0; white-space: pre-wrap; }
.s { white-space: pre-wrap; font-size: 16px; }
</style></head><body>
<p><span class="s">    FourLeadingSpaces</span></p>
</body></html>"#;
    
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG SPAN PRE LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
