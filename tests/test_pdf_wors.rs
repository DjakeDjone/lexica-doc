use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_wors() {
    let html = r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body { margin: 0; padding: 0; color: #24272e; font-family: 'LiberationSans-Regular'; }
p { margin: 0; white-space: pre-wrap; }.p-0 { text-align:Left;margin-top:0.00px;margin-bottom:0.00px; }
.s-0-0 { font-weight: bold; color: #24272e; font-size: 11.5pt; font-family: 'Carlito-Bold'; }
</style></head><body><p class="p-0"><span class="s-0-0">Einzug: </span></p></body></html>"#;
    
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG WORS LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
