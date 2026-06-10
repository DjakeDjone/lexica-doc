use printpdf::*;
use std::collections::BTreeMap;
use std::fs;

#[test]
fn test_pdf_fonts() {
    let html = r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body { margin: 0; padding: 0; color: #24272e; font-family: 'LiberationSans-Regular'; }
p { margin: 0; white-space: pre-wrap; }.p-0 { text-align:left;margin-top:0.00px;margin-bottom:0.00px; }
.s-0-0 { font-weight: bold; color: #24272e; font-size: 11.5pt; font-family: 'Carlito-Bold'; }
</style></head><body><p class="p-0"><span class="s-0-0">Einzug: </span></p></body></html>"#;
    
    let mut fonts: BTreeMap<String, Base64OrRaw> = BTreeMap::new();
    let font_bytes = fs::read("assets/fonts/Carlito-Bold.ttf").unwrap();
    fonts.insert("Carlito-Bold".to_owned(), Base64OrRaw::Raw(font_bytes));

    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &fonts, &options, &mut warnings).unwrap();
    println!("DEBUG WORS FONTS LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
