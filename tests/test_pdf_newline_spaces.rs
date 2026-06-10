use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_newline_spaces() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; white-space: pre-wrap; }
        </style>
        </head>
        <body>
            <p>Line1
    SpacesOnLine2</p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG NEWLINE SPACES LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
