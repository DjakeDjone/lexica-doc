use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_center() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; width: 500px; }
            .centered { text-align: center; }
            .right { text-align: right; }
        </style>
        </head>
        <body>
            <p>Left Text</p>
            <p class="centered">Center Text</p>
            <p class="right">Right Text</p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG CENTER LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
