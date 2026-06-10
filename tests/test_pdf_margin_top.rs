use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_margin_top() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; margin: 0; }
            .mt { margin-top: 50px; }
            .mb { margin-bottom: 50px; }
        </style>
        </head>
        <body>
            <p>Line1</p>
            <p class="mt">Line2 with margin top</p>
            <p class="mb">Line3 with margin bottom</p>
            <p>Line4</p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG MARGIN TOP LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
