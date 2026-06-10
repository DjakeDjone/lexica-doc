use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_pt() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; margin: 0; }
            .pt { font-size: 24pt; }
            .px { font-size: 32px; }
        </style>
        </head>
        <body>
            <p>Line1</p>
            <p><span class="pt">Line2 PT</span></p>
            <p><span class="px">Line3 PX</span></p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG PT LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
