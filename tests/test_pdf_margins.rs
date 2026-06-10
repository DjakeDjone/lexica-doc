use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_margins() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { margin: 0; padding: 0; }
            .my-class { font-size: 24.00px; }
        </style>
        </head>
        <body>
            <p><span class="my-class">Huge Red Bold Text</span></p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions {
        margin_top: Some(25.4),
        margin_left: Some(25.4),
        ..Default::default()
    };
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG LISTS WITH MARGINS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
