use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_missing_font() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; margin: 0; }
            .missing { font-family: 'MissingFont'; font-size: 24pt; color: #ff0000; font-weight: bold; }
        </style>
        </head>
        <body>
            <p><span class="missing">StyledText</span></p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG MISSING FONT LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
