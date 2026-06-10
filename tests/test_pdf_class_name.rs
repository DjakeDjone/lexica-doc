use printpdf::*;
use std::collections::BTreeMap;

#[test]
fn test_pdf_class_name() {
    let html = r#"
        <html lang="en">
        <head>
        <style>
            body { font-size: 16px; margin: 0; }
            .s-0-0 { font-size: 24pt; color: #ff0000; font-weight: bold; }
        </style>
        </head>
        <body>
            <p><span class="s-0-0">StyledText</span></p>
            <p>NormalText</p>
        </body>
        </html>
    "#;
    let mut warnings = Vec::new();
    let options = GeneratePdfOptions::default();
    let (_, debug_info) = PdfDocument::from_html_debug(html, &BTreeMap::new(), &BTreeMap::new(), &options, &mut warnings).unwrap();
    println!("DEBUG CLASS NAME LISTS:");
    for d in debug_info.display_list_debug {
        println!("{}", d);
    }
}
