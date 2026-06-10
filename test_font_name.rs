use std::fs;

fn main() {
    let bytes = fs::read("assets/fonts/LiberationSans-Regular.ttf").unwrap();
    let parsed = rust_fontconfig::FcParseFontBytes(&bytes, "LiberationSans-Regular");
    if let Some(fonts) = parsed {
        for f in fonts {
            println!("Parsed font: family='{}', name='{}', weight={}, style='{:?}'", 
                     f.font_family, f.font_name, f.font_weight, f.font_style);
        }
    }
}
