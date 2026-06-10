use wors::document::DocumentState;

fn main() {
    let mut doc = DocumentState::bootstrap();
    doc.delete_range(0..doc.plain_text().chars().count());
    
    let lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
    
    doc.insert_text(0, lorem, Default::default());
    
    let bytes = doc.to_pdf_bytes().expect("Failed to generate PDF");
    std::fs::write("lorem_ipsum.pdf", bytes).expect("Failed to write PDF");
    println!("Successfully created lorem_ipsum.pdf!");
}
