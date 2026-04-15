use std::fs;

use insta::assert_yaml_snapshot;
use walkdir::WalkDir;
use wors::document::docx::docx_to_document;

#[test]
fn snapshot_all_docx_fixtures() {
    let fixtures_dir = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/docx");
    let mut found = false;

    for entry in WalkDir::new(fixtures_dir)
        .into_iter()
        .filter_map(|e| e.ok())
        .filter(|e| {
            e.path()
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("docx"))
        })
    {
        found = true;
        let path = entry.path();
        let stem = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("unknown");

        let bytes =
            fs::read(path).unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
        let imported = docx_to_document(&bytes)
            .unwrap_or_else(|e| panic!("failed to parse {}: {e}", path.display()));

        // Snapshot name is derived from the .docx filename.
        let snapshot_name = stem.replace(|c: char| !c.is_alphanumeric() && c != '_', "_");
        assert_yaml_snapshot!(snapshot_name, imported);
    }

    if !found {
        eprintln!(
            "warning: no .docx fixtures found in {fixtures_dir}; \
             add files to tests/fixtures/docx/ and re-run"
        );
    }
}
