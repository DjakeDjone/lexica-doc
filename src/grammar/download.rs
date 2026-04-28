use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, bail, Context, Result};
use zip::ZipArchive;

pub const LT_STABLE_ZIP_URL: &str = "https://languagetool.org/download/LanguageTool-stable.zip";

pub async fn download_languagetool_server_jar(target_jar_path: PathBuf) -> Result<()> {
    let client = reqwest::Client::new();
    let response = client
        .get(LT_STABLE_ZIP_URL)
        .send()
        .await
        .context("Failed to start LanguageTool download")?
        .error_for_status()
        .context("LanguageTool download returned error status")?;
    let zip_bytes = response
        .bytes()
        .await
        .context("Failed to read LanguageTool archive bytes")?;
    let archive = zip_bytes.to_vec();

    tokio::task::spawn_blocking(move || {
        let target_dir = target_jar_path
            .parent()
            .ok_or_else(|| anyhow!("Invalid target JAR path: {}", target_jar_path.display()))?;
        fs::create_dir_all(target_dir)
            .with_context(|| format!("Failed to create {}", target_dir.display()))?;
        extract_distribution_into_dir(&archive, target_dir)?;

        if !target_jar_path.exists() {
            bail!(
                "Downloaded LanguageTool archive did not produce {}",
                target_jar_path.display()
            );
        }

        Ok::<(), anyhow::Error>(())
    })
    .await
    .map_err(|join_error| anyhow!("LanguageTool download task join error: {join_error}"))??;

    Ok(())
}

fn extract_distribution_into_dir(zip_bytes: &[u8], target_dir: &Path) -> Result<()> {
    let reader = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(reader).context("Failed to open LanguageTool ZIP archive")?;
    let mut extracted_server_jar = false;

    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .with_context(|| format!("Failed to read ZIP entry at index {index}"))?;
        let Some(relative_path) = strip_archive_root(entry.name()) else {
            continue;
        };
        let output_path = target_dir.join(relative_path);

        if entry.is_dir() {
            fs::create_dir_all(&output_path)
                .with_context(|| format!("Failed to create {}", output_path.display()))?;
            continue;
        }

        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("Failed to create {}", parent.display()))?;
        }

        let tmp_path = output_path.with_extension("download");
        let mut output_file = fs::File::create(&tmp_path)
            .with_context(|| format!("Failed to create {}", tmp_path.display()))?;
        std::io::copy(&mut entry, &mut output_file)
            .with_context(|| format!("Failed to extract {}", output_path.display()))?;
        fs::rename(&tmp_path, &output_path)
            .with_context(|| format!("Failed to finalize {}", output_path.display()))?;

        if output_path
            .file_name()
            .and_then(|name| name.to_str())
            .is_some_and(|name| name.eq_ignore_ascii_case("languagetool-server.jar"))
        {
            extracted_server_jar = true;
        }
    }

    if !extracted_server_jar {
        bail!("Could not find languagetool-server.jar in downloaded archive");
    }

    Ok(())
}

fn strip_archive_root(path: &str) -> Option<&str> {
    let (first, rest) = path.split_once('/')?;
    if first.is_empty() || rest.is_empty() {
        None
    } else {
        Some(rest)
    }
}
