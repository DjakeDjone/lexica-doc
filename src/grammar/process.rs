use std::{
    path::PathBuf,
    process::{Child, Command, Stdio},
    time::Duration,
};

use anyhow::{bail, Context, Result};
use tokio::time::sleep;

use super::GrammarConfig;

const HEALTH_ATTEMPTS: usize = 10;
const HEALTH_WAIT_MS: u64 = 300;

#[cfg(target_os = "windows")]
const JAVA_BIN: &str = "javaw";

#[cfg(not(target_os = "windows"))]
const JAVA_BIN: &str = "java";

pub fn default_lt_jar_path() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|exe| {
            exe.parent()
                .map(|parent| parent.join("languagetool-server.jar"))
        })
        .unwrap_or_else(|| PathBuf::from("languagetool-server.jar"))
}

pub fn spawn_languagetool(config: &GrammarConfig) -> Result<Child> {
    if !config.lt_jar_path.exists() {
        bail!(
            "LanguageTool server JAR not found: {}",
            config.lt_jar_path.display()
        );
    }

    let Some(parent_dir) = config.lt_jar_path.parent() else {
        bail!(
            "LanguageTool server JAR has no parent directory: {}",
            config.lt_jar_path.display()
        );
    };
    let libs_dir = parent_dir.join("libs");
    if !libs_dir.is_dir() {
        bail!(
            "LanguageTool distribution is incomplete: missing {}. Re-download LanguageTool from the Grammer tab.",
            libs_dir.display()
        );
    }

    Command::new(JAVA_BIN)
        .current_dir(parent_dir)
        .arg("-jar")
        .arg(&config.lt_jar_path)
        .arg("--allow-origin")
        .arg("*")
        .arg("--port")
        .arg(config.port.to_string())
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .with_context(|| {
            format!(
                "Failed to spawn LanguageTool server from {}",
                config.lt_jar_path.display()
            )
        })
}

pub fn kill_languagetool(child: &mut Child) {
    let _ = child.kill();
    let _ = child.wait();
}

pub async fn wait_until_ready(client: &reqwest::Client, port: u16) -> bool {
    let url = format!("http://localhost:{port}/v2/languages");
    for _ in 0..HEALTH_ATTEMPTS {
        if let Ok(response) = client.get(&url).send().await {
            if response.status().is_success() {
                return true;
            }
        }
        sleep(Duration::from_millis(HEALTH_WAIT_MS)).await;
    }
    false
}
