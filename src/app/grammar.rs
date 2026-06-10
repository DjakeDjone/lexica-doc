use std::path::PathBuf;
#[cfg(not(target_arch = "wasm32"))]
use tokio::sync::mpsc;

use super::WorsApp;
use crate::grammar::GrammarStatus;
#[cfg(not(target_arch = "wasm32"))]
use crate::grammar::{
    download::{download_languagetool_server_jar, LT_STABLE_ZIP_URL},
    process::{kill_languagetool, spawn_languagetool},
    task::{run_grammar_task, GrammarRequest, GrammarTaskResult},
    GrammarChecker,
};

const GRAMMAR_QUEUE_CAPACITY: usize = 8;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum GrammarDownloadStatus {
    Idle,
    Downloading,
}

#[derive(Debug)]
#[cfg(not(target_arch = "wasm32"))]
pub(crate) enum GrammarDownloadResult {
    Ready(PathBuf),
    Failed(String),
}

impl WorsApp {
    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn start_grammar_service(&mut self) -> Result<(), String> {
        self.stop_grammar_service();

        if !self.grammar_config.lt_jar_path.exists() {
            return Err(format!(
                "LanguageTool JAR not found at {}",
                self.grammar_config.lt_jar_path.display()
            ));
        }

        let Some(runtime) = self._grammar_runtime.as_ref() else {
            return Err("Grammar runtime is unavailable".to_owned());
        };

        let child = spawn_languagetool(&self.grammar_config)
            .map_err(|error| format!("Grammar unavailable: {error}"))?;
        let (tx, rx) = mpsc::channel(GRAMMAR_QUEUE_CAPACITY);
        let (results_tx, results_rx) = mpsc::channel(GRAMMAR_QUEUE_CAPACITY);

        runtime.spawn(run_grammar_task(
            rx,
            results_tx,
            GrammarChecker::new(self.grammar_config.port),
            self.grammar_config.port,
        ));

        self.grammar_process = Some(child);
        self.grammar_tx = Some(tx);
        self.grammar_results_rx = Some(results_rx);
        self.grammar_status = GrammarStatus::Idle;
        Ok(())
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn start_grammar_service(&mut self) -> Result<(), String> {
        Err("Grammar checking is not available in the web build".to_owned())
    }

    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn stop_grammar_service(&mut self) {
        self.grammar_tx = None;
        self.grammar_results_rx = None;
        if let Some(child) = self.grammar_process.as_mut() {
            kill_languagetool(child);
        }
        self.grammar_process = None;
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn stop_grammar_service(&mut self) {}

    pub(crate) fn restart_grammar_service(&mut self) {
        match self.start_grammar_service() {
            Ok(()) => {
                self.grammar_warning_message = None;
                self.show_grammar_warning = false;
                self.status_message = "Grammar server restarted".to_owned();
            }
            Err(message) => {
                self.grammar_status = GrammarStatus::Unavailable(message.clone());
                self.grammar_warning_message = Some(message);
                self.show_grammar_warning = true;
            }
        }
    }

    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn poll_grammar_results(&mut self) {
        let Some(results_rx) = self.grammar_results_rx.as_mut() else {
            return;
        };

        while let Ok(message) = results_rx.try_recv() {
            match message {
                GrammarTaskResult::Completed(errors) => {
                    self.grammar_errors = errors;
                    self.grammar_status = GrammarStatus::Done;
                }
                GrammarTaskResult::Unavailable(message) => {
                    self.grammar_status = GrammarStatus::Unavailable(message);
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn poll_grammar_results(&mut self) {}

    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn start_grammar_download(&mut self) {
        if self.grammar_download_status == GrammarDownloadStatus::Downloading {
            return;
        }

        let Some(runtime) = self._grammar_runtime.as_ref() else {
            self.grammar_status = GrammarStatus::Unavailable(
                "Cannot download LanguageTool because runtime is unavailable".to_owned(),
            );
            return;
        };

        let target_path = self.grammar_config.lt_jar_path.clone();
        let (tx, rx) = mpsc::unbounded_channel::<GrammarDownloadResult>();
        runtime.spawn(async move {
            let result = match download_languagetool_server_jar(target_path.clone()).await {
                Ok(()) => GrammarDownloadResult::Ready(target_path),
                Err(error) => GrammarDownloadResult::Failed(error.to_string()),
            };
            let _ = tx.send(result);
        });

        self.grammar_download_rx = Some(rx);
        self.grammar_download_status = GrammarDownloadStatus::Downloading;
        self.show_grammar_warning = true;
        self.grammar_warning_message = Some(format!(
            "Downloading LanguageTool from {LT_STABLE_ZIP_URL}. This can take a while."
        ));
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn start_grammar_download(&mut self) {
        self.grammar_status = GrammarStatus::Unavailable(
            "Grammar downloads are not available in the web build".to_owned(),
        );
        self.status_message = "Grammar download unavailable on web".to_owned();
    }

    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn poll_grammar_download(&mut self) {
        let mut drained = Vec::new();
        if let Some(rx) = self.grammar_download_rx.as_mut() {
            while let Ok(message) = rx.try_recv() {
                drained.push(message);
            }
        }
        if drained.is_empty() {
            return;
        }

        for message in drained {
            match message {
                GrammarDownloadResult::Ready(path) => {
                    self.grammar_download_status = GrammarDownloadStatus::Idle;
                    self.grammar_download_rx = None;
                    self.grammar_warning_message =
                        Some(format!("LanguageTool downloaded to {}", path.display()));
                    self.show_grammar_warning = false;
                    self.status_message = "LanguageTool downloaded".to_owned();
                    if let Err(error_message) = self.start_grammar_service() {
                        self.grammar_status = GrammarStatus::Unavailable(error_message);
                        self.show_grammar_warning = true;
                    } else {
                        self.grammar_status = GrammarStatus::Idle;
                        self.request_grammar_check(true);
                    }
                }
                GrammarDownloadResult::Failed(error_message) => {
                    self.grammar_download_status = GrammarDownloadStatus::Idle;
                    self.grammar_download_rx = None;
                    self.grammar_status = GrammarStatus::Unavailable(format!(
                        "LanguageTool download failed: {error_message}"
                    ));
                    self.show_grammar_warning = true;
                    self.grammar_warning_message =
                        Some(format!("LanguageTool download failed: {error_message}"));
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn poll_grammar_download(&mut self) {}

    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn request_grammar_check(&mut self, force: bool) {
        if !force && !self.grammar_auto_check {
            return;
        }
        if self.grammar_download_status == GrammarDownloadStatus::Downloading {
            self.status_message = "Grammar download in progress".to_owned();
            return;
        }

        if self.grammar_tx.is_none() {
            if let Err(message) = self.start_grammar_service() {
                self.grammar_status = GrammarStatus::Unavailable(message.clone());
                self.grammar_warning_message = Some(message);
                self.show_grammar_warning = true;
                return;
            }
        }

        let text = self.document.plain_text();
        let language = self
            .grammar_config
            .language
            .to_languagetool_code(&text)
            .to_owned();
        let request = GrammarRequest { text, language };

        let Some(tx) = self.grammar_tx.clone() else {
            return;
        };

        match tx.try_send(request.clone()) {
            Ok(()) => {
                self.grammar_status = GrammarStatus::Checking;
            }
            Err(mpsc::error::TrySendError::Full(_)) => {
                self.grammar_status = GrammarStatus::Checking;
            }
            Err(mpsc::error::TrySendError::Closed(_)) => {
                if let Err(message) = self.start_grammar_service() {
                    self.grammar_status = GrammarStatus::Unavailable(message.clone());
                    self.grammar_warning_message = Some(message);
                    self.show_grammar_warning = true;
                    return;
                }

                if let Some(restarted_tx) = self.grammar_tx.clone() {
                    match restarted_tx.try_send(request) {
                        Ok(()) | Err(mpsc::error::TrySendError::Full(_)) => {
                            self.grammar_status = GrammarStatus::Checking;
                        }
                        Err(mpsc::error::TrySendError::Closed(_)) => {
                            self.grammar_status = GrammarStatus::Unavailable(
                                "Grammar worker channel closed unexpectedly".to_owned(),
                            );
                        }
                    }
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    pub(crate) fn request_grammar_check(&mut self, _force: bool) {}
}

impl Drop for WorsApp {
    fn drop(&mut self) {
        self.stop_grammar_service();
    }
}
