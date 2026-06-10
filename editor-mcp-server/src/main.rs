use anyhow::{Context, Result};
use reqwest::Client;
use serde_json::json;
use std::io::{self, BufRead, Write};

#[tokio::main]
async fn main() -> Result<()> {
    // 1. Discover the port
    let port = if let Some(arg) = std::env::args().nth(1) {
        arg
    } else {
        let mut path = dirs::home_dir().context("No home directory")?;
        path.push(".wors");
        path.push("mcp-port");
        std::fs::read_to_string(path)
            .context("Failed to read ~/.wors/mcp-port. Is the editor running?")?
            .trim()
            .to_string()
    };

    let base_url = format!("http://127.0.0.1:{}", port);
    let client = Client::new();

    // 2. Read stdin and process MCP JSON-RPC
    let stdin = io::stdin();
    let stdout = io::stdout();
    let mut handle = stdout.lock();

    for line in stdin.lock().lines() {
        let line = line?;
        if line.trim().is_empty() {
            continue;
        }

        let req: serde_json::Value = match serde_json::from_str(&line) {
            Ok(r) => r,
            Err(_) => continue,
        };

        if let Some(method) = req.get("method").and_then(|m| m.as_str()) {
            let id = req.get("id").cloned();
            let default_params = json!({});
            let params = req.get("params").unwrap_or(&default_params);

            let response = match method {
                "initialize" => {
                    Some(json!({
                        "jsonrpc": "2.0",
                        "id": id,
                        "result": {
                            "protocolVersion": "2024-11-05",
                            "capabilities": {},
                            "serverInfo": {
                                "name": "editor-mcp-server",
                                "version": "0.1.0"
                            }
                        }
                    }))
                }
                "tools/list" => {
                    Some(json!({
                        "jsonrpc": "2.0",
                        "id": id,
                        "result": {
                            "tools": [
                                {
                                    "name": "editor_state",
                                    "description": "Return the current editor state, including title, dirty flag, cursor position, selection, and undo/redo availability.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {}
                                    }
                                },
                                {
                                    "name": "document_get_text",
                                    "description": "Return the current document as plain text.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {}
                                    }
                                },
                                {
                                    "name": "document_replace_range",
                                    "description": "Replace a byte range in the current document.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {
                                            "start": { "type": "integer", "minimum": 0 },
                                            "end": { "type": "integer", "minimum": 0 },
                                            "text": { "type": "string" }
                                        },
                                        "required": ["start", "end", "text"]
                                    }
                                },
                                {
                                    "name": "selection_set",
                                    "description": "Set the editor selection.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {
                                            "anchor": { "type": "integer", "minimum": 0 },
                                            "focus": { "type": "integer", "minimum": 0 }
                                        },
                                        "required": ["anchor", "focus"]
                                    }
                                },
                                {
                                    "name": "editor_command",
                                    "description": "Run a safe named editor command.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {
                                            "command": {
                                                "type": "string",
                                                "enum": [
                                                    "toggle_bold",
                                                    "toggle_italic",
                                                    "undo",
                                                    "redo",
                                                    "save"
                                                ]
                                            }
                                        },
                                        "required": ["command"]
                                    }
                                },
                                {
                                    "name": "export_pdf",
                                    "description": "Export the document as a PDF to the specified path.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {
                                            "path": { "type": "string" }
                                        },
                                        "required": ["path"]
                                    }
                                },
                                {
                                    "name": "ui_tree",
                                    "description": "Return the visible semantic UI tree with stable IDs.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {}
                                    }
                                },
                                {
                                    "name": "ui_invoke",
                                    "description": "Invoke a safe UI action by stable UI ID.",
                                    "inputSchema": {
                                        "type": "object",
                                        "properties": {
                                            "id": { "type": "string" },
                                            "action": {
                                                "type": "string",
                                                "enum": ["click", "focus", "toggle"]
                                            }
                                        },
                                        "required": ["id", "action"]
                                    }
                                }
                            ]
                        }
                    }))
                }
                "tools/call" => {
                    let tool_name = params.get("name").and_then(|n| n.as_str()).unwrap_or("");
                    let default_args = json!({});
                    let args = params.get("arguments").unwrap_or(&default_args);
                    
                    let (content, is_error) = match tool_name {
                        "editor_state" => {
                            match client.get(format!("{}/state", base_url)).send().await {
                                Ok(resp) => {
                                    let text = resp.text().await.unwrap_or_default();
                                    (text, false)
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "document_get_text" => {
                            match client.get(format!("{}/document/text", base_url)).send().await {
                                Ok(resp) => {
                                    let text = resp.text().await.unwrap_or_default();
                                    (text, false)
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "document_replace_range" => {
                            let start = args.get("start").and_then(|v| v.as_u64()).unwrap_or(0);
                            let end = args.get("end").and_then(|v| v.as_u64()).unwrap_or(0);
                            let text = args.get("text").and_then(|v| v.as_str()).unwrap_or("");
                            let body = json!({
                                "type": "replace_range",
                                "start": start,
                                "end": end,
                                "text": text
                            });
                            match client.post(format!("{}/command", base_url)).json(&body).send().await {
                                Ok(resp) => {
                                    if resp.status().is_success() {
                                        ("Range replaced".to_string(), false)
                                    } else {
                                        let err = resp.text().await.unwrap_or_default();
                                        (err, true)
                                    }
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "selection_set" => {
                            let anchor = args.get("anchor").and_then(|v| v.as_u64()).unwrap_or(0);
                            let focus = args.get("focus").and_then(|v| v.as_u64()).unwrap_or(0);
                            let body = json!({
                                "type": "set_selection",
                                "anchor": anchor,
                                "focus": focus
                            });
                            match client.post(format!("{}/command", base_url)).json(&body).send().await {
                                Ok(resp) => {
                                    if resp.status().is_success() {
                                        ("Selection set".to_string(), false)
                                    } else {
                                        let err = resp.text().await.unwrap_or_default();
                                        (err, true)
                                    }
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "editor_command" => {
                            let cmd = args.get("command").and_then(|v| v.as_str()).unwrap_or("");
                            let body = json!({
                                "type": cmd
                            });
                            match client.post(format!("{}/command", base_url)).json(&body).send().await {
                                Ok(resp) => {
                                    if resp.status().is_success() {
                                        (format!("Executed {}", cmd), false)
                                    } else {
                                        let err = resp.text().await.unwrap_or_default();
                                        (err, true)
                                    }
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "export_pdf" => {
                            let path = args.get("path").and_then(|v| v.as_str()).unwrap_or("");
                            let body = json!({
                                "type": "export_pdf",
                                "path": path
                            });
                            match client.post(format!("{}/command", base_url)).json(&body).send().await {
                                Ok(resp) => {
                                    if resp.status().is_success() {
                                        (format!("Exported PDF to {}", path), false)
                                    } else {
                                        let err = resp.text().await.unwrap_or_default();
                                        (err, true)
                                    }
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "ui_tree" => {
                            match client.get(format!("{}/ui/tree", base_url)).send().await {
                                Ok(resp) => {
                                    let text = resp.text().await.unwrap_or_default();
                                    (text, false)
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        "ui_invoke" => {
                            let body = json!({
                                "id": args.get("id").and_then(|v| v.as_str()).unwrap_or(""),
                                "action": args.get("action").and_then(|v| v.as_str()).unwrap_or("click")
                            });
                            match client.post(format!("{}/ui/invoke", base_url)).json(&body).send().await {
                                Ok(resp) => {
                                    if resp.status().is_success() {
                                        ("UI action invoked".to_string(), false)
                                    } else {
                                        let err = resp.text().await.unwrap_or_default();
                                        (err, true)
                                    }
                                },
                                Err(e) => (e.to_string(), true)
                            }
                        }
                        _ => ("Unknown tool".to_string(), true)
                    };

                    Some(json!({
                        "jsonrpc": "2.0",
                        "id": id,
                        "result": {
                            "content": [
                                {
                                    "type": "text",
                                    "text": content
                                }
                            ],
                            "isError": is_error
                        }
                    }))
                }
                _ => None
            };

            if let Some(res) = response {
                writeln!(handle, "{}", serde_json::to_string(&res).unwrap())?;
                handle.flush()?;
            }
        }
    }

    Ok(())
}
