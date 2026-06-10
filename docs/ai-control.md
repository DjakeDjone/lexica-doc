# AI Control Layer (MCP)

`wors` exposes an AI-controllable API to allow local intelligent agents (such as Antigravity and Codex) to read the editor state, document content, semantic UI tree, and invoke actions.

This document describes how the AI control layer works, how to configure your AI tools to use it, and the security model around it.

## Architecture

The AI control layer consists of two parts:
1. **Local HTTP Editor API**: When `wors` runs, it starts a local background Axum HTTP server bound to `127.0.0.1:0`. The port it binds to is written to `~/.wors/mcp-port`.
2. **Editor MCP Server**: A standalone binary `editor-mcp-server` that communicates over stdin/stdout using the Model Context Protocol (MCP). It reads the port from `~/.wors/mcp-port` and proxies JSON-RPC MCP calls as HTTP requests to the editor API.

This design cleanly decouples the egui editor process from the standard MCP stdio communication protocol.

## Features

The MCP server exposes the following tools to the AI:
- `editor_state`: Returns current editor metadata (title, selection range, undo/redo availability).
- `document_get_text`: Retrieves the entire plaintext contents of the document.
- `document_replace_range`: Replaces a specific byte range in the text with new text.
- `selection_set`: Sets the anchor and focus points of the cursor selection.
- `editor_command`: Executes safe predefined commands (e.g., `toggle_bold`, `undo`, `redo`).
- `ui_tree`: Returns an accessibility-style semantic UI tree with stable element IDs.
- `ui_invoke`: Triggers an action (like `click`) on a given UI tree element by ID.

## Setup for Antigravity

To configure Antigravity or Codex to use the `editor-mcp-server`, add the following to your MCP configuration file (typically in `~/.gemini/antigravity/mcp/` or your editor's MCP config):

```json
{
  "mcpServers": {
    "wors-editor": {
      "command": "/path/to/vibed/browser/target/debug/editor-mcp-server",
      "args": []
    }
  }
}
```

Make sure that `wors` is running before the AI attempts to use these tools, otherwise the MCP server will fail to find `~/.wors/mcp-port`.

## Security Model

Allowing an external process to arbitrarily mutate an open document or run commands can be a security risk. The `wors` AI control layer mitigates this through several design choices:

1. **Local-Only Binding**: The internal HTTP API strictly binds to `127.0.0.1` and will reject or ignore remote requests. It relies on the OS user boundary for protection.
2. **Strictly Typed Commands**: The API only accepts a predefined set of semantic commands (via the `EditorCommand` enum). It does not permit arbitrary memory reads, arbitrary shell execution, or arbitrary file system access outside of normal editor export flows.
3. **Range Validation**: Operations like `ReplaceRange` perform boundary checks (e.g., ensuring ranges are valid indices in the document) to prevent out-of-bounds crashes or memory exploits.
4. **No Raw Pointer Actions**: Rather than allowing AI to synthesize raw mouse clicks or keypresses at specific screen coordinates (which can be fragile and exploited to click other applications), it only permits invoking semantic UI elements by their stable ID.

### Future Work
If needed, a future update may add a low-level automation layer (e.g., synthesizing OS-level key presses and generic screenshots). For security reasons, this would be hidden behind an explicit CLI permission flag (e.g., `--unsafe-ai-automation`).
