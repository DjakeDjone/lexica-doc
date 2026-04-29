# wors

Minimal desktop document editor built with Rust + `eframe/egui`.

## Install

Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install.sh | bash
```

The installer downloads the source and installs `wors` to `$HOME/.local/bin`.
It also registers a desktop launcher (`wors.desktop`) so the app appears in your system menu.

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install.ps1 | iex
```

The Windows installer downloads the source, builds it with Cargo, installs `wors.exe` to `%USERPROFILE%\.cargo\bin`, adds that directory to your user `PATH` if needed, and creates a Start Menu shortcut.

## Run

```bash
wors
```

## Run In The Browser

Install Trunk if needed:

```bash
cargo install trunk
rustup target add wasm32-unknown-unknown
```

Serve the editor locally:

```bash
trunk serve --open
```

The web build uses the same editor UI as the desktop app. Desktop-only integrations such as native file dialogs and the local LanguageTool process are disabled in the browser build.

## Todos

- [ ] Add support for opening files from the command line
- [ ] desktop icon
- [ ] more formatting options
- [ ] export to PDF
- *and more...*
