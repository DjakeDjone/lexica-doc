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


## Todos

- [ ] Add support for opening files from the command line
- [ ] desktop icon
- [ ] more formatting options
- [ ] export to PDF
- *and more...*
