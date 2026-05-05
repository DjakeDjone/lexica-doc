# wors

Minimal desktop document editor built with Rust + `eframe/egui`.

## Install

### Prebuilt Binaries

The GitHub workflow builds Linux and Windows desktop binaries on each push to `main`,
for pull requests, and when run manually.

Install the latest successful `main` build without downloading the artifact manually:

Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install-prebuilt.sh | bash
```

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install-prebuilt.ps1 | iex
```

The prebuilt installer downloads the latest non-expired GitHub Actions artifact for your
platform and installs `wors` to a local user binary directory. GitHub requires
authenticated artifact downloads, so run `gh auth login` first or set `GH_TOKEN` /
`GITHUB_TOKEN`.

To install manually instead:

1. Open the latest successful **Build desktop app** run in GitHub Actions.
2. Download the artifact for your platform:
   - `wors-linux-x86_64`
   - `wors-windows-x86_64`
3. Extract the downloaded `.zip` archive.

Linux:

```bash
unzip wors-linux-x86_64.zip -d wors-linux-x86_64
chmod +x wors-linux-x86_64/wors
mkdir -p "$HOME/.local/bin"
mv wors-linux-x86_64/wors "$HOME/.local/bin/wors"
```

Windows (PowerShell):

```powershell
Expand-Archive .\wors-windows-x86_64.zip -DestinationPath .
.\wors-windows-x86_64\wors.exe
```

Move `wors.exe` to a directory on your `PATH` if you want to run it from any terminal.

### Build From Source Installer

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
