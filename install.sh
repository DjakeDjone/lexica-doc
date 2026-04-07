#!/usr/bin/env bash
set -euo pipefail

REPO_OWNER="${REPO_OWNER:-DjakeDjone}"
REPO_NAME="${REPO_NAME:-lexica-doc}"
BRANCH="${BRANCH:-main}"
INSTALL_ROOT="${INSTALL_ROOT:-$HOME/.local}"
XDG_DATA_HOME="${XDG_DATA_HOME:-$HOME/.local/share}"
BIN_NAME="wors"
APP_NAME="Wors"
DESKTOP_DIR="${XDG_DATA_HOME}/applications"
ICON_DIR="${XDG_DATA_HOME}/icons/hicolor/256x256/apps"
DESKTOP_FILE="${DESKTOP_DIR}/${BIN_NAME}.desktop"
ICON_FILE="${ICON_DIR}/${BIN_NAME}.png"

if ! command -v curl >/dev/null 2>&1; then
  echo "error: curl is required" >&2
  exit 1
fi

if ! command -v tar >/dev/null 2>&1; then
  echo "error: tar is required" >&2
  exit 1
fi

if ! command -v cargo >/dev/null 2>&1; then
  echo "error: cargo is required (install Rust from https://rustup.rs)" >&2
  exit 1
fi

tmp_dir="$(mktemp -d)"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT

archive_url="https://github.com/${REPO_OWNER}/${REPO_NAME}/archive/refs/heads/${BRANCH}.tar.gz"
archive_file="${tmp_dir}/source.tar.gz"
source_root="${tmp_dir}/${REPO_NAME}-${BRANCH}"
crate_dir="$source_root"

echo "Downloading ${archive_url}"
curl -fsSL "$archive_url" -o "$archive_file"
tar -xzf "$archive_file" -C "$tmp_dir"

if [[ ! -f "${crate_dir}/Cargo.toml" && -f "${source_root}/browser/Cargo.toml" ]]; then
  crate_dir="${source_root}/browser"
fi

if [[ ! -f "${crate_dir}/Cargo.toml" ]]; then
  echo "error: Cargo.toml not found in downloaded source" >&2
  exit 1
fi

echo "Installing ${BIN_NAME} to ${INSTALL_ROOT}/bin"
cargo install --path "$crate_dir" --locked --force --root "$INSTALL_ROOT"

mkdir -p "$DESKTOP_DIR" "$ICON_DIR"

if [[ -f "${crate_dir}/assets/logo.png" ]]; then
  cp "${crate_dir}/assets/logo.png" "$ICON_FILE"
else
  echo "warning: ${crate_dir}/assets/logo.png not found; launcher icon will not be updated" >&2
fi

cat >"$DESKTOP_FILE" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=${APP_NAME}
Comment=Minimal desktop document editor
Exec=${INSTALL_ROOT}/bin/${BIN_NAME}
TryExec=${INSTALL_ROOT}/bin/${BIN_NAME}
Icon=${ICON_FILE}
Terminal=false
Categories=Office;WordProcessor;
StartupNotify=true
EOF

chmod 644 "$DESKTOP_FILE"
if [[ -f "$ICON_FILE" ]]; then
  chmod 644 "$ICON_FILE"
fi

if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database "$DESKTOP_DIR" >/dev/null 2>&1 || true
fi

if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q "${XDG_DATA_HOME}/icons/hicolor" >/dev/null 2>&1 || true
fi

echo "Installed ${BIN_NAME} and desktop launcher."
if [[ ":${PATH}:" != *":${INSTALL_ROOT}/bin:"* ]]; then
  echo "Add ${INSTALL_ROOT}/bin to your PATH to run ${BIN_NAME}."
fi
echo "You can launch ${APP_NAME} from your desktop app menu."
