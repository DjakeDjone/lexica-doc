#!/usr/bin/env bash
set -euo pipefail

REPO_OWNER="${REPO_OWNER:-DjakeDjone}"
REPO_NAME="${REPO_NAME:-lexica-doc}"
BRANCH="${BRANCH:-main}"
INSTALL_ROOT="${INSTALL_ROOT:-$HOME/.local}"
BIN_NAME="wors"

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

echo "Installed ${BIN_NAME}."
if [[ ":${PATH}:" != *":${INSTALL_ROOT}/bin:"* ]]; then
  echo "Add ${INSTALL_ROOT}/bin to your PATH to run ${BIN_NAME}."
fi
