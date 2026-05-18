#!/usr/bin/env bash
set -euo pipefail

REPO_OWNER="${REPO_OWNER:-DjakeDjone}"
REPO_NAME="${REPO_NAME:-lexica-doc}"
BRANCH="${BRANCH:-main}"
WORKFLOW_FILE="${WORKFLOW_FILE:-build.yml}"
INSTALL_ROOT="${INSTALL_ROOT:-$HOME/.local}"
BIN_NAME="wors"
ARTIFACT_NAME="${ARTIFACT_NAME:-wors-linux-x86_64}"

if ! command -v curl >/dev/null 2>&1; then
  echo "error: curl is required" >&2
  exit 1
fi

if ! command -v unzip >/dev/null 2>&1; then
  echo "error: unzip is required" >&2
  exit 1
fi

if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  echo "error: python3 or python is required" >&2
  exit 1
fi

api_root="https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}"
github_token="${GH_TOKEN:-${GITHUB_TOKEN:-}}"
curl_progress=(-fL --progress-bar)
tmp_dir="$(mktemp -d)"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT

run_json="${tmp_dir}/runs.json"
artifacts_json="${tmp_dir}/artifacts.json"
artifact_zip="${tmp_dir}/${ARTIFACT_NAME}.zip"
extract_dir="${tmp_dir}/artifact"
install_bin_dir="${INSTALL_ROOT}/bin"

mkdir -p "$extract_dir"

if command -v gh >/dev/null 2>&1; then
  echo "Finding latest successful ${WORKFLOW_FILE} run on ${BRANCH}"
  run_id="$(gh run list \
    --repo "${REPO_OWNER}/${REPO_NAME}" \
    --workflow "$WORKFLOW_FILE" \
    --branch "$BRANCH" \
    --status success \
    --event push \
    --limit 1 \
    --json databaseId \
    --jq '.[0].databaseId')"

  if [[ -z "$run_id" || "$run_id" == "null" ]]; then
    echo "error: no successful workflow runs found" >&2
    exit 1
  fi

  echo "Downloading ${ARTIFACT_NAME}"
  artifact_url="$(gh api \
    "repos/${REPO_OWNER}/${REPO_NAME}/actions/runs/${run_id}/artifacts" \
    --jq ".artifacts[] | select(.name == \"${ARTIFACT_NAME}\" and .expired == false) | .archive_download_url" \
    | head -n 1)"

  if [[ -z "$artifact_url" ]]; then
    echo "error: artifact '${ARTIFACT_NAME}' not found or has expired" >&2
    exit 1
  fi

  curl "${curl_progress[@]}" \
    -H "Accept: application/vnd.github+json" \
    -H "Authorization: Bearer $(gh auth token)" \
    "$artifact_url" \
    -o "$artifact_zip"
  unzip -q "$artifact_zip" -d "$extract_dir"
else
  if [[ -z "$github_token" ]]; then
    echo "error: install-prebuilt.sh requires gh or GH_TOKEN/GITHUB_TOKEN to download GitHub Actions artifacts" >&2
    exit 1
  fi

  curl_headers=(
    -H "Accept: application/vnd.github+json"
    -H "Authorization: Bearer ${github_token}"
  )

  runs_url="${api_root}/actions/workflows/${WORKFLOW_FILE}/runs?branch=${BRANCH}&status=success&event=push&per_page=1"
  echo "Finding latest successful ${WORKFLOW_FILE} run on ${BRANCH}"
  curl -fsSL "${curl_headers[@]}" "$runs_url" -o "$run_json"

  artifacts_url="$("$PYTHON_BIN" - "$run_json" <<'PY'
import json
import sys

with open(sys.argv[1], "r", encoding="utf-8") as handle:
    runs = json.load(handle).get("workflow_runs", [])

if not runs:
    sys.exit("error: no successful workflow runs found")

print(runs[0]["artifacts_url"])
PY
  )"

  curl -fsSL "${curl_headers[@]}" "$artifacts_url" -o "$artifacts_json"

  download_url="$("$PYTHON_BIN" - "$artifacts_json" "$ARTIFACT_NAME" <<'PY'
import json
import sys

with open(sys.argv[1], "r", encoding="utf-8") as handle:
    artifacts = json.load(handle).get("artifacts", [])

artifact_name = sys.argv[2]
for artifact in artifacts:
    if artifact.get("name") == artifact_name and not artifact.get("expired", False):
        print(artifact["archive_download_url"])
        break
else:
    sys.exit(f"error: artifact {artifact_name!r} not found or has expired")
PY
  )"

  echo "Downloading ${ARTIFACT_NAME}"
  curl "${curl_progress[@]}" "${curl_headers[@]}" "$download_url" -o "$artifact_zip"
  unzip -q "$artifact_zip" -d "$extract_dir"
fi

binary_path="$(find "$extract_dir" -type f -name "$BIN_NAME" -perm /111 | head -n 1)"
if [[ -z "$binary_path" ]]; then
  binary_path="$(find "$extract_dir" -type f -name "$BIN_NAME" | head -n 1)"
fi

if [[ -z "$binary_path" ]]; then
  echo "error: ${BIN_NAME} not found in ${ARTIFACT_NAME}" >&2
  exit 1
fi

mkdir -p "$install_bin_dir"
install -m 755 "$binary_path" "${install_bin_dir}/${BIN_NAME}"

# Install .desktop file and icon for file association (Open With)
XDG_DATA_HOME="${XDG_DATA_HOME:-$HOME/.local/share}"
DESKTOP_DIR="${XDG_DATA_HOME}/applications"
ICON_DIR="${XDG_DATA_HOME}/icons/hicolor/256x256/apps"
DESKTOP_FILE="${DESKTOP_DIR}/${BIN_NAME}.desktop"
ICON_FILE="${ICON_DIR}/${BIN_NAME}.png"

mkdir -p "$DESKTOP_DIR" "$ICON_DIR"

# Copy logo from extracted artifact if present
if [[ -f "${extract_dir}/logo.png" ]]; then
  cp "${extract_dir}/logo.png" "$ICON_FILE"
  chmod 644 "$ICON_FILE"
fi

cat >"$DESKTOP_FILE" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=Wors
Comment=Minimal desktop document editor
Exec=${install_bin_dir}/${BIN_NAME} %f
TryExec=${install_bin_dir}/${BIN_NAME}
Icon=${ICON_FILE}
Terminal=false
Categories=Office;WordProcessor;
MimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/vnd.oasis.opendocument.text;text/markdown;text/plain;
StartupNotify=true
EOF

chmod 644 "$DESKTOP_FILE"

if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database "$DESKTOP_DIR" >/dev/null 2>&1 || true
fi

if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q "${XDG_DATA_HOME}/icons/hicolor" >/dev/null 2>&1 || true
fi

echo "Installed ${BIN_NAME} to ${install_bin_dir}/${BIN_NAME}."
if [[ ":${PATH}:" != *":${install_bin_dir}:"* ]]; then
  echo "Add ${install_bin_dir} to your PATH to run ${BIN_NAME}."
fi
