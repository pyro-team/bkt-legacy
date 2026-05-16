#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
STABLE_DIR="${ROOT_DIR}/30_builds/10_stable"
MACOS_DIR="${ROOT_DIR}/20_development/src/macos"
DIST_DIR="${ROOT_DIR}/30_builds/mac"
PACKAGE_NAME="BKT-Legacy-mac"
PACKAGE_DIR="${DIST_DIR}/${PACKAGE_NAME}"
ZIP_FILE="${DIST_DIR}/${PACKAGE_NAME}.zip"

require_file() {
  local path="$1"
  if [[ ! -f "${path}" ]]; then
    echo "Missing required file: ${path}" >&2
    exit 1
  fi
}

require_file "${STABLE_DIR}/BKT-Legacy.ppam"
require_file "${STABLE_DIR}/Templates.pptx"
require_file "${MACOS_DIR}/BKTKeyState.applescript"
require_file "${MACOS_DIR}/install.command"
require_file "${MACOS_DIR}/README-Mac.md"

if ! command -v osacompile >/dev/null 2>&1; then
  echo "Error: osacompile is not available on this Mac." >&2
  exit 1
fi

echo "Compiling BKTKeyState.scpt..."
osacompile -o "${MACOS_DIR}/BKTKeyState.scpt" "${MACOS_DIR}/BKTKeyState.applescript"

echo "Preparing ${PACKAGE_DIR}..."
rm -rf "${PACKAGE_DIR}"
mkdir -p "${PACKAGE_DIR}"

cp "${STABLE_DIR}/BKT-Legacy.ppam" "${PACKAGE_DIR}/"
cp "${STABLE_DIR}/Templates.pptx" "${PACKAGE_DIR}/"
cp "${MACOS_DIR}/BKTKeyState.scpt" "${PACKAGE_DIR}/"
cp "${MACOS_DIR}/install.command" "${PACKAGE_DIR}/"
cp "${MACOS_DIR}/README-Mac.md" "${PACKAGE_DIR}/"

chmod +x "${PACKAGE_DIR}/install.command"

echo "Creating ${ZIP_FILE}..."
rm -f "${ZIP_FILE}"
(
  cd "${DIST_DIR}"
  zip -r "${ZIP_FILE}" "${PACKAGE_NAME}" >/dev/null
)

echo "Done:"
echo "  ${PACKAGE_DIR}"
echo "  ${ZIP_FILE}"
