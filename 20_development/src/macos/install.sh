#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_SCRIPT="${SCRIPT_DIR}/BKTKeyState.applescript"
COMPILED_SCRIPT="${SCRIPT_DIR}/BKTKeyState.scpt"
TARGET_DIR="${HOME}/Library/Application Scripts/com.microsoft.Powerpoint"
TARGET_SCRIPT="${TARGET_DIR}/BKTKeyState.scpt"

if ! command -v osacompile >/dev/null 2>&1; then
  echo "Error: osacompile is not available on this Mac." >&2
  exit 1
fi

if [[ ! -f "${SOURCE_SCRIPT}" ]]; then
  echo "Error: ${SOURCE_SCRIPT} was not found." >&2
  exit 1
fi

mkdir -p "${TARGET_DIR}"

echo "Compiling ${SOURCE_SCRIPT}..."
rm -f "${COMPILED_SCRIPT}"
osacompile -o "${COMPILED_SCRIPT}" "${SOURCE_SCRIPT}"

echo "Installing ${TARGET_SCRIPT}..."
cp -f "${COMPILED_SCRIPT}" "${TARGET_SCRIPT}"

echo "Installed BKTKeyState.scpt to ${TARGET_DIR}"
