#!/usr/bin/env bash

set -euo pipefail

SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT_TARGET_DIR="${HOME}/Library/Application Scripts/com.microsoft.Powerpoint"
ADDIN_TARGET_DIR="${HOME}/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins"
ADDIN_SOURCE_PATH="${SOURCE_DIR}/BKT-Legacy.ppam"
ADDIN_INSTALL_PATH="${ADDIN_TARGET_DIR}/BKT-Legacy.ppam"

fail() {
  echo "Error: $1" >&2
  echo
  read -r -p "Press Enter to close this window..."
  exit 1
}

require_file() {
  local path="$1"
  [[ -f "${path}" ]] || fail "Missing required file: ${path}"
}

require_file "${SOURCE_DIR}/BKT-Legacy.ppam"
require_file "${SOURCE_DIR}/Templates.pptx"
require_file "${SOURCE_DIR}/BKTKeyState.scpt"

echo "BKT Legacy Mac installer"
echo
echo "Choose where PowerPoint should load the add-in from:"
echo "1. Copy it to the Microsoft Office Add-Ins folder (recommended)"
echo "2. Keep it in the current folder"
echo

while true; do
  read -r -p "Enter 1 or 2 [1]: " install_choice
  install_choice="${install_choice:-1}"

  case "${install_choice}" in
    1)
      ADDIN_LOAD_PATH="${ADDIN_INSTALL_PATH}"
      COPY_ADDIN_FILES=true
      break
      ;;
    2)
      ADDIN_LOAD_PATH="${ADDIN_SOURCE_PATH}"
      COPY_ADDIN_FILES=false
      break
      ;;
    *)
      echo "Please enter 1 or 2."
      ;;
  esac
done

mkdir -p "${SCRIPT_TARGET_DIR}"

echo "Installing BKTKeyState.scpt..."
cp -f "${SOURCE_DIR}/BKTKeyState.scpt" "${SCRIPT_TARGET_DIR}/BKTKeyState.scpt"

if [[ "${COPY_ADDIN_FILES}" == true ]]; then
  mkdir -p "${ADDIN_TARGET_DIR}"

  echo "Installing BKT-Legacy.ppam..."
  cp -f "${SOURCE_DIR}/BKT-Legacy.ppam" "${ADDIN_TARGET_DIR}/BKT-Legacy.ppam"

  echo "Installing Templates.pptx..."
  cp -f "${SOURCE_DIR}/Templates.pptx" "${ADDIN_TARGET_DIR}/Templates.pptx"
else
  echo "Keeping BKT-Legacy.ppam in the current folder."
fi

echo
echo "BKT Legacy files were installed."
if [[ "${COPY_ADDIN_FILES}" == false ]]; then
  echo "Keep this folder in place. Templates.pptx must stay next to BKT-Legacy.ppam."
fi
echo
echo "Next steps:"
echo "1. Open PowerPoint."
echo "2. Go to Tools > PowerPoint Add-ins."
echo "3. Click + and choose BKT-Legacy.ppam from:"
echo "   ${ADDIN_LOAD_PATH}"
echo "4. If the file picker does not show the folder, press Command-Shift-G and paste that path."
echo "5. Confirm the macro security prompts."
echo
read -r -p "Press Enter to close this window..."
