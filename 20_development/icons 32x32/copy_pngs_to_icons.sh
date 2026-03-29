#!/usr/bin/env bash

set -euo pipefail

script_dir="$(cd "$(dirname "$0")" && pwd)"
target_dir="${script_dir%/icons 32x32}/icons"

if [[ ! -d "$target_dir" ]]; then
  echo "Target folder not found: $target_dir" >&2
  exit 1
fi

shopt -s nullglob
png_files=("$script_dir"/*.png)

if [[ ${#png_files[@]} -eq 0 ]]; then
  echo "No PNG files found in: $script_dir"
  exit 0
fi

for file in "${png_files[@]}"; do
  target_file="$target_dir/$(basename "$file")"
  sips -z 16 16 \
    --setProperty dpiWidth 192 \
    --setProperty dpiHeight 192 \
    "$file" \
    --out "$target_file" >/dev/null
  echo "Converted $(basename "$file") -> $(basename "$target_file")"
done
