#!/bin/bash
# Build native/hwp2pdf (universal arm64 + x86_64) for export_to_pdf on macOS.
# Needs Xcode command line tools only; it links nothing from 한/글 — the engine is dlopen'd at run time from
# the 한/글 named by HWP_APP, and its libraries resolve from that same app (see hwp2pdf.m). No rpaths here:
# an rpath fixed at build time would load libraries from the builder's install, not the user's.
set -euo pipefail
cd "$(dirname "$0")"
if [ "$(uname)" != Darwin ]; then
  echo "native/build.sh: not macOS, skipping hwp2pdf" >&2
  exit 0
fi
if ! command -v clang >/dev/null; then
  echo "native/build.sh: clang not found (xcode-select --install); export_to_pdf will be unavailable" >&2
  exit 0
fi
clang -fobjc-arc -O2 -arch arm64 -arch x86_64 -mmacosx-version-min=11.0 -o hwp2pdf hwp2pdf.m \
  -framework Foundation -framework AppKit -framework CoreGraphics
echo "built $(pwd)/hwp2pdf ($(lipo -archs hwp2pdf))"
