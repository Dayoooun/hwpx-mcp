#!/bin/bash
# Build native/hwp2pdf (universal arm64 + x86_64) for export_to_pdf on macOS.
# Needs Xcode command line tools only; it links nothing from 한/글 at build time — the engine is
# dlopen'd at run time from the user's installed app (rpaths below cover its dependent libraries).
set -euo pipefail
cd "$(dirname "$0")"
if [ "$(uname)" != Darwin ]; then
  echo "native/build.sh: not macOS, skipping hwp2pdf" >&2
  exit 0
fi
APP="${HWP_APP:-/Applications/Hancom Office HWP.app}"
clang -fobjc-arc -O2 -arch arm64 -arch x86_64 -mmacosx-version-min=11.0 -o hwp2pdf hwp2pdf.m \
  -framework Foundation -framework AppKit -framework CoreGraphics \
  -Wl,-rpath,"$APP/Contents/Frameworks/Hnc/Bin" -Wl,-rpath,"$APP/Contents/Frameworks"
echo "built $(pwd)/hwp2pdf ($(lipo -archs hwp2pdf))"
