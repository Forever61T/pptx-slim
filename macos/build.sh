#!/bin/bash
# Build the drag-and-drop macOS app for pptx-slim.
#
# Usage:
#   macos/build.sh                 # -> ~/Applications/PPT Slimmer.app
#   macos/build.sh /path/to/out    # -> /path/to/out/PPT Slimmer.app
#
# Two traps this script exists to avoid:
#   1. osacompile inside a copy of the bundle, or copying the bundle with `cp -R`,
#      leaves resource forks / Finder info that make codesign refuse the bundle.
#      So: build and sign in a scratch dir under /tmp, then move with `ditto`.
#   2. The signature covers Info.plist. Editing Info.plist after signing
#      invalidates it, so it is edited before signing.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(dirname "$HERE")"
APP_NAME="PPT Slimmer"
DEST_DIR="${1:-$HOME/Applications}"
DEST="$DEST_DIR/$APP_NAME.app"

echo "building  $APP_NAME.app"
echo "  engine  $ROOT/pptx_slim.py"

if [ ! -f "$ROOT/pptx_slim.py" ]; then
  echo "error: pptx_slim.py not found next to macos/" >&2
  exit 1
fi

WORK="$(mktemp -d "${TMPDIR:-/tmp}/pptxslim-build.XXXXXX")"
trap 'rm -rf "$WORK"' EXIT

# 1. Compile the AppleScript droplet in the scratch dir.
osacompile -o "$WORK/$APP_NAME.app" "$HERE/droplet.applescript"

# 2. Embed the engine so the app works standalone.
cp "$ROOT/pptx_slim.py" "$WORK/$APP_NAME.app/Contents/Resources/pptx_slim.py"

# 3. Fix up Info.plist (before signing, because the signature covers it).
PLIST="$WORK/$APP_NAME.app/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier local.pptxslim.droplet" "$PLIST" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :CFBundleIdentifier string local.pptxslim.droplet" "$PLIST"
/usr/libexec/PlistBuddy -c "Set :CFBundleName $APP_NAME" "$PLIST" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :CFBundleName string $APP_NAME" "$PLIST"
/usr/libexec/PlistBuddy -c "Set :LSMinimumSystemVersion 10.13" "$PLIST" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :LSMinimumSystemVersion string 10.13" "$PLIST"
# Droplets must advertise that they accept file drops.
/usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes array" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0 dict" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeName string Presentation" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeRole string Viewer" "$PLIST" 2>/dev/null || true

# 4. Strip quarantine/resource-fork cruft, then ad-hoc sign.
xattr -cr "$WORK/$APP_NAME.app"
codesign --force --deep --sign - "$WORK/$APP_NAME.app"

# 5. Move into place with ditto, which preserves the signature.
mkdir -p "$DEST_DIR"
rm -rf "$DEST"
ditto "$WORK/$APP_NAME.app" "$DEST"

codesign --verify --verbose=1 "$DEST" 2>&1 | sed 's/^/  /'
echo "done      $DEST"
