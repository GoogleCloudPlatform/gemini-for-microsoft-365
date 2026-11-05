#!/usr/bin/env bash
#
# Copyright 2026 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

# Sideloads Gemini Office 365 Add-in manifests to macOS Office sandbox containers
# Supports Word, PowerPoint, Excel, and Outlook.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
MANIFEST_ARG="${1:-}"

resolve_path() {
  local target="$1"
  if [ -z "$target" ]; then
    echo ""
  elif [ -f "$target" ]; then
    echo "$target"
  elif [ -f "$REPO_ROOT/$target" ]; then
    echo "$REPO_ROOT/$target"
  elif [ -f "$REPO_ROOT/microsoft-addin/$target" ]; then
    echo "$REPO_ROOT/microsoft-addin/$target"
  else
    echo ""
  fi
}

OFFICE_MANIFEST=""
OUTLOOK_MANIFEST=""

if [ -n "$MANIFEST_ARG" ]; then
  RESOLVED_ARG="$(resolve_path "$MANIFEST_ARG")"
  if [ -z "$RESOLVED_ARG" ]; then
    echo "❌ Error: Specified manifest file not found: $MANIFEST_ARG"
    echo "   Usage: $0 [path/to/manifest.xml] [path/to/manifest-outlook.xml]"
    exit 1
  fi
  OFFICE_MANIFEST="$RESOLVED_ARG"

  # Optional second argument for Outlook manifest
  if [ -n "${2:-}" ]; then
    OUTLOOK_MANIFEST="$(resolve_path "$2")"
  else
    # Auto-derive sibling Outlook manifest if present
    DERIVED_OUTLOOK="${OFFICE_MANIFEST%.xml}-outlook.xml"
    if [ -f "$DERIVED_OUTLOOK" ]; then
      OUTLOOK_MANIFEST="$DERIVED_OUTLOOK"
    fi
  fi
else
  # Auto-discover default Office manifest
  OFFICE_CANDIDATES=(
    "$REPO_ROOT/microsoft-addin/manifest-gemini.xml"
    "$REPO_ROOT/manifest-gemini.xml"
    "$REPO_ROOT/microsoft-addin/manifest-dev.xml"
    "$REPO_ROOT/manifest-wif.xml"
    "$REPO_ROOT/manifest.xml"
  )
  for c in "${OFFICE_CANDIDATES[@]}"; do
    if [ -f "$c" ]; then
      OFFICE_MANIFEST="$c"
      break
    fi
  done

  # Auto-discover default Outlook manifest
  OUTLOOK_CANDIDATES=(
    "$REPO_ROOT/microsoft-addin/manifest-gemini-outlook.xml"
    "$REPO_ROOT/manifest-gemini-outlook.xml"
    "$REPO_ROOT/microsoft-addin/manifest-outlook.xml"
  )
  for c in "${OUTLOOK_CANDIDATES[@]}"; do
    if [ -f "$c" ]; then
      OUTLOOK_MANIFEST="$c"
      break
    fi
  done
fi

if [ -z "$OFFICE_MANIFEST" ] || [ ! -f "$OFFICE_MANIFEST" ]; then
  echo "❌ Error: Office manifest file not found. Please provide path to your manifest XML file:"
  echo "   Usage: $0 [path/to/manifest.xml] [path/to/manifest-outlook.xml]"
  exit 1
fi

echo "🚀 Installing Gemini for Office 365 Add-in on macOS..."
echo "   • Office Manifest (Word, PowerPoint, Excel): $(basename "$OFFICE_MANIFEST")"
if [ -n "$OUTLOOK_MANIFEST" ] && [ -f "$OUTLOOK_MANIFEST" ]; then
  echo "   • Outlook Manifest:                          $(basename "$OUTLOOK_MANIFEST")"
fi

OFFICE_APPS=("com.microsoft.Word" "com.microsoft.Powerpoint" "com.microsoft.Excel")

for app in "${OFFICE_APPS[@]}"; do
  WEF_DIR="$HOME/Library/Containers/$app/Data/Documents/wef"
  echo "📁 Target directory ($app): $WEF_DIR"
  mkdir -p "$WEF_DIR"
  rm -rf "$WEF_DIR"/*
  cp -f "$OFFICE_MANIFEST" "$WEF_DIR/gemini-manifest.xml"
  echo "   ✅ Office manifest copied to $WEF_DIR/gemini-manifest.xml"

  # Clear WebKit Network and Storage caches for the app
  CACHE_DIR="$HOME/Library/Containers/$app/Data/Library/Caches/WebKit"
  if [ -d "$CACHE_DIR" ]; then
    rm -rf "$CACHE_DIR/NetworkCache"/* "$CACHE_DIR/CacheStorage"/* 2>/dev/null || true
    echo "   🧹 Cleared WebKit cache for $app"
  fi

  # Enable Web Inspector / Developer Extras
  defaults write "$app" OfficeWebAddinDeveloperExtras -bool true 2>/dev/null || true
  echo "   🛠️  Enabled OfficeWebAddinDeveloperExtras for $app"
done

# Outlook host container
OUTLOOK_APP="com.microsoft.Outlook"
OUTLOOK_WEF_DIR="$HOME/Library/Containers/$OUTLOOK_APP/Data/Documents/wef"
echo "📁 Target directory ($OUTLOOK_APP): $OUTLOOK_WEF_DIR"
mkdir -p "$OUTLOOK_WEF_DIR"
rm -rf "$OUTLOOK_WEF_DIR"/*

if [ -n "$OUTLOOK_MANIFEST" ] && [ -f "$OUTLOOK_MANIFEST" ]; then
  cp -f "$OUTLOOK_MANIFEST" "$OUTLOOK_WEF_DIR/gemini-manifest.xml"
  echo "   ✅ Outlook manifest copied to $OUTLOOK_WEF_DIR/gemini-manifest.xml"
else
  # Fallback to primary manifest if Outlook-specific manifest doesn't exist
  cp -f "$OFFICE_MANIFEST" "$OUTLOOK_WEF_DIR/gemini-manifest.xml"
  echo "   ℹ️ Copied primary manifest to $OUTLOOK_WEF_DIR/gemini-manifest.xml"
fi

# Clear WebKit Cache & enable dev extras for Outlook
OUTLOOK_CACHE_DIR="$HOME/Library/Containers/$OUTLOOK_APP/Data/Library/Caches/WebKit"
if [ -d "$OUTLOOK_CACHE_DIR" ]; then
  rm -rf "$OUTLOOK_CACHE_DIR/NetworkCache"/* "$OUTLOOK_CACHE_DIR/CacheStorage"/* 2>/dev/null || true
  echo "   🧹 Cleared WebKit cache for $OUTLOOK_APP"
fi
defaults write "$OUTLOOK_APP" OfficeWebAddinDeveloperExtras -bool true 2>/dev/null || true
echo "   🛠️  Enabled OfficeWebAddinDeveloperExtras for $OUTLOOK_APP"

echo ""
echo "✨ Sideloading complete!"
echo "👉 Next steps:"
echo "   1. Restart Microsoft Word, PowerPoint, Excel, or Outlook."
echo "   2. Navigate to the 'Home' tab on the ribbon."
echo "   3. Click 'Gemini' in the ribbon."
echo "   (Tip: Right-click anywhere in the taskpane and choose 'Inspect Element' to view console logs)"
