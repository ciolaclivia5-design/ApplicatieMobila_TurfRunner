#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$ROOT_DIR/backend"
APP_DIR="$ROOT_DIR/turf_runner"

for tool in node npm xcrun osascript; do
  if ! command -v "$tool" >/dev/null 2>&1; then
    echo "Required tool not found: $tool"
    echo "Install Xcode Command Line Tools and Node.js, then run this script again."
    read -r -p "Press Enter to close..."
    exit 1
  fi
done

if [ ! -x "$BACKEND_DIR/node_modules/.bin/ts-node-dev" ]; then
  echo "Installing backend dependencies..."
  (cd "$BACKEND_DIR" && npm install)
fi

if [ ! -x "$APP_DIR/node_modules/.bin/react-native" ]; then
  echo "Installing app dependencies..."
  (cd "$APP_DIR" && npm install)
fi

if [ ! -d "$APP_DIR/ios/Pods" ]; then
  if ! command -v bundle >/dev/null 2>&1; then
    echo "Bundler is required to install iOS dependencies. Install Ruby Bundler and try again."
    read -r -p "Press Enter to close..."
    exit 1
  fi
  echo "Installing iOS native dependencies..."
  (cd "$APP_DIR" && bundle install)
  (cd "$APP_DIR/ios" && bundle exec pod install)
fi

DEVELOPER_DIR="$(xcode-select -p)"
SIMULATOR_APP="$DEVELOPER_DIR/Applications/Simulator.app"
DEVICE_HUB_APP="$DEVELOPER_DIR/../Applications/DeviceHub.app"
if [ ! -d "$SIMULATOR_APP" ] && [ ! -d "$DEVICE_HUB_APP" ]; then
  echo "Neither Simulator.app nor DeviceHub.app was found in the selected Xcode installation."
  read -r -p "Press Enter to close..."
  exit 1
fi

osascript - "$ROOT_DIR" <<'APPLESCRIPT'
on run argv
  set projectRoot to item 1 of argv
  set backendCommand to "cd " & quoted form of (projectRoot & "/backend") & " && npm start"
  set appCommand to "cd " & quoted form of (projectRoot & "/turf_runner") & " && PATH=" & quoted form of (projectRoot & "/scripts/mac-compat") & ":$PATH npx react-native run-ios"

  tell application "Terminal"
    activate
    do script backendCommand
    delay 1
    do script appCommand
  end tell
end run
APPLESCRIPT

echo "Started the backend and iOS app in Terminal. Stop each process with Control+C in its Terminal window."
