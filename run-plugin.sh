#!/usr/bin/env bash

osascript <<'EOF'
if application "Figma" is running then
    tell application "Figma" to activate
    tell application "System Events" to tell process "Figma"
        keystroke "p" using {command down, option down}
    end tell
    delay 0.2
    tell application "Visual Studio Code" to activate
    return "Running"
else
    return "Not running"
end if
EOF