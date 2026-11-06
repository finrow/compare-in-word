#!/bin/sh
# uninstall.sh — removes the Compare in Word Finder action and support files.
set -e
rm -rf "$HOME/Library/Services/Compare in Word.workflow"
rm -rf "$HOME/Library/Application Support/CompareInWord"
/System/Library/CoreServices/pbs -flush >/dev/null 2>&1 || true
/System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true
echo "Compare in Word removed."
