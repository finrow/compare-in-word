#!/bin/sh
#
# install.sh — installs Compare in Word for the current user.
#
#   1. copies the scripts to ~/Library/Application Support/CompareInWord;
#   2. installs the Finder Quick Action into ~/Library/Services.
#
# Re-run it any time to update. settings.conf is never overwritten.

set -e
REPO=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
SUPPORT="$HOME/Library/Application Support/CompareInWord"
SERVICES="$HOME/Library/Services"

say() { printf '\033[1m%s\033[0m\n' "$1"; }

say "Installing support files to $SUPPORT"
mkdir -p "$SUPPORT"
for f in compare-in-word.sh compare-in-word.applescript show-dialog.applescript; do
    cp "$REPO/support/$f" "$SUPPORT/$f"
done
chmod +x "$SUPPORT/compare-in-word.sh"
cp "$REPO/README.md" "$SUPPORT/README.md" 2>/dev/null || true
if [ -f "$SUPPORT/settings.conf" ]; then
    cp "$REPO/support/settings.conf" "$SUPPORT/settings.conf.default"
else
    cp "$REPO/support/settings.conf" "$SUPPORT/settings.conf"
fi
# Files copied from a downloaded zip carry a quarantine flag; clear it.
xattr -dr com.apple.quarantine "$SUPPORT" 2>/dev/null || true

say "Installing the Finder action to $SERVICES"
mkdir -p "$SERVICES"
dest="$SERVICES/Compare in Word.workflow"
rm -rf "$dest"
mkdir -p "$dest/Contents/Resources"
cp "$REPO/services/template/Contents/Info.plist" "$dest/Contents/Info.plist"
u1=$(uuidgen); u2=$(uuidgen); u3=$(uuidgen)
sed -e "s/__UUID1__/$u1/g" -e "s/__UUID2__/$u2/g" -e "s/__UUID3__/$u3/g" \
    "$REPO/services/template/Contents/Resources/document.wflow" > "$dest/Contents/Resources/document.wflow"
plutil -convert binary1 "$dest/Contents/Resources/document.wflow" >/dev/null 2>&1 || true
xattr -dr com.apple.quarantine "$dest" 2>/dev/null || true

# Make Finder notice the new service.
/System/Library/CoreServices/pbs -flush >/dev/null 2>&1 || true
/System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true

say "Done."
cat <<'MSG'

Select two Word documents in Finder and right-click: "Compare in Word" is in
the menu (under Services if your menu is long).

The first run asks for two permissions, which macOS prompts for:
  - Automation: Finder controlling Microsoft Word and System Events.
  - Accessibility: System Settings > Privacy & Security > Accessibility,
    with Finder enabled (add /System/Library/CoreServices/Finder.app if it
    is not listed). Without it the dialog opens with empty fields.

Settings: ~/Library/Application Support/CompareInWord/settings.conf
Log of the last run: ~/Library/Application Support/CompareInWord/trace.txt
MSG
