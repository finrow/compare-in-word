#!/bin/sh
#
# compare-in-word.sh — entry point for the "Compare in Word" Finder Quick Action.
#
# Finder hands us the two selected Word documents as "$@". Finder's selection
# order follows the window's current sort, not the order you clicked, so we
# work out which document is the baseline ourselves and let the user correct
# it from the confirmation dialog.
#
#   Rank order: filename version token  >  filename date stamp  >  mtime
#
# Debug:  compare-in-word.sh --print-order fileA fileB

SELF_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
APPLESCRIPT="$SELF_DIR/compare-in-word.applescript"
TRACE="$SELF_DIR/trace.txt"
LOCKDIR="$SELF_DIR/.running.lock"

PRINT_ONLY=0
if [ "$1" = "--print-order" ]; then
    PRINT_ONLY=1
    shift
fi

# --- single-instance lock ----------------------------------------------------
# Two runs at once fight over the same Compare Documents dialog and BOTH wedge.
# mkdir is atomic, so it works as a lock without flock (which macOS sh lacks).
#
# The lock records its owner's PID. A run that is killed cannot clean up after
# itself, and a fixed staleness timeout either disables the tool for minutes or
# lets a live run be trampled — checking whether the owner is still alive is
# exact and immediate.
if [ "$PRINT_ONLY" -eq 0 ]; then
    if ! mkdir "$LOCKDIR" 2>/dev/null; then
        _owner=$(cat "$LOCKDIR/pid" 2>/dev/null)
        if [ -n "$_owner" ] && kill -0 "$_owner" 2>/dev/null; then
            # Genuinely running. Do NOT touch $TRACE — clobbering it here
            # destroys the evidence of what the stuck run was doing, which is
            # exactly when that evidence matters most.
            /usr/bin/osascript \
                -e 'on run {t, m}' \
                -e 'display alert t message m as warning' \
                -e 'end run' \
                "Compare in Word" \
                "A comparison is already running.

Finish or cancel the open Compare Documents dialog, then try again." >/dev/null 2>&1
            exit 0
        fi
        # Owner is gone: reclaim.
        rm -rf "$LOCKDIR" 2>/dev/null
        mkdir "$LOCKDIR" 2>/dev/null || exit 0
    fi
    echo "$$" > "$LOCKDIR/pid" 2>/dev/null
    # Not `exec`ed below, so this trap does get a chance to fire.
    trap 'rm -rf "$LOCKDIR" 2>/dev/null' EXIT INT TERM

    # Clear last run's flattened copies; they are disposable and would
    # otherwise pile up in the support folder.
    rm -rf "$SELF_DIR/tmp" 2>/dev/null

    # Only now, holding the lock, is it safe to reset the trace.
    : > "$TRACE"
    printf 'shell: start %s (pid %s)\n' "$(date '+%H:%M:%S')" "$$" >> "$TRACE"
    printf 'shell: argc=%s\n' "$#" >> "$TRACE"
    for _a in "$@"; do printf 'shell: arg=[%s]\n' "$_a" >> "$TRACE"; done
fi

alert() {
    [ "$PRINT_ONLY" -eq 1 ] && { printf 'ALERT: %s\n' "$1" >&2; return; }
    /usr/bin/osascript \
        -e 'on run {t, m}' \
        -e 'display alert t message m as warning' \
        -e 'end run' \
        "Compare in Word" "$1" >/dev/null 2>&1
}

# Strip the extension, lowercase, and pull out the last "v4" / "rev 3" /
# "draft2" style token. Prints a bare number, or nothing if there is no
# version token. A trailing copy suffix ("Report 2.docx") has no v/rev/draft
# prefix, so it is correctly ignored.
version_of() {
    printf '%s' "$1" \
        | /usr/bin/sed -E 's/\.[A-Za-z0-9]{1,5}$//' \
        | /usr/bin/tr '[:upper:]' '[:lower:]' \
        | /usr/bin/grep -oE '(^|[^a-z0-9])(v|ver|version|rev|revision|draft)[ _.-]?[0-9]+(\.[0-9]+)?' \
        | /usr/bin/tail -1 \
        | /usr/bin/grep -oE '[0-9]+(\.[0-9]+)?' \
        | /usr/bin/tail -1
}

# Pull out the last ISO-ish date stamp (2026-07-16, 2026_07_16, 20260716) and
# normalise it to YYYYMMDD for numeric comparison. Ambiguous US-style stamps
# (7.16.26) are deliberately not parsed — they collide with version numbers
# like "v1.2", and the mtime fallback plus the confirmation dialog cover them.
date_of() {
    printf '%s' "$1" \
        | /usr/bin/sed -E 's/\.[A-Za-z0-9]{1,5}$//' \
        | /usr/bin/grep -oE '(19|20)[0-9]{2}[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12][0-9]|3[01])' \
        | /usr/bin/tail -1 \
        | /usr/bin/tr -d '\-_.'
}

# True when $1 sorts before $2 numerically.
lt() {
    /usr/bin/awk -v a="$1" -v b="$2" 'BEGIN { exit !((a + 0) < (b + 0)) }'
}

if [ "$#" -ne 2 ]; then
    alert "Select exactly two Word documents, then choose Compare in Word.

You selected $# item(s)."
    exit 0
fi

A="$1"
B="$2"

for f in "$A" "$B"; do
    if [ ! -f "$f" ]; then
        alert "This file could not be found:

$f"
        exit 0
    fi
    if [ ! -r "$f" ]; then
        alert "This file is not readable:

$f"
        exit 0
    fi
done

# --- OneDrive Files-On-Demand pre-flight ------------------------------------
# An online-only ("dataless") file opens fine and only stalls inside read(),
# so Word would beachball for the whole download with no cancel affordance and
# then report an opaque error. Fault the file in here instead, where we own the
# timeout and the message. Reading one byte pulls the entire file; there is no
# byte-level progress to report, so we show an indeterminate notice after 5s.
MAX_WAIT=90

is_dataless() {
    case "$(/usr/bin/stat -f '%Sf' -- "$1" 2>/dev/null)" in
        *dataless*) return 0 ;;
        *) return 1 ;;
    esac
}

notify() {
    [ "$PRINT_ONLY" -eq 1 ] && { printf 'NOTIFY: %s\n' "$1" >&2; return; }
    /usr/bin/osascript \
        -e 'on run {m}' \
        -e 'display notification m with title "Compare in Word"' \
        -e 'end run' \
        "$1" >/dev/null 2>&1
}

# 0 = local, 3 = timed out, 4 = provider refused
materialize() {
    mf="$1"
    is_dataless "$mf" || return 0
    /bin/dd if="$mf" bs=1 count=1 of=/dev/null 2>/dev/null &
    mpid=$!
    mi=0
    while kill -0 "$mpid" 2>/dev/null; do
        sleep 1
        mi=$((mi + 1))
        [ "$mi" -eq 5 ] && notify "Downloading $(/usr/bin/basename -- "$mf") from OneDrive…"
        if [ "$mi" -ge "$MAX_WAIT" ]; then
            # SIGTERM aborts the fetch cleanly and leaves no partial state.
            kill -TERM "$mpid" 2>/dev/null
            wait "$mpid" 2>/dev/null
            return 3
        fi
    done
    wait "$mpid" 2>/dev/null || return 4
    return 0
}

for f in "$A" "$B"; do
    materialize "$f"
    case "$?" in
        3)
            alert "Timed out downloading this file from OneDrive:

$(/usr/bin/basename -- "$f")

Check your network, and that OneDrive is running and signed in."
            exit 0
            ;;
        4)
            alert "OneDrive would not provide this file:

$(/usr/bin/basename -- "$f")

You may be offline or signed out of OneDrive."
            exit 0
            ;;
    esac
done

NA=$(/usr/bin/basename -- "$A")
NB=$(/usr/bin/basename -- "$B")

VA=$(version_of "$NA")
VB=$(version_of "$NB")
DA=$(date_of "$NA")
DB=$(date_of "$NB")

if [ -n "$VA" ] && [ -n "$VB" ] && [ "$VA" != "$VB" ]; then
    REASON="Ordered by version number in the filename (v$VA before v$VB)."
    if lt "$VA" "$VB"; then
        ORIG="$A"; REV="$B"
    else
        ORIG="$B"; REV="$A"
        REASON="Ordered by version number in the filename (v$VB before v$VA)."
    fi
elif [ -n "$DA" ] && [ -n "$DB" ] && [ "$DA" != "$DB" ]; then
    REASON="Ordered by the date in the filename."
    if lt "$DA" "$DB"; then
        ORIG="$A"; REV="$B"
    else
        ORIG="$B"; REV="$A"
    fi
else
    MA=$(/usr/bin/stat -f %m -- "$A" 2>/dev/null || echo 0)
    MB=$(/usr/bin/stat -f %m -- "$B" 2>/dev/null || echo 0)
    REASON="No version number found, so ordered by last-modified date."
    if lt "$MB" "$MA"; then
        ORIG="$B"; REV="$A"
    else
        ORIG="$A"; REV="$B"
    fi
fi

if [ "$PRINT_ONLY" -eq 1 ]; then
    printf 'ORIGINAL: %s\nREVISED : %s\nREASON  : %s\n' \
        "$(/usr/bin/basename -- "$ORIG")" "$(/usr/bin/basename -- "$REV")" "$REASON"
    exit 0
fi


printf 'shell: handing off to applescript\n' >> "$TRACE"
# Deliberately NOT `exec`: the EXIT trap must still run so the lock is released.
/usr/bin/osascript "$APPLESCRIPT" "$ORIG" "$REV" "$REASON"
_rc=$?
printf 'shell: applescript exited %s\n' "$_rc" >> "$TRACE"
# Always exit 0: the AppleScript reports its own failures in a readable dialog,
# and a non-zero status here just makes Automator add a second, useless
# "Run Shell Script encountered an error" alert on top of it.
exit 0
