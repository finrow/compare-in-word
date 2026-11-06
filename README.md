# Compare in Word

A right-click action for the macOS Finder: select two Word documents, choose
**Compare in Word**, and Microsoft Word opens both, fills in its Compare
Documents dialog, presses OK, and leaves you looking at the redline.

It works out which document is the baseline from the filenames (version
tokens, then date stamps, then modification time), downloads online-only
OneDrive files before Word touches them, and handles the case where the
baseline still contains tracked changes, which otherwise makes Word's Compare
silently produce nothing.

Everything runs locally through Word's own Compare feature. Requires
Microsoft Word for Mac.

## Install

```bash
git clone https://github.com/finrow/compare-in-word.git
cd compare-in-word
./install.sh
```

Select two `.docx`, `.docm` or `.doc` files in Finder and right-click. The
action is in the context menu (inside the **Services** submenu if your menu
is long). The first run asks for two permissions: Automation (Finder
controlling Word and System Events) and Accessibility for Finder, under
System Settings > Privacy & Security. See Troubleshooting below.

Set `AutoClickOK=0` in `settings.conf` if you would rather check the dialog
over and press OK yourself.

To remove everything: `./uninstall.sh`.

## Why it drives the real dialog

An earlier version tried to skip the dialog and produce the redline outright. That is not
possible: Word *declares* a `compare` Apple Event but does not implement it, and calling it
fails with -1708 "doesn't understand the compare message". Word's dictionary is generated
from the VBA object model and is full of stubs like this, so a clean `osacompile` proves
nothing about whether a command actually works.

Going through the real dialog turns out to be better anyway. Only three of its fourteen
controls are reachable from AppleScript, so any headless route would have been guessing at
the other eleven; the dialog uses Word's genuine defaults for all of them.

## Files

| Path | What it is |
|---|---|
| `~/Library/Services/Compare in Word.workflow` | The Quick Action. Registers the right-click menu item. |
| `~/Library/Application Support/CompareInWord/compare-in-word.sh` | Entry point. Picks the baseline, materializes OneDrive files. |
| `~/Library/Application Support/CompareInWord/compare-in-word.applescript` | Opens both documents, then fills the dialog's two fields. |
| `~/Library/Application Support/CompareInWord/show-dialog.applescript` | Runs detached; holds the dialog open. See below. |
| `~/Library/Application Support/CompareInWord/settings.conf` | Your comparison settings. Never overwritten by `install.sh`. |

The Quick Action is a two-line stub that calls the shell script, so you can edit the
installed scripts freely without reinstalling. In this repository the same files live in
`support/` and `services/template/`; `install.sh` copies them into place.

## Why there are two AppleScripts

Word's `show <dialog>` command blocks for as long as the dialog is on screen. Left in the
main script it held the Quick Action open — that spinning "0% completed" chip in the menu
bar — and eventually failed with -1712 when it outlasted AppleScript's two-minute default.

So `show` runs in its own detached process. The main script fills the fields and exits
straight away; the helper keeps the dialog alive until you dismiss it.

The dialog is opened via `get dialog`, not by clicking a menu item, because modern Word for
Mac has **no Compare Documents entry in the Tools menu at all** — it is a Review-ribbon
control. An earlier version searched the Tools menu for it and always failed.

## Which document becomes the baseline

Finder does **not** give a Service the files in the order you clicked them — the order is
undefined, and differs between cmd-click and shift-click. So the direction is worked out
from the files themselves, and the result is identical whichever order you select them in.

First rule that discriminates wins:

1. **Version token in the filename** — `v4` before `v5`, `draft3` before `draft11`,
   `rev 2` before `rev 7`. Compared numerically, so `v1` sorts before `v10`.
   Requires an explicit keyword (`v`, `ver`, `version`, `rev`, `revision`, `draft`), which is
   what stops a Finder copy suffix like `Report 2.docx` being read as version 2.
2. **ISO-ish date stamp** — `2026-07-01` before `2026-07-16`.
3. **Last-modified date** — older is the baseline.

Deliberately *not* parsed: US-style `7.16.26`, because it collides with version numbers
like `v1.2`. Those fall through to the modified date.

Both fields are then selected explicitly by name. Matching is strict — exact name first,
then an exact stem, and a substring only when exactly one menu entry matches. A loose
"contains" test once picked the wrong file, because `Report v6` is a substring of
`Annual Report v6`, and the comparison silently ran against a
different document that merely shared a suffix.

## Documents that already contain tracked changes

If the **original** still has tracked changes in it, Word's Compare silently does nothing —
no warning, no error, no result. This was the cause of a long stretch of "it ran but there
is no redline": the baseline carried a colleague's tracked changes, and every comparison quietly
produced zero.

So when the original has any, the script copies it to
`~/Library/Application Support/CompareInWord/tmp/`, accepts the tracked changes **in that
copy**, and compares against it. Your file is never modified. This matches what Word's own
dialog says it does — it "considers tracked changes as accepted" — it just fails to do it
when driven this way.

The copy is named `… (changes accepted).docx` and is deleted at the start of the next run.

Because the check reads the file on disk, the original never has to be opened at all — which
is why you see **two** documents during a run rather than three.

## After the comparison

The baseline document is closed and you are left looking at the redline with the cursor at
the top.

A document you already had open is never closed — only one this run opened. Closing uses
`saving no`, so closing a window you were working in could otherwise discard unsaved edits.

Two Word quirks are load-bearing here: `close <document variable>` fails with "doesn't
understand the close message" while `close document "<name>" saving no` works, and
`activate object window "<name>"` fails because window titles carry suffixes like
`  -  Compatibility Mode` — `window 1 of document "<name>"` is exact.

## Comparison settings

The script fills in the two document fields and applies `settings.conf`, and nothing else. Comparison Settings and the
"Show changes At / In" radios are left exactly as you set them — they are your sticky
preferences, and this tool does not second-guess them. Note that "In: Revised document"
marks up the revised file in place rather than producing a separate document.

Because you get Word's real dialog, every setting is genuinely whatever Word has — including
the eleven Comparison Settings checkboxes and the Character/Word-level granularity that
AppleScript cannot reach at all.

For reference, decoding the dialog's nib
(`Word.app/Contents/Resources/Base.lproj/doccomp.nib`) confirms all eleven checkboxes ship
**checked**. The granularity radio encodes "Character level" in the nib, but that is likely
an Interface Builder first-cell artifact that Word overrides at runtime from the
`CompareChangeLevel` key in
`~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg`.

## Troubleshooting

**The menu item doesn't appear.** Re-register and restart Finder:

```bash
/System/Library/CoreServices/pbs -flush && killall Finder
```

It lives under **Quick Actions** in the right-click menu. It only shows when the selection
is `.docx`, `.docm` or `.doc` — and only in Finder.

**Nothing happens on first use.** The first run asks for permission to control Microsoft
Word. Approve it, then run it again. You can review this under
System Settings → Privacy & Security → Automation.

**The dialog opens but the two fields are empty.** Filling them needs a second permission:
System Settings → Privacy & Security → **Accessibility**, with **Finder** enabled (add it
with `+` → `/System/Library/CoreServices/Finder.app` if it isn't listed). The Quick Action
runs as Finder, so that is the process that needs the grant. Without it the script falls
back to opening the dialog empty, which still works — you just pick both documents yourself.

Whenever the fill fails, it writes what it saw to
`~/Library/Application Support/CompareInWord/ui-dump.txt`. That file is the fastest way to
diagnose a change in Word's dialog layout.

**Wrong number of files.** It requires exactly two; anything else shows an explanation.

## OneDrive

Online-only files are downloaded before Word opens them. A placeholder opens fine and only
stalls inside `read()`, so without this Word would beachball for the length of the download
with no way to cancel. There's a 90-second timeout with a distinct message for "timed out"
versus "OneDrive refused".

## Uninstall

```bash
./uninstall.sh
```

## License

MIT.
