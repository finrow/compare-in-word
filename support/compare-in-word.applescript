-- compare-in-word.applescript
--
-- Driven by compare-in-word.sh, which has already worked out which document is
-- the baseline. argv = { originalPOSIXPath, revisedPOSIXPath, reasonText }
--
-- Opens both documents in Word, brings up Compare Documents, and selects them
-- in the two pickers so the redline is one click away.
--
-- Everything below was learned by running this against Word 16.110 and by
-- decoding Word's own doccomp.nib. None of it is guesswork:
--
--  1. Word DECLARES a `compare` Apple Event but does not implement it (-1708).
--     Its dictionary is generated from the VBA object model and is full of
--     stubs, so a clean compile proves nothing.
--
--  2. Word's `open` returns no reply and AppleScript leaves the variable
--     UNBOUND rather than raising, so `set d to open f` fails later with -2753.
--
--  3. `show <dialog>` BLOCKS until dismissed, so it lives in a detached helper.
--
--  4. Word does NOT seed the dialog's fields. The pickers' menus are empty in
--     the nib; Word fills them at runtime from the OPEN documents, which is why
--     both files must be opened first.
--
--  5. The dialog is a top-level `window "Compare Documents"`, not a sheet, and
--     there is no Compare Documents item in the Tools menu to click.
--
--  6. The pickers are NON-EDITABLE NSPopUpButtons (nib outlets
--     originalDocDropList obj296 / revisedDocDropList obj320, cell flag bit 28
--     clear). They cannot be typed into and ignore `set value` — the only way
--     is to open the menu and click an item.
--
--  7. Identify them by ACCESSIBILITY NAME, never by index or position. In the
--     nib's subview array Revised comes first (idx 6) and Original near last
--     (idx 17), so `pop up button 1` is Revised. The nib's
--     NSNibAXRelationshipConnectors confirm obj296 is titled "Original
--     document:" and obj320 "Revised document:", so name matching is exact.
--
--  8. NEVER press Escape unless a menu is confirmed open. With no menu open it
--     cancels the whole dialog.

property startedAt : missing value
property kDocWait : 30 -- ×0.1s to wait for both documents to open
property kWinWait : 100 -- ×0.1s to wait for the dialog to appear
property kMenuWait : 20 -- ×0.1s to wait for a picker's menu to drop down

on run argv
	if (count of argv) < 2 then
		display alert "Compare in Word" message "Internal error: expected two file paths." as warning
		return
	end if

	try
		compareDocs(item 1 of argv, item 2 of argv)
	on error errMsg number errNum
		if errNum is -128 then return
		display alert "Compare in Word" message ¬
			"Word could not compare these documents." & return & return & ¬
			errMsg & " (" & errNum & ")" & return & return & ¬
			"Both documents should still be open in Word — you can compare them by hand from the Review tab ▸ Compare." as warning
	end try
end run


on compareDocs(originalPath, revisedPath)
	set startedAt to (current date)
	trace("applescript: start")
	set origName to nameOf(originalPath)
	set revName to nameOf(revisedPath)

	-- Decide the baseline BEFORE opening anything.
	--
	-- If the original still contains tracked changes, Word's Compare silently
	-- does nothing — no warning, no error, no result. Detecting that from the
	-- file on disk instead of from an opened document means the original never
	-- has to be opened at all: that is why two documents now appear where three
	-- did before (original + revised + accepted copy).
	set baselinePath to originalPath
	set baselineName to origName
	set needsAccept to false
	if my hasTracked(originalPath) then
		set c to my makeFlatCopy(originalPath, origName)
		if c is not missing value then
			set baselinePath to c
			set baselineName to my nameOf(c)
			set needsAccept to true
			trace("original has tracked changes — using " & baselineName)
		end if
	end if

	-- Was the baseline ALREADY open before we touched anything? If so it is the
	-- user's own window, possibly with unsaved edits, and must never be closed
	-- from under them — `close … saving no` would discard their work. Only a
	-- document this run opened is ours to tidy away.
	set baselineWasOpen to false
	try
		tell application "Microsoft Word"
			if (exists document baselineName) then set baselineWasOpen to true
		end tell
	end try

	set baseAlias to (POSIX file baselinePath) as alias
	set revAlias to (POSIX file revisedPath) as alias

	tell application "Microsoft Word"
		activate
		-- Baseline first and REVISED last, so the revised document is active
		-- when the Compare dialog is built.
		ignoring application responses
			open baseAlias
			open revAlias
		end ignoring
	end tell

	-- They must be open before the dialog is built, because that is where its
	-- menus come from.
	repeat kDocWait times
		try
			tell application "Microsoft Word"
				if (exists document baselineName) and (exists document revName) then exit repeat
			end tell
		end try
		delay 0.1
	end repeat

	if needsAccept then my acceptAndSave(baselineName)
	set origName to baselineName

	-- Put the dialog up from a detached process, because `show` blocks for as
	-- long as the dialog is on screen.
	--
	-- Do NOT try to open it by walking the menu bar instead. Descending into
	-- every submenu over the accessibility API takes so long in Word that the
	-- script hangs and no dialog ever appears — and Word's MainMenu.strings
	-- contains no Compare entry anyway, so there is nothing there to find.
	trace("documents opened")
	-- nohup matters: the Quick Action's process group is torn down when the
	-- workflow finishes, and a plain `&` child goes with it. That would leave
	-- the dialog on screen but orphaned — still showing, yet with nothing
	-- listening, so pressing OK does nothing at all.
	do shell script "/usr/bin/nohup /usr/bin/osascript " & ¬
		quoted form of (POSIX path of (path to application support from user domain) & ¬
			"CompareInWord/show-dialog.applescript") & " >/dev/null 2>&1 &"

	set filled to true
	try
		fillDialog(origName, revName, "")
	on error uiErr number uiNum
		set filled to false
		logLine("Fill failed: " & uiErr & " (" & uiNum & ")" & linefeed & (my dumpUI()))
	end try

	-- Only tidy up when we actually pressed OK. With AutoClickOK=0 the dialog
	-- is still the user's to deal with, so closing documents underneath them
	-- would be wrong.
	if filled and ((my cfgValue("AutoClickOK", "1")) is "1") then
		my finishUp(baselineName, revName, baselineWasOpen)
	end if
end compareDocs








-- After the comparison: drop the baseline document and leave the user looking
-- at the redline, cursor at the top.
--
-- `close` is not in Word's dictionary as a document command in the usual sense —
-- `close <a document variable>` fails with "doesn't understand the close
-- message", but `close document "<name>" saving no` works. Same for focus:
-- `activate object window "<name>"` fails because window titles carry suffixes
-- like "  -  Compatibility Mode", while `window 1 of document "<name>"` is exact.
on finishUp(baselineName, revName, baselineWasOpen)
	set dest to my cfgValue("In", "Revised document")
	set target to revName
	if dest is "Original document" then set target to baselineName

	-- Word runs the comparison after the dialog closes, so wait for the redline
	-- to actually land before disturbing anything.
	if dest is not "New document" then
		repeat 150 times
			try
				tell application "Microsoft Word"
					if (count of revisions of document target) > 0 then exit repeat
				end tell
			end try
			delay 0.1
		end repeat
	end if

	if (baselineName is not target) and (not baselineWasOpen) then
		try
			tell application "Microsoft Word"
				close document baselineName saving no
			end tell
			my trace("closed baseline " & baselineName)
		on error e
			my trace("could not close baseline: " & e)
		end try
	else if baselineWasOpen then
		my trace("left " & baselineName & " open — it was already open before this run")
	end if

	-- With "New document" the redline is an untitled document whose name we
	-- cannot predict, and Word already brings it to the front.
	if dest is "New document" then return

	try
		tell application "Microsoft Word"
			activate object (window 1 of document target)
			home key selection move unit a story
		end tell
		my trace("showing " & target & " with the cursor at the top")
	on error e
		my trace("could not focus the redline: " & e)
	end try
end finishUp


on fillDialog(origName, revName, nativeNote)
	tell application "System Events"
		tell process "Microsoft Word"
			set ok to false
			repeat kWinWait times
				if (exists window "Compare Documents") then
					set ok to true
					exit repeat
				end if
				delay 0.1
			end repeat
			if not ok then error "Compare Documents window never appeared."
			my trace("dialog appeared")
		end tell
	end tell

	set trouble to (my pickWithRetry("Original", origName))
	trace("picked original")
	set trouble to trouble & (my pickWithRetry("Revised", revName))
	trace("picked revised")
	set trouble to trouble & (my applySettings())
	trace("settings applied")

	-- OK stays disabled until Word accepts both documents, so it is the honest
	-- signal for whether the fill worked. It must be POLLED, not sampled:
	-- Word re-enables it a beat after the last settings click, and checking
	-- once 0.3s later reported false while the dialog was in fact ready.
	--
	-- The dialog carries TWO OK and TWO Cancel buttons stacked at identical
	-- frames (collapsed vs expanded settings), so ask whether ANY button named
	-- OK is enabled rather than addressing one by index.
	set okOn to false
	repeat 40 times
		try
			tell application "System Events"
				tell process "Microsoft Word"
					tell window "Compare Documents"
						repeat with bRef in (every button)
							try
								set b to contents of bRef
								if ((name of b) as text) is "OK" and (enabled of b) then
									set okOn to true
									exit repeat
								end if
							end try
						end repeat
					end tell
				end tell
			end tell
		end try
		if okOn then exit repeat
		delay 0.25
	end repeat

	trace("OK enabled: " & okOn)
	if okOn and ((my cfgValue("AutoClickOK", "1")) is "1") then
		set trouble to trouble & (my clickOK())
		trace("clickOK returned")
	end if
	trace("done. trouble=" & trouble)

	if (not okOn) or (trouble is not "") or (nativeNote is not "") then
		logLine("OK enabled: " & okOn & linefeed & ¬
			"Wanted original [" & origName & "], revised [" & revName & "]" & ¬
			linefeed & linefeed & trouble & linefeed & nativeNote & linefeed & ¬
			(my dumpUI()))
	end if
end fillDialog


-- Press OK so the comparison actually runs.
--
-- A real accessibility click is what makes Word perform the comparison; the
-- `show` Apple Event claims to "display and carry out" but observably does not.
--
-- The dialog carries TWO OK buttons stacked at identical frames — one for the
-- collapsed state and one for the expanded Comparison Settings state — and only
-- one of them is live. So try each enabled OK in turn and stop as soon as the
-- dialog goes away, rather than assuming the first is the real one.
on clickOK()
	tell application "System Events"
		tell process "Microsoft Word"
			set theOK to missing value
			try
				tell window "Compare Documents"
					repeat with bRef in (every button)
						set b to contents of bRef
						try
							if ((name of b) as text) is "OK" and (enabled of b) then
								set theOK to b
								exit repeat
							end if
						end try
					end repeat
				end tell
			end try
			if theOK is missing value then return "No enabled OK button to click." & linefeed

			-- Click exactly ONE OK and never a second.
			--
			-- The dialog carries two OK buttons stacked at identical frames
			-- (collapsed vs expanded settings). An earlier version clicked each
			-- enabled one in turn until the window went away — but the first
			-- click already runs the comparison, and Word takes a moment to tear
			-- the dialog down, so the second click landed afterwards and
			-- re-invoked Compare, leaving a fresh empty dialog on screen.
			try
				click theOK
			on error e
				return "Could not click OK: " & e & linefeed
			end try

			-- Word runs the comparison before dismissing, so allow real time.
			--
			-- It may also put up a follow-up warning first. Comparing documents
			-- that already contain tracked changes triggers "Word will consider
			-- the tracked changes as accepted… continue?" — common with real
			-- baselines. Left unanswered, that prompt
			-- silently aborts the comparison, which is exactly the symptom:
			-- dialog closes, script reports success, zero revisions produced.
			set warned to ""
			set lastWl to ""
			repeat with i from 1 to 150
				delay 0.1
				try
					repeat with wRef in (every window)
						set w to contents of wRef
						repeat with shRef in (every sheet of w)
							set sh to contents of shRef
							set warned to warned & "  warning sheet on [" & ¬
								(my safeText(name of w)) & "]" & linefeed
							-- Affirmative in preference order; never "Cancel".
							repeat with lbl in {"Yes", "Continue", "OK"}
								try
									click (first button of sh whose name is (lbl as text))
									my trace("dismissed warning with " & lbl)
									exit repeat
								end try
							end repeat
						end repeat
					end repeat
				end try
				-- Word's post-OK warning is a top-level window with an EMPTY
				-- name, not a sheet — e.g. comparing documents that already
				-- contain tracked changes asks "Word will consider the tracked
				-- changes as accepted… continue?". Any baseline that still
				-- carries tracked changes triggers it. Left
				-- unanswered it aborts the comparison silently: the dialog
				-- closes, the script reports success, and zero revisions appear.
				try
					repeat with wRef in (every window)
						set w to contents of wRef
						if ((my safeText(name of w)) is "") then
							repeat with lbl in {"Yes", "Continue", "OK"}
								try
									click (first button of w whose name is (lbl as text))
									set warned to warned & "  alert dismissed with " & lbl & linefeed
									my trace("dismissed alert window with " & lbl)
									exit repeat
								end try
							end repeat
						end if
					end repeat
				end try

				-- Log every Word window seen in the first moments after OK.
				if (i mod 5) is 0 then
					try
						set wl to ""
						repeat with wRef in (every window)
							set wl to wl & "[" & (my safeText(name of (contents of wRef))) & "]"
						end repeat
						if wl is not lastWl then
							my trace("  windows: " & wl)
							set lastWl to wl
						end if
					end try
				end if
				if not (exists window "Compare Documents") then
					if warned is not "" then my trace("warning was shown:" & linefeed & warned)
					return ""
				end if
			end repeat
		end tell
	end tell
	return "Clicked OK but the dialog was still open after 15s." & linefeed
end clickOK


-- Read one key from settings.conf, falling back to a default.
on cfgValue(k, dflt)
	set cfgPath to (POSIX path of (path to application support from user domain)) & ¬
		"CompareInWord/settings.conf"
	try
		set v to do shell script "/usr/bin/grep -E " & quoted form of ("^" & k & "=") & ¬
			" " & quoted form of cfgPath & " | /usr/bin/tail -1 | /usr/bin/cut -d= -f2-"
		if v is not "" then return v
	end try
	return dflt
end cfgValue


-- Restore the user's Comparison Settings from settings.conf.
--
-- Opening the dialog via `get dialog` makes Word build a fresh one with factory
-- defaults rather than the user's sticky preferences, so they are reapplied
-- here. If the file is missing, nothing is touched — that is the opt-out.
on applySettings()
	set cfgPath to (POSIX path of (path to application support from user domain)) & ¬
		"CompareInWord/settings.conf"
	set raw to ""
	try
		set raw to do shell script "/bin/cat " & quoted form of cfgPath
	on error
		return ""
	end try

	set wantAt to ""
	set wantIn to ""
	set cbNames to {}
	set cbVals to {}
	set savedDelims to AppleScript's text item delimiters
	repeat with lnRef in (paragraphs of raw)
		set ln to (contents of lnRef) as text
		if ln is not "" and (ln does not start with "#") and (ln contains "=") then
			set AppleScript's text item delimiters to "="
			set k to text item 1 of ln
			set v to text item 2 of ln
			set AppleScript's text item delimiters to savedDelims
			if k is "At" then
				set wantAt to v
			else if k is "In" then
				set wantIn to v
			else if k is "AutoClickOK" then
				-- behaviour flag read via cfgValue; not a dialog control
			else
				set end of cbNames to k
				set end of cbVals to v
			end if
		end if
	end repeat
	set AppleScript's text item delimiters to savedDelims

	set cfgNote to ""
	-- Each control is set through its own fresh lookup. Clicking ANY control in
	-- this dialog makes Word rebuild it, which invalidates every cached
	-- reference — so a single cached pass silently stops working partway
	-- through. That is why the left-hand checkboxes were being applied and
	-- "White space" (later in the list) was not.
	-- Read every checkbox ONCE, then only re-resolve the ones actually wrong.
	-- Re-resolving per control is what makes this reliable (Word rebuilds the
	-- dialog after each click), but paying that cost for settings that already
	-- match dominated the run time on real documents: ~70s of a 79s run.
	set curNames to {}
	set curVals to {}
	try
		tell application "System Events"
			tell process "Microsoft Word"
				tell window "Compare Documents"
					repeat with cbRef in (every checkbox)
						set cb to contents of cbRef
						try
							set end of curNames to ((name of cb) as text)
							set end of curVals to ((value of cb) as integer)
						end try
					end repeat
				end tell
			end tell
		end tell
	end try
	my trace("read " & (count of curNames) & " checkboxes in one pass")

	repeat with idx from 1 to (count of cbNames)
		set nm to item idx of cbNames
		set wantOn to ((item idx of cbVals) is "1")
		set needs to true
		repeat with j from 1 to (count of curNames)
			if (item j of curNames) is nm then
				set cur to item j of curVals
				if (wantOn and cur is 1) or ((not wantOn) and cur is 0) then set needs to false
				exit repeat
			end if
		end repeat
		if needs then set cfgNote to cfgNote & (my setCheckbox(nm, wantOn))
	end repeat

	-- Read both radio groups in one pass too, and only pay setRadio's full
	-- re-resolve for a radio that is actually wrong. "Already set" was costing
	-- a whole window re-resolve just to discover there was nothing to do.
	set atOK to false
	set inOK to false
	try
		tell application "System Events"
			tell process "Microsoft Word"
				tell window "Compare Documents"
					repeat with rgRef in (every radio group)
						repeat with rbRef in (every radio button of (contents of rgRef))
							set rb to contents of rbRef
							set nm to ""
							try
								set nm to (name of rb) as text
							end try
							set cur to 0
							try
								set cur to (value of rb) as integer
							end try
							if (nm is wantAt) and (cur is 1) then set atOK to true
							if (nm is wantIn) and (cur is 1) then set inOK to true
						end repeat
					end repeat
				end tell
			end tell
		end tell
	end try
	if not atOK then set cfgNote to cfgNote & (my setRadio(wantAt))
	if not inOK then set cfgNote to cfgNote & (my setRadio(wantIn))
	return cfgNote
end applySettings


-- Wait for the dialog window to be resolvable again. Word tears the window
-- down and rebuilds it after certain clicks, so a reference that worked a
-- moment ago can vanish; every interaction re-acquires it.
on waitWindow(tenths)
	tell application "System Events"
		tell process "Microsoft Word"
			repeat tenths times
				if (exists window "Compare Documents") then return true
				delay 0.1
			end repeat
		end tell
	end tell
	return false
end waitWindow


-- Set one checkbox by name, resolving the window and the control fresh, and
-- clicking only when it is in the wrong state.
on setCheckbox(wantName, wantOn)
	if wantName is "" then return ""
	if not (my waitWindow(30)) then
		my trace("setCheckbox " & wantName & ": window gone")
		return "Dialog closed before setting " & wantName & "." & linefeed
	end if
	tell application "System Events"
		tell process "Microsoft Word"
			try
				tell window "Compare Documents"
					repeat with cbRef in (every checkbox)
						set cb to contents of cbRef
						set nm to ""
						try
							set nm to (name of cb) as text
						end try
						if nm is wantName then
							set cur to 0
							try
								set cur to (value of cb) as integer
							end try
							if (wantOn and cur is 0) or ((not wantOn) and cur is 1) then
								-- "Insertions and deletions" is permanently
								-- disabled; clicking it fails harmlessly.
								try
									click cb
									my trace("setCheckbox clicked " & wantName)
								on error e2
									my trace("setCheckbox " & wantName & " click failed: " & e2)
								end try
								-- Confirm the click actually took before anyone
								-- touches the next control. Every click makes
								-- Word REBUILD the dialog, and a click dispatched
								-- while the previous rebuild is still in flight
								-- can land on a re-bound element — observed
								-- 2026-08-16 hitting OK and running the
								-- comparison with half-applied settings. Reading
								-- the fresh value back is what proves the
								-- rebuild has settled.
								set confirmed to false
								repeat 20 times
									delay 0.15
									try
										repeat with cb2Ref in (every checkbox)
											set cb2 to contents of cb2Ref
											if ((name of cb2) as text) is wantName then
												if ((value of cb2) as integer) is (wantOn as integer) then ¬
													set confirmed to true
												exit repeat
											end if
										end repeat
									end try
									if confirmed then exit repeat
								end repeat
								if not confirmed then ¬
									my trace("setCheckbox " & wantName & " NOT CONFIRMED")
							end if
							return ""
						end if
					end repeat
				end tell
			on error e
				my trace("setCheckbox " & wantName & " failed: " & e)
				return "Could not set " & wantName & ": " & e & linefeed
			end try
		end tell
	end tell
	my trace("setCheckbox " & wantName & ": not found")
	return ""
end setCheckbox


-- Select one radio button by name, resolving the window and the control fresh.
-- Radios live inside an AXRadioGroup (they come from NSMatrix), so they are not
-- direct children of the window — `every radio button of window` finds none.
on setRadio(wantName)
	if wantName is "" then return ""
	if not (my waitWindow(30)) then
		my trace("setRadio " & wantName & ": window gone")
		return "Dialog closed before setting " & wantName & "." & linefeed
	end if
	tell application "System Events"
		tell process "Microsoft Word"
			try
				tell window "Compare Documents"
					repeat with rgRef in (every radio group)
						repeat with rbRef in (every radio button of (contents of rgRef))
							set rb to contents of rbRef
							set nm to ""
							try
								set nm to (name of rb) as text
							end try
							if nm is wantName then
								set cur to 0
								try
									set cur to (value of rb) as integer
								end try
								if cur is 0 then
									click rb
									my trace("setRadio clicked " & wantName)
									-- Same rebuild-settle confirmation as the
									-- checkboxes: do not return until the radio
									-- reads back as selected, so the next
									-- interaction cannot race the rebuild.
									set confirmed to false
									repeat 20 times
										delay 0.15
										try
											repeat with rg2Ref in (every radio group)
												repeat with rb2Ref in (every radio button of (contents of rg2Ref))
													set rb2 to contents of rb2Ref
													if ((name of rb2) as text) is wantName then
														if ((value of rb2) as integer) is 1 then ¬
															set confirmed to true
													end if
												end repeat
											end repeat
										end try
										if confirmed then exit repeat
									end repeat
									if not confirmed then ¬
										my trace("setRadio " & wantName & " NOT CONFIRMED")
								else
									my trace("setRadio " & wantName & " already set")
								end if
								return ""
							end if
						end repeat
					end repeat
				end tell
			on error e
				my trace("setRadio " & wantName & " failed: " & e)
				return "Could not set " & wantName & ": " & e & linefeed
			end try
		end tell
	end tell
	my trace("setRadio " & wantName & ": not found")
	return "Radio " & wantName & " not found." & linefeed
end setRadio


-- Word rebuilds the dialog after every interaction, so the control the next
-- pick needs can be momentarily absent or half-built — the failure looks like
-- "Can't get menu 1 … Invalid index". Retry rather than give up, letting the
-- dialog settle in between, and dismiss any menu left hanging open by the
-- previous attempt (checking first that one IS open — a stray Escape with no
-- menu open cancels the whole dialog).
on pickWithRetry(labelWord, fullName)
	set lastErr to ""
	repeat 4 times
		set lastErr to my pickInto(labelWord, fullName)
		if lastErr is "" then return ""
		my trace("retrying " & labelWord & " after: " & lastErr)
		try
			tell application "System Events"
				tell process "Microsoft Word"
					if (exists menu 1 of window "Compare Documents") then key code 53
				end tell
			end tell
		end try
		delay 0.8
	end repeat
	return lastErr
end pickWithRetry


-- Select `fullName` in the picker whose accessibility name contains
-- `labelWord` ("Original" or "Revised"). Returns "" on success, otherwise a
-- description of what went wrong including every menu item Word offered.
--
-- The control is re-resolved on each call rather than cached: Word rebuilds
-- parts of the dialog after a selection, and a stale reference then fails with
-- "Invalid index".
on pickInto(labelWord, fullName)
	set stem to stemOf(fullName)
	set seen to ""
	tell application "System Events"
		tell process "Microsoft Word"
			tell window "Compare Documents"
				set pb to missing value
				try
					repeat with bRef in (every pop up button)
						set b to contents of bRef
						set bn to ""
						try
							set bn to (name of b) as text
						end try
						if bn contains labelWord then
							set pb to b
							exit repeat
						end if
					end repeat
				end try
				if pb is missing value then
					return "No pop up button named like [" & labelWord & "]." & linefeed
				end if

				-- Drop the menu down. Some pop up buttons expose `menu 1`
				-- straight away; otherwise AXShowMenu is the canonical action,
				-- with a plain click as the fallback.
				set theMenu to missing value
				try
					set theMenu to menu 1 of pb
				end try
				if theMenu is missing value then
					try
						perform action "AXShowMenu" of pb
					on error
						try
							click pb
						end try
					end try
					repeat kMenuWait times
						try
							set theMenu to menu 1 of pb
							exit repeat
						end try
						delay 0.1
					end repeat
				end if
				if theMenu is missing value then
					return "Could not open the " & labelWord & " menu." & linefeed
				end if

				-- Word's menu titles drop the .docx extension, hence the stem.
				--
				-- Match STRICTLY, best-first. A bare "contains" test picked the
				-- wrong document: the stem "Report v6" is a substring of
				-- "Annual Report v6", so a comparison
				-- silently ran against a different file that merely shared a
				-- suffix. Containment is only used as a last resort, and only
				-- when exactly one entry matches.
				set hit to missing value
				set loose to {}
				set looseN to 0
				try
					repeat with miRef in (menu items of theMenu)
						set mi to contents of miRef
						set t to ""
						try
							set t to (name of mi) as text
						end try
						if t is not "" then
							set seen to seen & "    [" & t & "]" & linefeed
							if (t is fullName) or (t is stem) then
								set hit to mi
								exit repeat
							else if (t starts with (stem & " ")) or (t starts with (fullName & " ")) then
								-- e.g. "Name  -  Compatibility Mode"
								if hit is missing value then set hit to mi
							else if (t contains stem) then
								set end of loose to mi
								set looseN to looseN + 1
							end if
						end if
					end repeat
				end try
				if (hit is missing value) and (looseN is 1) then set hit to item 1 of loose
				if (hit is missing value) and (looseN > 1) then
					key code 53
					delay 0.2
					return "[" & fullName & "] is ambiguous — " & looseN & ¬
						" menu entries contain the stem [" & stem & "]:" & linefeed & seen
				end if

				if hit is missing value then
					-- Safe here and ONLY here: the menu is confirmed open, so
					-- Escape closes the menu rather than the dialog.
					key code 53
					delay 0.2
					return "No match for [" & fullName & "] / stem [" & stem & ¬
						"] in the " & labelWord & " menu, which offered:" & linefeed & seen
				end if

				click hit
				-- Settle for the dialog rebuild the click triggers. 0.3s is
				-- load-bearing: 0.15s was tried and made the NEXT pick's menu
				-- fail to open, and the retry that healed it cost ~4s — far
				-- more than the 0.15s saved. Do not trim this again.
				delay 0.3
				return ""
			end tell
		end tell
	end tell
end pickInto


-- Does this .docx carry tracked changes? Read straight from the file, so the
-- document does not have to be opened in Word to find out.
on hasTracked(posixPath)
	try
		set n to (do shell script "/usr/bin/unzip -p " & quoted form of posixPath & ¬
			" word/document.xml 2>/dev/null | /usr/bin/grep -cE '<w:(ins|del)[ />]' || true") as integer
		return (n > 0)
	on error
		return false
	end try
end hasTracked


-- Disposable copy of the original, for accepting tracked changes into. Returns
-- its POSIX path, or missing value.
on makeFlatCopy(originalPath, origName)
	set flatName to (my stemOf(origName)) & " (changes accepted).docx"
	set tmpDir to (POSIX path of (path to application support from user domain)) & ¬
		"CompareInWord/tmp"
	set flatPath to tmpDir & "/" & flatName
	try
		do shell script "/bin/mkdir -p " & quoted form of tmpDir & ¬
			" && /bin/cp " & quoted form of originalPath & " " & quoted form of flatPath & ¬
			" && /bin/chmod u+w " & quoted form of flatPath
		return flatPath
	on error e
		my trace("could not copy original: " & e)
		return missing value
	end try
end makeFlatCopy


on acceptAndSave(docName)
	try
		tell application "Microsoft Word"
			accept all revisions document docName
			save document docName
		end tell
		my trace("accepted tracked changes in the copy")
	on error e
		my trace("could not accept changes: " & e)
	end try
end acceptAndSave


on openAndWait(theAlias, theName)
	tell application "Microsoft Word"
		try
			open theAlias
		end try
		repeat 40 times
			if (exists document theName) then return
			delay 0.1
		end repeat
	end tell
end openAndWait


on dumpUI()
	set out to ""
	try
		tell application "System Events"
			tell process "Microsoft Word"
				repeat with wRef in (every window)
					set w to contents of wRef
					set wn to my safeText(name of w)
					set out to out & "WINDOW: " & wn & linefeed
					if wn contains "Compare" then set out to out & (my describeWin(w))
				end repeat
			end tell
		end tell
	on error e
		set out to out & "dump failed: " & e & linefeed
	end try
	return out
end dumpUI


-- Full subtree of a window. `repeat with el in someList` would bind el to a
-- REFERENCE to the list item and every read below would silently yield
-- nothing, so index into the list instead. `role` is plain text, unlike
-- `class of x`, which throws when coerced. Each read gets its own try: a single
-- throwing attribute would otherwise abort the whole dump.
on describeWin(win)
	set out to ""
	tell application "System Events"
		try
			set allEls to entire contents of win
			repeat with i from 1 to (count of allEls)
				set el to item i of allEls
				set ln to "?"
				try
					set ln to my safeText(role of el)
				end try
				try
					set ln to ln & " title=" & my safeText(title of el)
				end try
				try
					set ln to ln & " name=" & my safeText(name of el)
				end try
				try
					set ln to ln & " value=" & my safeText(value of el)
				end try
				try
					set ln to ln & " enabled=" & my safeText(enabled of el)
				end try
				set out to out & "  " & ln & linefeed
			end repeat
		on error e
			set out to out & "  describe failed: " & e & linefeed
		end try
	end tell
	return out
end describeWin


on safeText(v)
	try
		if v is missing value then return "-"
		return v as text
	on error
		return "?"
	end try
end safeText


-- Append a progress marker. Written incrementally and flushed immediately, so
-- if the script stalls the file still shows the last step that completed —
-- ui-dump.txt only gets written at the very end and tells you nothing about a
-- hang.
on trace(t)
	if startedAt is missing value then set startedAt to (current date)
	set el to ((current date) - startedAt) as text
	set t to "[" & el & "s] " & t
	set p to (POSIX path of (path to application support from user domain)) & ¬
		"CompareInWord/trace.txt"
	try
		set f to open for access (POSIX file p) with write permission
		write (t & linefeed) to f starting at eof as «class utf8»
		close access f
	on error
		try
			close access (POSIX file p)
		end try
	end try
end trace


on resetTrace()
	try
		do shell script "/bin/rm -f " & quoted form of ¬
			((POSIX path of (path to application support from user domain)) & ¬
				"CompareInWord/trace.txt")
	end try
end resetTrace


on logLine(t)
	set p to (POSIX path of (path to application support from user domain)) & "CompareInWord/ui-dump.txt"
	try
		set f to open for access (POSIX file p) with write permission
		set eof f to 0
		write t to f as «class utf8»
		close access f
	on error
		try
			close access (POSIX file p)
		end try
	end try
end logLine


on nameOf(posixPath)
	set savedDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "/"
	set n to last text item of posixPath
	set AppleScript's text item delimiters to savedDelims
	return n
end nameOf


on stemOf(fileName)
	if fileName contains "." then
		set savedDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "."
		set parts to text items of fileName
		set n to ""
		repeat with k from 1 to ((count of parts) - 1)
			if k > 1 then set n to n & "."
			set n to n & (item k of parts)
		end repeat
		set AppleScript's text item delimiters to savedDelims
		if n is not "" then return n
	end if
	return fileName
end stemOf
