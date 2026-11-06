-- show-dialog.applescript
--
-- Spawned detached by compare-in-word.applescript, because displaying a Word
-- dialog BLOCKS for as long as it is on screen. Running it in its own process
-- lets the main script fill the fields and exit while this one keeps the dialog
-- alive and then applies the result.
--
-- Why two steps rather than `show`:
--
--   `show` is documented as "Displays and carries out actions initiated in the
--   specified built-in Word dialog box". The second half is not true here.
--   Proven on 2026-07-30: the dialog appeared, both documents were selected, OK
--   was pressed and the dialog closed — and the revised .docx still contained
--   ZERO tracked changes (`<w:ins>`/`<w:del>` counts both nil). The comparison
--   never ran.
--
--   So use the explicit pair instead: `display Word dialog` shows it WITHOUT
--   carrying anything out, and `execute dialog` then applies the settings the
--   dialog was left holding. -1 means OK was pressed; 0 means Cancel, in which
--   case nothing must be applied.
--
-- The dialog is opened this way rather than from a menu because modern Word for
-- Mac has no Compare Documents menu item (it is a Review-ribbon control), and
-- crawling the menu bar over the accessibility API is slow enough to hang.

on trace(t)
	set p to (POSIX path of (path to application support from user domain)) & ¬
		"CompareInWord/trace.txt"
	try
		set f to open for access (POSIX file p) with write permission
		write ("dialog: " & t & linefeed) to f starting at eof as «class utf8»
		close access f
	on error
		try
			close access (POSIX file p)
		end try
	end try
end trace

on run
	tell application "Microsoft Word"
		set dlg to get dialog dialog tools compare documents
		set r to 0
		with timeout of 3600 seconds
			-- `show` is the only verb that even claims to carry the action
			-- out ("Displays and carries out actions initiated in the
			-- specified built-in Word dialog box"). `display Word dialog`
			-- explicitly only displays, and with `execute dialog` being a stub
			-- there would then be no way to apply the result.
			try
				set r to (show dlg)
				my trace("show returned " & r)
			on error e
				my trace("show failed: " & e)
				return
			end try
		end timeout

		-- `execute dialog` is ALSO an unimplemented stub on Word 16.110:
		-- "dialog 198 doesn't understand the execute dialog message". So there
		-- is no way to apply a dialog's settings after the fact, and the
		-- display/execute pair is a dead end. Whether the comparison runs is
		-- therefore entirely down to whichever display verb is used above.
		my trace("dismissed with button code " & r)
	end tell
end run
