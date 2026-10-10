-- pptx-slim drag-and-drop droplet
--
-- Drop a .pptx on the built app, type a target size in MB, get a smaller file.
-- The engine (pptx_slim.py) is embedded in Contents/Resources by build.sh.

property appTitle : "PPT Slimmer"

-- Find an interpreter that actually has Pillow, rather than assuming one.
on findPython()
	set candidates to {"/opt/homebrew/bin/python3", "/usr/local/bin/python3", "/usr/bin/python3"}
	repeat with p in candidates
		try
			set probe to do shell script quoted form of (p as text) & " -c 'import PIL' 2>/dev/null && echo ok || echo no"
			if probe is "ok" then return p as text
		end try
	end repeat
	return "/usr/bin/python3"
end findPython

on lastLines(txt, n)
	set AppleScript's text item delimiters to linefeed
	set theLines to text items of txt
	set AppleScript's text item delimiters to ""
	if (count of theLines) > n then
		set theLines to items ((count of theLines) - n + 1) thru -1 of theLines
	end if
	set AppleScript's text item delimiters to linefeed
	set out to theLines as text
	set AppleScript's text item delimiters to ""
	return out
end lastLines

on open theFiles
	set resPath to (path to me as text) & "Contents:Resources:pptx_slim.py"
	set engine to POSIX path of resPath
	set py to my findPython()

	set n to count of theFiles
	if n > 1 then
		set promptText to "Target size in MB for each of the " & n & " files?"
	else
		set promptText to "Target size in MB?"
	end if

	try
		set answer to display dialog promptText default answer "5" with title appTitle ¬
			buttons {"Cancel", "Shrink"} default button "Shrink" with icon note
	on error
		return
	end try
	set targetMB to text returned of answer

	-- Accept "5", "5.5", "3 MB" alike; refuse anything else before we start work.
	try
		set targetNum to (targetMB as real)
	on error
		display dialog "That is not a number: " & targetMB with title appTitle buttons {"OK"} default button "OK" with icon stop
		return
	end try
	if targetNum ≤ 0 then
		display dialog "The target must be greater than zero." with title appTitle buttons {"OK"} default button "OK" with icon stop
		return
	end if

	set report to ""
	repeat with f in theFiles
		set posixFile to POSIX path of f
		set report to report & "── " & (do shell script "basename " & quoted form of posixFile) & " ──" & linefeed
		try
			set cmd to quoted form of py & " " & quoted form of engine & " " & ¬
				quoted form of posixFile & " " & (targetNum as text) & " 2>&1"
			set output to do shell script cmd
			set report to report & my lastLines(output, 3) & linefeed & linefeed
		on error errMsg
			set report to report & "Failed: " & my lastLines(errMsg, 3) & linefeed & linefeed
		end try
	end repeat

	display dialog report with title appTitle buttons {"OK"} default button "OK"
end open

on run
	display dialog "Drag one or more .pptx files onto this app." & return & return & ¬
		"It asks for a target size in MB, then writes \"<name> (compressed).pptx\"" & return & ¬
		"next to each original. Animated GIFs, transparent art and embedded" & return & ¬
		"video are never touched." with title appTitle buttons {"OK"} default button "OK" with icon note
end run
