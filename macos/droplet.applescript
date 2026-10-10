-- pptx-slim drag-and-drop droplet
--
-- Drop a .pptx on the built app, type a target size in MB, get a smaller file.
-- The engine (pptx_slim.py) is embedded in Contents/Resources by build.sh.
--
-- Dialogs follow the system language: Chinese on a Chinese Mac, English
-- elsewhere. Same rule the Python engine uses.

-- Shorthand: pick the string matching the system language.
on L(zh, en)
	if my isChinese() then return zh
	return en
end L

on isChinese()
	try
		set loc to user locale of (get system info)
		return (loc starts with "zh")
	end try
	return false
end isChinese

-- The app's own name, so a localised build shows the right title.
on appName()
	try
		return name of me
	end try
	return "PPT Slimmer"
end appName

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
	set ttl to my appName()

	set n to count of theFiles
	if n > 1 then
		set promptText to my L("这 " & n & " 个文件，每个压到多少 MB？", ¬
			"Target size in MB for each of the " & n & " files?")
	else
		set promptText to my L("压到多少 MB？", "Target size in MB?")
	end if

	try
		set answer to display dialog promptText default answer "5" with title ttl ¬
			buttons {my L("取消", "Cancel"), my L("开始", "Shrink")} ¬
			default button (my L("开始", "Shrink")) with icon note
	on error
		return
	end try
	set targetMB to text returned of answer

	-- Accept "5", "5.5", "3 MB" alike; refuse anything else before we start work.
	try
		set targetNum to (targetMB as real)
	on error
		display dialog my L("这不是个数字：" & targetMB, "That is not a number: " & targetMB) ¬
			with title ttl buttons {"OK"} default button "OK" with icon stop
		return
	end try
	if targetNum ≤ 0 then
		display dialog my L("目标大小必须大于 0。", "The target must be greater than zero.") ¬
			with title ttl buttons {"OK"} default button "OK" with icon stop
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
			set report to report & my L("失败：", "Failed: ") & my lastLines(errMsg, 3) & linefeed & linefeed
		end try
	end repeat

	display dialog report with title ttl buttons {"OK"} default button "OK"
end open

on run
	set msg to my L("把一个或多个 .pptx 文件拖到这个 App 上。" & return & return & ¬
		"它会问你目标大小（MB），然后在每个原文件旁边生成" & return & ¬
		"「原名 (compressed).pptx」。动图、透明素材和内嵌视频" & return & ¬
		"一律不动。", ¬
		"Drag one or more .pptx files onto this app." & return & return & ¬
		"It asks for a target size in MB, then writes \"<name> (compressed).pptx\"" & return & ¬
		"next to each original. Animated GIFs, transparent art and embedded" & return & ¬
		"video are never touched.")
	display dialog msg with title my appName() buttons {"OK"} default button "OK" with icon note
end run
