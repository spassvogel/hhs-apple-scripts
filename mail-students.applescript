

set _mailSuffix to "@student.hhs.nl"
set _mailSubjectResponse to display dialog "Mail subject" default answer "Resultaten voor <VAK>" with icon note buttons {"Cancel", "Continue"} default button "Continue"
set _mailBodyResponse to display dialog "Mail body" default answer "Beste student" & linefeed & linefeed & linefeed with icon note buttons {"Cancel", "Continue"} default button "Continue"
set _mailbody to replace_chars((text returned of _mailBodyResponse), linefeed, "<br/>")

set myFolder to (choose folder with prompt "Where are the assessments?")
log myFolder
tell application "Finder"
	set _pdfs to (every file in entire contents of myFolder whose name ends with ".pdf") as alias list
end tell
log _pdfs
repeat with _file in _pdfs
	tell application "Finder"
		set _fileName to (name of (_file))
	end tell
	
	set _fileNameNoExtension to remove_extension(_fileName)
	set _recipient to _fileNameNoExtension & _mailSuffix
	
	tell application "Microsoft Outlook"
		set _message to make new outgoing message with properties {subject:(text returned of _mailSubjectResponse), content:_mailbody}
		tell _message
			make new attachment with properties {file:_file}
			make new recipient at _message with properties {email address:{address:_recipient}}
		end tell
		open _message
	end tell
end repeat

on replace_chars(_text, _search_string, _replacement_string)
	set AppleScript's text item delimiters to the _search_string
	set the item_list to every text item of _text
	set AppleScript's text item delimiters to the _replacement_string
	set _this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return _this_text
end replace_chars

on remove_extension(this_name)
	if this_name contains "." then
		set this_name to Â
			(the reverse of every character of this_name) as string
		set x to the offset of "." in this_name
		set this_name to (text (x + 1) thru -1 of this_name)
		set this_name to (the reverse of every character of this_name) as string
	end if
	return this_name
end remove_extension
