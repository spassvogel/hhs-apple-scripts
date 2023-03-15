(*
====================================================
   [CSV] How to Read CSV File into AppleScript List
====================================================

PURPOSE:
   ‰Û¢ Provide a simple process to read a CSV file, and parse each line/row into a AS list
       for further use as needed by the programmer.
   ‰Û¢ This is a tool/example for the AppleScript programmer, not the end-user of a script
   ‰Û¢ See limitations below
   
DATE:    Fri, Dec 11, 2015            VER: 1.0
AUTHOR: JMichaelTX (on MacScripter.net and StackOverflow.com forums)

LIMITATIONS:
   ‰Û¢ I don't believe the parsing of the CSV data here conforms to the RFC 4180 de facto standard
   ‰Û¢ It doesn't allow for clear separation of strings from numbers
   ‰Û¢ If you need a more full-featured, compliant CSV parser, then see the below ref.

REF:    For a more robust, extensive process, see:
           CSV-to-list converter by Nigel Garvey, 2010-03-13
           [url]http://macscripter.net/viewtopic.php?pid=125444#p125444[/url]
           
SAMPLE INPUT DATA (from CSV file):
   Parent,Tag,Num Notes            <=== first line/row is column titles
   !SYMBOLS,SYM.ES,TBD
   ZIP_List,ZIP.77077,TBD
   FINANCE,FIN.Call,TBD
   Evernote,EN.UI,TBD
   HISTORY,HiS.NatlSec,TBD
   . . .

SAMPLE OUTPUT DATA (log)
   (*Number of Rows: 102*)
   (*!SYMBOLS, SYM.ES, TBD*)
   (*ZIP_List, ZIP.77077, TBD*)
   (*FINANCE, FIN.Call, TBD*)
   (*Evernote, EN.UI, TBD*)
   (*HISTORY, HiS.NatlSec, TBD*)
   . . .
====================================================
*)
--- GET THE CSV FILE YOU WANT TO READ ---
set pathInputFile to (choose file with prompt "Select the CSV file" of type "csv")
-- set _mailSubjectResponse to display dialog "Mail subject" default answer "Resultaten voor <VAK>" with icon note buttons {"Cancel", "Continue"} default button "Continue"

-- set _mailBodyResponse to display dialog "Mail body" default answer "Beste {1}, " & linefeed & "bij dezen jouw resultaat voor proeftoets 2: {3}. " & linefeed & "Voor beide proeftoetsen heb je gemiddeld een: {4}. " & linefeed & linefeed & "met vriendelijke groet," & linefeed & linefeed & "Wouter van den Heuvel" with icon note buttons {"Cancel", "Continue"} default button "Continue"
-- set _mailbodyTemplate to replace_chars((text returned of _mailBodyResponse), linefeed, "<br/>")

--- READ THE FILE CONTENTS ---
set strFileContents to read pathInputFile

--- BREAK THE FILE INTO PARAGRAPHS (i.e., ROWS or LINES) ---
--        (AS Paragraphs are separated by LF or CR)

set parFileContents to (paragraphs of strFileContents)
set numRows to count of parFileContents
log "Number of Rows: " & numRows
set lstFieldsinHeader to parseCSV(item 1 of parFileContents as text)
log lstFieldsinHeader

--- PROCESS EACH PARAGRAPH (AKA LINE or ROW) OF INPUT FILE ---
choose from list lstFieldsinHeader with title "Email column" with prompt "Please choose email column" default items "E-mailadres" with multiple selections allowed
(*
-- repeat with iPar from 2 to 3
repeat with iPar from 2 to number of items in parFileContents
	--        Skip first row since it has column titles, data starts in 2nd row
	
	set lstRow to item iPar of parFileContents
	if lstRow = "" then exit repeat -- EXIT Loop if Row is empty, like the last line
	
	set lstFieldsinRow to parseCSV(lstRow as text)
	set _mailbody to _mailbodyTemplate
	set _recipient to item 2 of lstFieldsinRow
	
	(*
	todo: allow index to be input by user, 
	make suggestions based on column names
	*)
	set strName to item 5 of lstFieldsinRow -- COL 1 of CSV file
	set strGrade to item 3 of lstFieldsinRow -- COL 2 of CSV file
	set strGrade2 to item 4 of lstFieldsinRow -- COL 2 of CSV file
	set numNotes to item 3 of lstFieldsinRow -- COL 3 of CSV file
	
	repeat with i from 1 to count of lstFieldsinRow
		set _mailbody to replace_chars(_mailbody, "{" & i & "}", item i of lstFieldsinRow)
	end repeat
	log _recipient
	
	tell application "Microsoft Outlook"
		set _message to make new outgoing message with properties {subject:(text returned of _mailSubjectResponse), content:_mailbody}
		tell _message
			make new recipient at _message with properties {email address:{address:_recipient}}
		end tell
		open _message
	end tell
	
end repeat -- with iPar
*)
--=============== END OF MAIN SCRIPT ==============

on parseCSV(pstrRowText)
	set {od, my text item delimiters} to {my text item delimiters, ";"}
	set parsedText to text items of pstrRowText
	set my text item delimiters to od
	return parsedText
end parseCSV

--  set TestString to "1-2-3-5-8-13-21"
-- set myArray to my theSplit(TestString, "-")
on theSplit(theString, theDelimiter)
	-- save delimiters to restore old settings
	set oldDelimiters to AppleScript's text item delimiters
	-- set delimiters to delimiter to be used
	set AppleScript's text item delimiters to theDelimiter
	-- create the array
	set theArray to every text item of theString
	-- restore the old setting
	set AppleScript's text item delimiters to oldDelimiters
	-- return the result
	return theArray
end theSplit

on replace_chars(_text, _search_string, _replacement_string)
	set AppleScript's text item delimiters to the _search_string
	set the item_list to every text item of _text
	set AppleScript's text item delimiters to the _replacement_string
	set _this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return _this_text
end replace_chars
