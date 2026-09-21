-- riabroi-ratchakan-typeset: export .docx to PDF with Microsoft Word for Mac
--
-- Called by inspect_docx.py (never by hand):
--   osascript word_pdf_mac.applescript <docx> <pdf> <unique-stem> <timeout-seconds>
--
-- Word for Mac runs as a single instance (there is no DispatchEx as on Windows),
-- so this script must never disturb documents the user already has open:
--   * it opens only the private copy that Python placed in the work folder under a random name
--   * it finds that document by its unique name, never through "active document"
--   * it closes only documents carrying that name, without saving
--   * it quits Word only if this script launched Word and no other document is open
--
-- Keep this file ASCII only: osascript may not read a text script as UTF-8.

on run argv
	set inPath to item 1 of argv
	set outPath to item 2 of argv
	set docStem to item 3 of argv
	set waitSec to (item 4 of argv) as integer
	set wasRunning to application "Microsoft Word" is running
	set inFile to POSIX file inPath
	set errMsg to ""
	set errNum to 0
	with timeout of waitSec seconds
		try
			tell application "Microsoft Word" to open inFile
			set docName to my findDoc(docStem, 40)
			if docName is "" then error "riabroi: Word did not open the copy" number 9001
			tell application "Microsoft Word" to save as document docName file name outPath file format format PDF
		on error msg number num
			set errMsg to msg
			set errNum to num
		end try
		-- clean up even when a step above failed
		try
			tell application "Microsoft Word"
				repeat with n in (get name of every document)
					if (n as text) contains docStem then close document (n as text) saving no
				end repeat
				if (not wasRunning) and ((count of documents) is 0) then quit saving no
			end tell
		end try
	end timeout
	if errNum is not 0 then error errMsg number errNum
end run

-- Word may list the new document a moment after "open" returns
on findDoc(docStem, tries)
	repeat tries times
		tell application "Microsoft Word" to set docNames to name of every document
		repeat with n in docNames
			if (n as text) contains docStem then return (n as text)
		end repeat
		delay 0.5
	end repeat
	return ""
end findDoc
