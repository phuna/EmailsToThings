--
-- Properties to be adjusted:
--
-- Here we specify if we want tot truncate the message (1) or not (0)
property truncateMsg : 1

-- in case we truncate, specify the length (number of chars)
property truncateLength : 500

-- here set a due date offset in number of days
-- eg. 1 ==> Due date Today + 1, i.e. Tomorrow
-- if < 0 (eg. -1) then no due date
property dueDateOffset : 1

-- Variables
set todoNotes to ""
set todoName to ""
set todoTags to ""


tell application "Microsoft Outlook"
	activate
	
	set selectedMessages to current messages
	if selectedMessages is {} then
		display dialog "Please select one or more messages" with icon 1
		return
	end if
	
	-- repeat with theMessage in selectedMessages
	set theMessage to item 1 of selectedMessages
	set theSubject to subject of theMessage
	set theContent to ((content of theMessage) as text)
	
	-- Check for a blank subject line,
	if theSubject is missing value then
		set theSubject to "No Subject"
	end if
	-- End check
	
	set theID to id of theMessage as string
	
	-- Set sender
	set theSender to sender of theMessage
	try
		set theSenderName to name of theSender
	on error
		set theSenderName to address of theSender
	end try
	
	-- Convert HTML mail to text
	if theContent contains "<body " then
		set TID to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "<body "
		set theContent to (text item 2 of theContent)
		set AppleScript's text item delimiters to "</body>"
		set theContent to " " & (text item 1 of theContent)
		set AppleScript's text item delimiters to TID
		
	else if theContent contains "<body>" then
		set TID to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "<body>"
		set theContent to (text item 2 of theContent)
		set AppleScript's text item delimiters to "</body>"
		set theContent to ">" & (text item 1 of theContent)
		set AppleScript's text item delimiters to TID
	else
		set theContent to ">" & theContent
	end if
	set theContent to "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" /></head> <body" & theContent & "</body></html>"
	
	
	-- here we convert the HTML of the Message Content to plain text to insert into the Note section
	set theTxtContent to (do shell script ("echo " & (quoted form of theContent) & " |textutil -format html -convert txt -stdin -stdout") as «class utf8»)
	
	if (truncateMsg is 1) then
		if length of theTxtContent > truncateLength then
			set theTxtContent to (text 1 through truncateLength of theTxtContent) & "[...]"
		end if
	end if
	
	set todoName to theSubject
	
	-- set the task's note
	set todoNotes to "[url=outlooktothings://message?id=" & theID & "]" & theSenderName & " // " & theSubject & "[/url]" & linefeed & "Sent: " & time sent of theMessage & linefeed & linefeed & theTxtContent
	
	set todoTypes to {"Response", "Action", "Watch", "ToRead"}
	set answer to (choose from list todoTypes with prompt "Choose one." without multiple selections allowed) as text
	if answer is not "false" then
		set todoTags to answer
	end if
	
	if (dueDateOffset < 0) then
		set theDueDate to missing value
	else
		set theDueDate to (current date) + dueDateOffset * days
	end if
	
	if todoTags is not "" then
		tell application "Things"
			set newToDo to make new to do with properties {name:todoName, notes:todoNotes, tag names:todoTags} at beginning of list "Inbox"
			if (theDueDate is not missing value) then
				set due date of newToDo to theDueDate
			end if
		end tell
		
		display notification "Task \"" & todoName & "\" added" with title "Outlook to Things"
	end if
end tell

