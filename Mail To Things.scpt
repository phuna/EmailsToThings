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

tell application "Mail"
	activate
	set selectedMsgs to selection
	if selectedMsgs is {} then
		display dialog "Please select one or more messages" with icon 1
		return
	end if
	set theMessage to item 1 of selectedMsgs
	set theSubject to subject of theMessage
	set theContent to ((content of theMessage) as rich text)
	
	-- Check for a blank subject line,
	if theSubject is missing value then
		set theSubject to "No Subject"
	end if
	-- End check
	
	set theID to message id of theMessage as string
	set theSenderName to sender of theMessage
	set theTxtContent to theContent
	
	if (truncateMsg is 1) then
		if length of theTxtContent > truncateLength then
			set theTxtContent to (rich text 1 through truncateLength of theTxtContent) & "[...]"
		end if
	end if
	
	set todoName to theSubject
	
	set todoNotes to "[url=message:%3C" & theID & "%3E]" & theSenderName & " // " & theSubject & "[/url]" & linefeed & "Sent: " & date sent of theMessage & linefeed & linefeed & theTxtContent
	
	set todoTypes to {"Response", "Action", "Watch", "ToRead"}
	set answer to (choose from list todoTypes with prompt "Choose one." without multiple selections allowed) as rich text
	
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
		
		display notification "Task \"" & todoName & "\" added" with title "Mail to Things"
	end if
	
end tell

