-- url format: outlooktothings://message?id=<id>
on open location outlooktothingsURL
	set oldDelims to AppleScript's text item delimiters
	--This saves Applescript's old text item delimiters to the variable oldDelims.
	set newDelims to {"outlooktothings://", "?id="}
	--This sets the variable newDelims to our new custom url handler prefix and the prefix for the page number argument.
	set AppleScript's text item delimiters to newDelims
	--This sets Applescript's text item delimiters to the newDelims.
	set theId to item -1 of the text items of outlooktothingsURL
	
	tell application "Microsoft Outlook"
		activate
		open message id theId
	end tell
	
end open location
