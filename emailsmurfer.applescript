###
# EmailSmurfer
# Extract email addresses from email header on Apple Mail (OSX)
###

# Mark: 1hour for 100k emails with compiled version (on MacBookPro 2,8GHz i7)
# Select a message before launch the script

on lowerCase(this_text)
	set the comparison_string to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	set the source_string to "abcdefghijklmnopqrstuvwxyz"
	set the new_text to ""
	repeat with this_char in this_text
		set x to the offset of this_char in the comparison_string
		if x is not 0 then
			set the new_text to (the new_text & character x of the source_string) as string
		else
			set the new_text to (the new_text & this_char) as string
		end if
	end repeat
	return the new_text
end lowerCase


display dialog "Starting EmailSmurfer

Which input?
    - For \"All mailboxes\": will take time
    - For \"Selected messages\"" buttons �
	{"All mailboxes", "Selected messages"} default button 2

if button returned of result is "All mailboxes" then
	set myInput to "all"
else
	set myInput to "selected"
end if

tell application "Mail"
	if myInput is "all" then
		set allSenders to sender of every message in mailboxes
	else
		set arrayAllSenders to {}
		set selectionMessage to selection
		repeat with i from 1 to (count selectionMessage)
			set end of arrayAllSenders to sender of item i of selectionMessage
		end repeat
		set allSenders to arrayAllSenders
	end if
end tell

set listOfEmails to {}
set arrayListOfEmails to {}

repeat with i from 1 to (count allSenders)
	tell application "Mail" to set theFromEmail to (extract address from item i of allSenders)
	if listOfEmails does not contain lowerCase(theFromEmail) then
		set end of listOfEmails to lowerCase(theFromEmail)
		tell application "Mail" to set theFromName to (extract name from item i of allSenders)
		set end of arrayListOfEmails to {name:theFromName, email:lowerCase(theFromEmail)}
	end if
end repeat

set StoredListOfEmails to arrayListOfEmails
set pwd to path to documents folder as string
set theFile to pwd & "output.csv"
set theFileID to open for access theFile with write permission
try
	repeat with i from 1 to (count StoredListOfEmails)
		write "\"" & name of item i of StoredListOfEmails & "\",\"" & email of item i of StoredListOfEmails & "\"" & return to theFileID as �class utf8�
	end repeat
end try
close access theFileID

display dialog "Number of messages: " & (count allSenders) & "
Number of emails: " & (count arrayListOfEmails)
