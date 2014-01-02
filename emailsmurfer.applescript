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

on simpleSort(my_list)
	set the index_list to {}
	set the sorted_list to {}
	repeat (the number of items in my_list) times
		set the low_item to ""
		repeat with i from 1 to (number of items in my_list)
			if i is not in the index_list then
				set this_item to item i of my_list as text
				if the low_item is "" then
					set the low_item to this_item
					set the low_item_index to i
				else if this_item comes before the low_item then
					set the low_item to this_item
					set the low_item_index to i
				end if
			end if
		end repeat
		set the end of sorted_list to the low_item
		set the end of the index_list to the low_item_index
	end repeat
	return the sorted_list
end simpleSort


display dialog "Starting EmailSmurfer

What mailbox use?
    - For \"All mailboxes\": will take time
    - For \"Current mailbox\": select a message before" buttons Â
	{"All mailboxes", "Current mailbox"} default button 2

if button returned of result is "All mailboxes" then
	set myInput to "all"
else
	set myInput to "current"
end if

tell application "Mail"
	set selectionMessage to selection
	set thisMessage to item 1 of selectionMessage
	if myInput is "all" then
		set allSenders to sender of every message in mailboxes
	else
		set allSenders to sender of every message in mailbox of thisMessage
	end if
end tell

set listOfEmails to {}
set arrayListOfEmails to {}

repeat with i from 1 to (count allSenders)
	-- tell application "Mail" to set theFromName to (extract name from item i of allSenders)
	tell application "Mail" to set theFromEmail to (extract address from item i of allSenders)
	if listOfEmails does not contain lowerCase(theFromEmail) then
		-- set end of listOfEmails to (theFromName & theFromEmail)
		set end of listOfEmails to lowerCase(theFromEmail)
		tell application "Mail" to set theFromName to (extract name from item i of allSenders)
		set end of arrayListOfEmails to {name:theFromName, email:theFromEmail}
	end if
end repeat

-- set SortedListOfEmails to simpleSort(listOfEmails)
set StoredListOfEmails to arrayListOfEmails
set pwd to path to documents folder as string
set theFile to pwd & "output.csv"
set theFileID to open for access theFile with write permission
try
	repeat with i from 1 to (count StoredListOfEmails)
		write "\"" & name of item i of StoredListOfEmails & "\",\"" & email of item i of StoredListOfEmails & "\"" & return to theFileID as Çclass utf8È
	end repeat
end try
close access theFileID

display dialog "Number of messages: " & (count allSenders) & "
Number of emails: " & (count arrayListOfEmails)
