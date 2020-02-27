' This program to find how many duplicate friends 
' are there in friend list database
'
Option Explicit

Dim friend_list(4), friend, count, obj

count = 1

' Below is a small friend list database.

friend_list(0) = "John"
friend_list(1) = "Adam"
friend_list(2) = "John"
friend_list(3) = "Sara"
friend_list(4) = "Sara"

Dim friend_count
Set friend_count = CreateObject("Scripting.Dictionary")


For Each friend in friend_list
	if friend_count.exists(friend) Then
		count = count + 1
		friend_count(friend) = count
	Else
		friend_count.add friend, count
		count = 1
	End if	
Next	

For Each obj in friend_count.keys
    Msgbox "Friend: "& obj & vbcrlf &"How many: "& friend_count(obj)
Next