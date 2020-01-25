Option Explicit

Dim firstname, lastname

firstname = InputBox("What is your first name")
lastname  = InputBox("What is your last name")

' both value must be false to execute inside if block
If NOT(firstname = "Kaiser" OR lastname = "Alam") Then
	MsgBox "Both value must be false [NOT True]: "& firstname &" "&lastname
End If

' one value must be true to execute inside if block
If firstname = "Kaiser" OR lastname = "Alam" Then
	MsgBox "One value must be true [OR]: "& firstname &" "&lastname
End If

' both value must be true to execute inside if block
If firstname = "Kaiser" AND lastname = "Alam" Then
	MsgBox "Both value must be true [AND]: "& firstname &" "&lastname
End If
