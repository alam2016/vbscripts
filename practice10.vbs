Option Explicit


Dim age, gender, maritalStatus

age = InputBox("Enter your age:")
gender = InputBox("Enter your gender (M/F):")
maritalStatus = InputBox("Are you Married (Y/N):")

If gender = "F" Then
	MsgBox "You are Female. Married: "&maritalStatus&", You should work in Urban area"
Else
	If gender = "M" And CInt(age) => 20 And CInt(age) <= 40 Then
		MsgBox "You are Male. Married: "&maritalStatus&", age between 20 and 40. You should work Anywhere"
	ElseIf gender = "M" And CInt(age) => 40 And CInt(age) <= 60 Then
		MsgBox "You are Male. Married: "&maritalStatus&", age between 40 and 60. You should work Urban area."	
	Else
		MsgBox "Error in input data."
	End If	
End If

