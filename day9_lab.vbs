Option Explicit

Dim num, i, facto

facto = 1

num = InputBox("Enter a number to find the factorial:")

For i=1 to num
	facto = facto * i	
Next

MsgBox(facto)
