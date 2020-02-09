Option Explicit

Dim i
Dim clientName


Function greet()

	For i=1 To 5
		clientName = InputBox("What is your name:")
		call foo(clientName)	
	Next

End Function

Function foo(name)
	MsgBox(name&" How are you?")
	MsgBox(clientName)
End Function


call greet()