Option Explicit

Dim A(4), i

A(0) = 5
A(1) = 7
A(2) = 8
A(3) = 1
A(4) = 9

For i=0 to 4
	MsgBox "old value: "&A(i)
Next

' insert
For i=0 to 4
	A(i) = A(i) * 2
Next

For i=0 to 4
	MsgBox "new value: "&A(i)
Next