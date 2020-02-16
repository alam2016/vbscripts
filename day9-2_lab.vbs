Option Explicit

'Dim Arr, i
'Arr = Array(5, 2, 7, 9, 1, 56, 89, 85)
'MsgBox UBound(Arr)

Dim Arr(), i

ReDim Arr(3)

Arr(0) = 9
Arr(1) = 3
Arr(2) = 2
Arr(3) = 5

' Before ReDim
For i=0 to UBound(Arr)
	MsgBox(Arr(i))
Next

' After ReDim

ReDim Preserve Arr(5)

Arr(2) = 7777
Arr(4) = 9999
Arr(5) = 8888


For i=0 to UBound(Arr)
	MsgBox(Arr(i))
Next
