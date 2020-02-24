Option Explicit

Dim arrEmployeeName(5)
Dim employeeName

arrEmployeeName(0) = "John"
arrEmployeeName(1) = "Adam"
arrEmployeeName(2) = "William"
arrEmployeeName(3) = "Smith"
arrEmployeeName(4) = "Alam"
arrEmployeeName(5) = "Kaiser"


For Each employeeName In arrEmployeeName
	MsgBox employeeName
Next



