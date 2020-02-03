option explicit

dim strCreditCard

strCreditCard = InputBox("Credit Card Type:")

Select Case strCreditCard
	Case "Master" 
		MsgBox "MC" 
		call add(5, 2)
	Case "Visa" 
		MsgBox "Visa"
		call subtraction(5, 2)
	Case Else 	
		MsgBox "Unknown Credit Card "&date
End Select 


function add(num1, num2)
	MsgBox num1 + num2
end function

function subtraction(num1, num2)
	MsgBox num1 - num2
end function

function multi(num1, num2)
	MsgBox num1 * num2
end function

function division(num1, num2)
	MsgBox num1 / num2
end function
