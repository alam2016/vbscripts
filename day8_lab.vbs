Option Explicit

Dim name, num1, num2

name = InputBox("Enter your name:")

num1 = InputBox("Enter number:")
num2 = InputBox("Enter number:")


'While (num1 MOD num2) = 0
'
'	MsgBox "Hello, "&name
'	
'	num1 = InputBox("Enter another number:")
'	num2 = InputBox("Enter another number:")
'
'Wend


do 
	MsgBox "Hello, "&name&" num1:"&num1&" num2:"&num2
	
	num1 = InputBox("Enter another number:")
	num2 = InputBox("Enter another number:")

Loop while (num1 MOD num2) = 0
