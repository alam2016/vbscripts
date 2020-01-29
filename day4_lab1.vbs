' Write a program using function that will calculate addition of two numbers

' First create a function with two arguments
' call that function with two positive numbers
' display the result


Option Explicit

' function construction 
Function add(num1, num2)
	MsgBox(num1 + num2)
End Function

Function foo()
	MsgBox("Foo function does not have arguments")
	call greet("Kaiser")
End Function

call foo()
call add(10, 6)

' example nested function 
Function greet(name)
	MsgBox("Hello "&name)
End Function

