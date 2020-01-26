Option Explicit

Dim name, time_of_day

name = InputBox("What is your name:")
time_of_day = InputBox("What is the time of the day:")

call greet(time_of_day, name)

Function greet(time_of_the_day, your_name)
	MsgBox time_of_the_day&" "&name
End Function

