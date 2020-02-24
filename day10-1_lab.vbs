Option Explicit

Dim car, elm

Set car = CreateObject("Scripting.Dictionary")

car.Add "sedan", "BMW"
car.Add "sub", "GMC"
car.Add "sports", "Audi"
car.Add "bus", "MTA"

MsgBox car.Count
' MsgBox car.item("sports")

'MsgBox car.item("sedan")
