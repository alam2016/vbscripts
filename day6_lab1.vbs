option explicit

dim num

num = InputBox("Enter a number: ")

if (num mod 5) = 0 or (num mod 3) = 0 then

	if (num mod 3) = 0 and (num mod 5) = 0 then
		MsgBox "fizz buzz: "&num
		
	elseif (num mod 3) = 0 then
		MsgBox "fizz: "&num
		
	elseif (num mod 5) = 0 then
		MsgBox "buzz: "&num
		
	end if
else
	MsgBox num&", unknown number entered !!! "&date
	
end if
