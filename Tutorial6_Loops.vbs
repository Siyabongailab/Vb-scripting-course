
dim intNUmber, intTotal 
intTotal = 0

do 
intNUmber = CInt(InputBox("please enter an integer: "))

if intNUmber = 0 then
msgbox("goodbye!")
exit do 
end if
intTotal = intTotal + intNUmber
msgbox(intTotal)
Loop