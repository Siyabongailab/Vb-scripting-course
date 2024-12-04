dim intNumber
dim strResults

intNumber = 38

if intNumber > 20 then

if intNumber = 25
  strResults = "The number is: " & intNumber

elseif intNumber = 40 then
strResults = " the number is: "& intNumber

else
strResults = "not the number we want [" & intNumber & "]"
end if
else 
strResults = "Number [" & intNumber & "] is not greater than 28"

end if

msgbox(strResults)