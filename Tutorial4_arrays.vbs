'*******************************************
'             Arrays
'*******************************************

REM Array-> variable that can store multiple values of the same data type

' Declare
DIM arIntNumbers(5) 'empty array that will store 5 numbers
DIM arStrNames ' Declare array variable

' Assign values to arStrNames using the Array function
arStrNames = Array("John", "kelvin", "Siyabonga", "sizwe") 'array with initial values

' Assign values to arIntNumbers
arIntNumbers(0) = 45
arIntNumbers(1) = 38
arIntNumbers(2) = 67
arIntNumbers(3) = 76
arIntNumbers(4) = 12
'arIntNumbers(5)=21

' Use
MsgBox(arIntNumbers(3))


