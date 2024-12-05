
class Calculator

public sub welcomeMessage()

msgbox("welcome to our calculation app")

end sub

public function CalcSum(intNum1, intNum2)

CalcSum = intNum1 + intNUm2

end function

end class

'declaration
dim num1, Num2, dblResults
set objCalculator = new Calculator

' 'assignment
call objCalculator.welcomeMessage()

num1= CInt(inputbox("enter first number"))
num2= CInt(inputbox("enter second number"))

dblResults = objCalculator.CalcSum(num1,num2)

msgbox("Results: " & dblResults)