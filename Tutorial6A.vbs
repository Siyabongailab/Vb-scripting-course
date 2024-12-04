Dim StrOperator
Dim strContinue
Dim intNum1, intNum2, intTotal

Do
    ' Ask the user for the first number
    intNum1 = CInt(InputBox("Please enter the first integer:"))

    ' Ask the user to choose an operator (loop until valid operator is entered)
    StrOperator = "" ' Initialize StrOperator
    Do Until StrOperator = "+" Or StrOperator = "-" Or StrOperator = "*" Or StrOperator = "/"
        StrOperator = CStr(InputBox("Please choose one of these operators:" & vbNewLine & _
                                    "Addition: " & "+" & vbNewLine & _
                                    "Subtraction: " & "-" & vbNewLine & _
                                    "Multiplication: " & "*" & vbNewLine & _
                                    "Division: " & "/"))
    Loop

    ' Ask the user for the second number
    intNum2 = CInt(InputBox("Please enter the second integer:"))

    ' Perform the selected operation
    If StrOperator = "+" Then
        intTotal = intNum1 + intNum2
        MsgBox "The total is: " & intTotal
    ElseIf StrOperator = "-" Then
        intTotal = intNum1 - intNum2
        MsgBox "The total is: " & intTotal
    ElseIf StrOperator = "*" Then
        intTotal = intNum1 * intNum2
        MsgBox "The total is: " & intTotal
    ElseIf StrOperator = "/" Then
        ' Check for division by zero
        If intNum2 = 0 Then
            MsgBox "Cannot divide by zero!"
        Else
            intTotal = intNum1 / intNum2
            MsgBox "The total is: " & intTotal
        End If
    End If

    ' Ask the user if they want to continue
    strContinue = LCase(CStr(InputBox("Do you want to continue (YES or NO): ")))

Loop Until strContinue <> "yes"

' Final message when exiting the loop
MsgBox "The end."


