' Declare variables
Dim intAge
Dim strResults

' Assign value to intAge
intAge = 25

' Unary If (Single condition)
If intAge >= 18 Then
    strResults = "Adult"
End If

' Binary If (Two conditions)
If intAge >= 18 Then
    strResults = "Adult"
Else
    strResults = "Not an adult"
End If

' Multiple conditions (If-ElseIf-Else)
If intAge > 18 Then
    strResults = "Adult"
ElseIf intAge >= 13 Then
    strResults = "Teenager"
ElseIf intAge >= 0 Then
    strResults = "Kid"
Else
    strResults = "Age cannot be negative"
End If

' Display the results
MsgBox "Age: " & intAge & vbNewLine & " Results: " & strResults

