DIM strFullName
DIM NameSpliter, strFirstName, strLastName

'assignment
strFullName = inputbox("what is your full name: ")
NameSpliter = Split(strFullName,  " ")

'Splitting the name
strFirstName =NameSpliter(0)
strLastName = NameSpliter(1)

'output
msgbox("First Name:  " & strFirstName & vbNewLine & " Last Name " & strLastName)





