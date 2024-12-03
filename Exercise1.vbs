
DIM strFirstName
DIM strLastName

'inputting name
strFirstName = inputbox("what is your first name: ")
strFirstName = UCase(Left(strFirstName, 1)) & Mid(strFirstName, 2)

'inputing last name
strLastName = inputbox("what is your Last name: ")
strLastName = UCase(Left(strLastName, 1)) & Mid(strLastName, 2)

'Use
msgbox(strFirstName & " " & strLastName )
