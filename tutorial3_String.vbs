Tutorial 3 - REM This Tutorial teaches about string functions
'*************
    ' Tutorial 3 - String Functions
'**************

'Declare
    dim strFirstNname, strLastname
    dim strFullname 
    dim intStringPosition
    dim strSearchString
    dim strCoursename, strCourseInput
    dim strStringCompare

'Assign
strFirstNname = Trim(InputBox("Please enter your first name!"))
strLastname = Trim(InputBox("Please enter your Last Name! "))

strFullname = Ucase(strFirstNname & " " & strLastname)

strCoursename = "VB Scripting"




'Use/Consume
msgBox("First Name: " & strFirstNname & vbNewLine _
                    & "Last Name: " & strLastname & vbNewLine _
                    & "======================" & vbNewLine _
                    & "Full Name : " & strFullname & vbNewLine _
                    & strFirstNname &  " has " & Len(strFirstNname) & " Characters" & vbNewLine _
                    & "The first 3 characters of " & strFirstNname & " are: " & Left(strFirstNname, 3))
   strSearchString = InputBox("Type the letter to search from the name (" & strFirstNname & ")" )
   intStringPosition = InStr(1, strFirstNname, strSearchString)

'Zero means it doesn't exist

   msgBox(strSearchString & " Is found at position: " & intStringPosition)

   strCourseInput = InputBox("Guess the name of the course :) ")
   strStringCompare= strComp(strCoursename, strCourseInput)
   msgBox("Just comapred " & strCourseInput & " and " & strCoursename & " is: " & strStringCompare)


