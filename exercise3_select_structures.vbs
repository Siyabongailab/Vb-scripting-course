dim TestMark, assignment1, assignment2, Classatt
dim strResults

' Initialize test marks, assignment scores, and attendance for testing
TestMark = 54
assignment1 = 55
assignment2 = 40
Classatt = 65

' Check TestMark to determine if the student can submit assignment 1
if TestMark > 50 then
    strResults = "Student can submit assignment 1"
    msgbox(strResults)
else
    strResults = "Student cannot submit assignment 1"
    msgbox(strResults)
end if

' Check if the student passed assignment 1
if assignment1 >= 55 then
    strResults = "Please submit assignment 2"
    msgbox(strResults)
else
    strResults = "You failed to pass assignment 1"
    msgbox(strResults)
end if

' Check if the student passed assignment 2 and evaluate class attendance if failed
if assignment2 >= 60 then
    strResults = "Student passed assignment 2"
    msgbox(strResults)
else
    ' Check class attendance
    Select Case True ' checks for all conditions that are true. not really attendence. bad approach
        Case Classatt >= 75
            strResults = "Class attendance is at least 75%"
            msgbox(strResults)
        Case Else
            strResults = "Poor class attendance from student"
            msgbox(strResults)
    End Select
end if


 
 