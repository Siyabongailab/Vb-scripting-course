'Tutorial 2.                                   '*************
         ' Tutorial 2 - Variables
'**************

REM This tutorial teaches you about differemt data types
Option Explicit

'Declare
dim strName
dim intAge
dim dblHeight
dim isMarried



'Assign
'Get this information from user
strName = InputBox("Please enter your name!")
intAge = Cint(InputBox("Please enter your age"))
dblHeight = Cdbl(InputBox("Please enter your Height"))
isMarried = CBool(Cint(InputBox("is married? " &vbNewLine _ 
                                                & "1 - Yes " & vbNewLine _ 
                                                & "0 - No")))

'Use/Consume
msgBox("Name: " & strName & vbNewLine _
                & "age: " & intAge & vbNewLine _
                & "Height: " & dblHeight & vbNewLine _
                & "Is Married? " & isMarried)