DIM intDayNUmber
DIM strDay

' Assign
intDayNUmber = 5

SELECT CASE intDayNUmber
CASE 1
strDay =" MONDAY"

CASE 2
strDay = "TUESDAY"
CASE 3
strDay = "WEDNESDAY"
CASE 4
strDay = "THURSDAY"
CASE 5
strDay = "FRIDAY"
CASE 6
strDay = "SATURDAY"
CASE 7
strDay = "SUNDAY"
CASE ELSE strDay = "INCORRECT NUMBER: [" & intDayNUmber & "]"

END SELECT
msgbox(strDay)
