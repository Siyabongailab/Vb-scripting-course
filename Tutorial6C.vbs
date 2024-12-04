dim StrOperator

do until StrOperator = "+" Or StrOperator = "-" Or StrOperator = "*" Or StrOperator = "/"

StrOperator = Cstr(InputBox("please choose any of these operators: " & vbNewLine & _
                         "Addition: " & "+" & vbNewLine & _
                        "Subtraction: "& "-" & vbNewLine & _
                        "Multiply: " & "*" & vbNewLine & _
                        " Divide : " & "/"))
 Loop 



                      