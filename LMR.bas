Attribute VB_Name = "LMR"
'This functions returns the specified number of characters from string inside cell (the same as workbook function LEFT)
'But if output will be just number it will be number in cell. Ther is not necessery to convert cell to number.

'ALTERNATIVE FOR LEFT()
Public Function LEFT_NUM(cv As String, ch As Long)

cv = CStr(cv)
output = Left(cv, ch)

check = IsNumeric(output)

    If check = True Then
    
        LEFT_NUM = CInt(output)
    
    Else
    
        LEFT_NUM = output
        
    End If


End Function

'ALTERNATIVE FOR MID()
Public Function MID_NUM(cv As String, st As Long, ch As Long)

cv = CStr(cv)
output = Mid(cv, st, ch)

check = IsNumeric(output)

    If check = True Then
    
       MID_NUM = CInt(output)
    
    Else
    
        MID_NUM = output
        
    End If


End Function

'ALTERNATIVE FOR RIGHT()
Public Function RIGHT_NUM(cv As String, ch As Long)

cv = CStr(cv)
output = Right(cv, ch)

check = IsNumeric(output)

    If check = True Then
    
       RIGHT_NUM = CInt(output)
    
    Else
    
        RIGHT_NUM = output
        
    End If


End Function
