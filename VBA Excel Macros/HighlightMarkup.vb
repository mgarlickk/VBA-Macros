Public Sub ChgTxtColor2()
    Set myRange = Range("A1:A1000")  'The Range that contains the substring you want to change color
    stStr = "<"   'The first character of the substring you want to change color
    edStr = ">"   'The last character of the substring you want to change color
    txtColor = 3   'The ColorIndex which repsents the color you want to change, 3 represents Red
    
    On Error Resume Next
    
    For Each myString In myRange
        Dim stStrArr(), edStrArr()
            For i = 1 To Len(myString)
                tempString = Mid(myString, i, 1)
                If tempString = stStr Then
                    counter = counter + 1
                    ReDim Preserve stStrArr(counter)
                    stStrArr(counter) = i
                End If
            Next i
            
            For j = 1 To Len(myString)
                tempString2 = Mid(myString, j, 1)
                If tempString2 = edStr Then
                    counter2 = counter2 + 1
                    ReDim Preserve edStrArr(counter2)
                    edStrArr(counter2) = j
                End If
            Next j
            
            For k = 1 To counter
                If edStrArr(k) > stStrArr(k) Then
                    myString.Characters(Start:=stStrArr(k), Length:=edStrArr(k) - stStrArr(k) + 1).Font.ColorIndex = txtColor
                End If
            Next k
            
        Erase stStrArr()
        Erase edStrArr()
    Next myString
End Sub
