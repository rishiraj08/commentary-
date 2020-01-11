Sub test()
'SearchString As String, Char As String, Instance As Long
     'Function purpose:  To return the position of the (first character of the )
     'nth occurance of a specific character set in a range
    Dim x As Integer, n As Long
     
    
    
    With ThisWorkbook.Sheets("Sheet1")
        Set Rng = .Range("A2:A" & .Range("A" & Rows.Count).End(xlUp).Row)
        Rng.Font.ColorIndex = xlAutomatic
     End With
     
    For Each rCell In Rng
        Char = rCell.Offset(, 2)
        SearchString = rCell
        Instance = rCell.Offset(, 1)
     'Loop through each letter in the search string
        For x = 1 To Len(SearchString)
             
             'check if the next character(s) match the text being search for
             'and increase n if so (to count how many matches have been found
             loopChar = Mid(SearchString, x, Len(Char))
             Debug.Print loopChar
            If UCase(loopChar) = UCase(Char) Then
               If Instance = 0 Then
                    rCell.Characters(Start:=x - 1, Length:=Len(Char) + 1).Font.Color = -1003520
                    n = 0
                    Exit For
               Else
                    XplaceHolder = x
                    n = n + 1
               End If
            End If
             
             'Exit loop if instance matches number found
            If n = Instance Then
                rCell.Characters(Start:=x - 1, Length:=Len(Char) + 1).Font.Color = -1003520
                n = 0
                Exit For
            ElseIf x = Len(SearchString) Then
                rCell.Characters(Start:=XplaceHolder, Length:=Len(Char)).Font.Color = -1003520
                n = 0
            End If
        Next x
     Next rCell
     'The error below will only be triggered if the function was not
     'already exited due to success
    CharPos = CVErr(xlErrValue)
End Sub
