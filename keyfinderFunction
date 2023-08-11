Function keyfinder(txt, startRow, col)
    ret = -1
    exitFor = 0
    For i = startRow To 2851
        flixWord = Worksheets(3).Cells(i, col)
        countwords = Len(flixWord) - Len(Replace(flixWord, ",", "")) + 1
        commaPos = 0
        Words = ""
        strt = 0
        For j = 1 To countwords
            strt = commaPos + 1
            If InStr(strt, flixWord, ",") = 0 Then
                leng = Len(flixWord)
                commaPos = leng + 1
            Else
                leng = InStr(strt, flixWord, ",") - strt
                commaPos = leng + commaPos + 1
            End If
            Words = Mid(flixWord, strt, leng)
            If Len(Words) > 0 And InStr(1, UCase(txt), UCase(Words), vbTextCompare) > 0 Then
                ret = Worksheets(3).Cells(i, 1) & " (" & i & ")"
                exitFor = 1
                Exit For
            End If
        Next j
        If exitFor = 1 Then Exit For
    Next i
    keyfinder = ret
End Function
