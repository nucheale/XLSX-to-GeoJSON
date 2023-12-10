Sub geojson()
Dim i, j, lastCell, cellLength As Integer
Dim file As String

Sheets(2).Select
Range("A:A").ClearContents


Sheets(1).Select
lastCell = Cells(1, 1).CurrentRegion.Cells(Cells(1, 1).CurrentRegion.Cells.Count).Row
    
For i = 1 To lastCell - 1
    Sheets(2).Cells(i, 1).Value = "{///type///: ///feature///, ///id///:" & Sheets(1).Cells(i + 1, 6) & ", ///geometry///: { ///type///: " & "///" & Sheets(1).Cells(i + 1, 1) & "///, " & "///coordinates///: " & "[" & Sheets(1).Cells(i + 1, 3) & "," & Sheets(1).Cells(i + 1, 2) & "]}," & "///properties///: { ///description///: ///" & Sheets(1).Cells(i + 1, 4) & "///," & "///iconCaption///: ///" & Sheets(1).Cells(i + 1, 5) & "///," & "///marker-color///: ///" & Sheets(1).Cells(i + 1, 7) & "///}}" & ","
Next i


Sheets(2).Select

cellLength = Len(Cells(lastCell - 1, 1).Value)
Cells(lastCell - 1, 1).Value = Mid(Cells(lastCell - 1, 1).Value, 1, cellLength - 1)

file = ActiveWorkbook.Path & "\" & "Result" & ".geojson"
Open file For Output As #1
Print #1, Cells(1, 3).Value
    
For i = 1 To lastCell - 1
    Cells(i, 1).Value = Replace(Cells(i, 1).Value, "///", Chr(34))
    Print #1, Cells(i, 1).Value
Next i

Print #1, Cells(2, 3).Value
Close 1
OriginalFile = ActiveWorkbook.Path & "\" & "Result" & ".geojson"
EncodedFile = ChangeFileCharset(OriginalFile, "UTF-8", "Windows-1251")

End Sub


Function ChangeFileCharset(ByVal filename$, ByVal DestCharset$, Optional ByVal SourceCharset$) As Boolean
    On Error Resume Next: Err.Clear
     With CreateObject("ADODB.Stream")
         .Type = 2
         If Len(SourceCharset$) Then .Charset = SourceCharset$
        .Open
         .LoadFromFile filename$
        FileContent$ = .ReadText
        .Close
         .Charset = DestCharset$
        .Open
         .WriteText FileContent$
         .SaveToFile filename$, 2
        .Close
     End With
     ChangeFileCharset = Err = 0
End Function
