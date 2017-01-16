# excelshortcuts

'vba code where it will delete a line until the condition is met'

Do
        If Left(Sheets("Table 1").Range("A1").Value, 10) <> "  Organisa" Then
            Sheets("Table 1").Range("A1").EntireRow.Delete
        End If
    Loop Until Left(Sheets("Table 1").Range("A1").Value, 10) = "  Organisa"
