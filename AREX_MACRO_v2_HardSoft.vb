Function HardRejectprocessImport()
    
    'Formula Sample to check Hard Reject output page referencing the SoftRejectOutput Tab
    '=VLOOKUP(J4,SoftRejectOutput!$I$1:$J$10000,1,FALSE)
    
    Dim ImportRow As Integer
    Dim outputRow As Integer
    Dim lastImportRow As Integer
    Dim lastOutputRow As Integer
    Dim iRow As Integer
    Dim match As Boolean
    ImportRow = 2
    outputRow = 3
    
    Application.ScreenUpdating = False
    
    'Find last row of import data
    lastImportRow = ImportRow
    While Sheets("HardRejectImport").Range("A" & lastImportRow).Value <> ""
        lastImportRow = lastImportRow + 1
    Wend
    lastImportRow = lastImportRow - 1
    
    lastOutputRow = outputRow
    While Sheets("HardRejectOutput").Range("B" & lastOutputRow).Value <> ""
        lastOutputRow = lastOutputRow + 1
    Wend
    lastOutputRow = lastOutputRow - 1
    outputRow = lastOutputRow + 1
    For ImportRow = 2 To lastImportRow
        Sheets("HardRejectOutput").Activate
        
        'Check B D J / A C I
        match = False
        For iRow = 3 To lastOutputRow
            If Sheets("HardRejectOutput").Range("B" & iRow).Value = Sheets("HardRejectImport").Range("A" & ImportRow).Value _
            And Sheets("HardRejectOutput").Range("D" & iRow).Value = Sheets("HardRejectImport").Range("C" & ImportRow).Value _
            And Sheets("HardRejectOutput").Range("J" & iRow).Value = Sheets("HardRejectImport").Range("I" & ImportRow).Value Then match = True
        Next iRow
        
        If match = False Then
            Sheets("HardRejectImport").Activate
            Range("A" & ImportRow & ":J" & ImportRow).Copy
            Sheets("HardRejectOutput").Activate
            Sheets("HardRejectOutput").Range("B" & outputRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("HardRejectImport").Activate
            outputRow = outputRow + 1
        End If
        
    Next ImportRow
    
    Application.ScreenUpdating = True
    
    Sheets("HardRejectOutput").Activate
    Range("A1").Select
    
    MsgBox "Done"
    
End Function


Function SoftRejectprocessImport()
    Dim ImportRow As Integer
    Dim outputRow As Integer
    Dim lastImportRow As Integer
    Dim lastOutputRow As Integer
    Dim iRow As Integer
    Dim match As Boolean
    ImportRow = 2
    outputRow = 3
    
    Application.ScreenUpdating = False
    
    'Find last row of Import data
    lastImportRow = ImportRow
    While Sheets("SoftRejectImport").Range("A" & lastImportRow).Value <> ""
        lastImportRow = lastImportRow + 1
    Wend
    lastImportRow = lastImportRow - 1
    
    lastOutputRow = outputRow
    While Sheets("SoftRejectOutput").Range("B" & lastOutputRow).Value <> ""
        lastOutputRow = lastOutputRow + 1
    Wend
    lastOutputRow = lastOutputRow - 1
    outputRow = lastOutputRow + 1
    For ImportRow = 2 To lastImportRow
        Sheets("SoftRejectOutput").Activate
        
        'Check B D J / A C I
        match = False
        For iRow = 3 To lastOutputRow
            If Sheets("SoftRejectOutput").Range("B" & iRow).Value = Sheets("SoftRejectImport").Range("A" & ImportRow).Value _
            And Sheets("SoftRejectOutput").Range("D" & iRow).Value = Sheets("SoftRejectImport").Range("C" & ImportRow).Value _
            And Sheets("SoftRejectOutput").Range("J" & iRow).Value = Sheets("SoftRejectImport").Range("I" & ImportRow).Value Then match = True
        Next iRow
        
        If match = False Then
            Sheets("SoftRejectImport").Activate
            Range("A" & ImportRow & ":J" & ImportRow).Copy
            Sheets("SoftRejectOutput").Activate
            Sheets("SoftRejectOutput").Range("B" & outputRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("SoftRejectImport").Activate
            outputRow = outputRow + 1
        End If
        
    Next ImportRow
    
    Sheets("SoftRejectOutput").Activate
    Range("A1").Select
    
    
    Application.ScreenUpdating = True
    
    MsgBox "Done"
    
End Function



Function ClearData()

Application.ScreenUpdating = False

    Sheets("SoftRejectImport").Select
    Range("A2:J60000").Select
    Selection.ClearContents
    
    Sheets("SoftRejectOutput").Select
    Range("B3:M60000").Select
    Selection.ClearContents
    

    Sheets("HardRejectImport").Select
    Range("A2:J60000").Select
    Selection.ClearContents
    
    Sheets("HardRejectOutput").Select
    Range("B3:M60000").Select
    Selection.ClearContents
    
    Sheets("SoftRejectImport").Select
    Range("A1").Select
    
Application.ScreenUpdating = True

End Function
