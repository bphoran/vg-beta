Option Explicit 'Require variable declarations

Dim toarStr As Object            'String to hold TOAR screen
Dim toarStr_Created As Boolean   'Ensure string object created
                                 'Default value is False

Sub Copy_From_Clipboard()

If toarStr_Created = False Then
    Set toarStr = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    toarStr_Created = True
End If

'Put a string in the clipboard
'toarStr.SetText "Hello!"
'toarStr.PutInClipboard

'Get a string from the clipboard
toarStr.GetFromClipboard
'Debug.Print toarStr.GetText
'ActiveSheet.Range("A10").Value = toarStr.GetText

'Worksheets("BETA").Shapes("Clipboard").TextFrame.Characters.Text = CmtObj.GetText

End Sub

Sub Parse_REX()

Call Copy_From_Clipboard

Dim vbaStr As String

vbaStr = toarStr.GetText
'ActiveSheet.Range("A10").Value = vbaStr.Chars(1)
MsgBox ":" & Mid(vbaStr, 1, 10)

End Sub

Sub Delete_Rows()

On Error Resume Next

 'Range("A2:A300").Select
 Range("A2:A500").Select

 Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete


End Sub

Sub Delete_Duplicates()

'ActiveSheet.Range("A1:D100").RemoveDuplicates Columns:=Array(1, 4), Header:=xlGuess
ActiveSheet.Range("A1:D500").RemoveDuplicates Columns:=Array(1, 4), Header:=xlGuess

End Sub

Sub Button1_Click()
    Call Delete_Rows 'Macro1
    Call Delete_Duplicates 'Macro2
    
End Sub

Sub Compress()
    
    With Range("A:A")
        'Set c = .Find("new", LookIn:=xlValues)
        MsgBox .Find("new", After:=Range("A2"), LookIn:=xlValues).Value
    End With

End Sub
