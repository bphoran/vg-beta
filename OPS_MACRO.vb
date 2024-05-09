'ToaMenuModule
'
'Author: Brendan Horan
'Last Update: 12/09/2020
 
Option Explicit
 
Public Const NOT_LOGGED_IN = -2
Public Const PULL_ERROR = -1
 
Const AREX_BUCKET_SIZE As Integer = 50  'Default size of dynamic array sArexAccts
Const AREX_CONTROL_COL As Integer = 4   'Control number column on AREX sheet
Const TOAR_CONTROL_COL As Integer = 10  'Control number column on remaining sheets
 
Dim sArexAccts() As String      'Dynamic array holding unique VBAs from TOAR *REX
Dim bControlFound() As Boolean  'Array that will indicate if the control number was found
Dim intArexSize As Integer      'Capacity of sArexAccts
Dim intArexCount As Integer     'Count of unique AREX VBAs
 
'Counts of new records
Public intREX As Integer
Public intREJ As Integer
Public intRCL As Integer
Public intREQR As Integer
Public intRHRD As Integer
 
'PullNewRecords
'
'Run through each screen, grabbing and filtering new records
'
'Return values less than zero indicate failure
Public Function PullNewRecords(TpxLink) As Integer
   
    Dim bPullToday As Boolean
   
    intREJ = 0
    intRHRD = 0
   
    PullNewRecords = PULL_ERROR 'Default error code
   
    'Check if user logged in
    If TpxModule.CheckLoginStatus(TpxLink) < 0 Then
        PullNewRecords = NOT_LOGGED_IN
        Exit Function
    End If
   
    bPullToday = RemoveOldRecords(shArex, AREX_CONTROL_COL, 7)
    bPullToday = RemoveOldRecords(shReclaim, TOAR_CONTROL_COL, 7)
    bPullToday = RemoveOldRecords(shSoftReject, TOAR_CONTROL_COL, 14)
    bPullToday = RemoveOldRecords(shRedundantHard, TOAR_CONTROL_COL, 90)
    bPullToday = RemoveOldRecords(shHardReject, TOAR_CONTROL_COL, 14)
   
    intREX = PullArex(TpxLink)
    intRCL = PullStatus(TpxLink, "TOAD *MENU-RCL", "", "", True, shReclaim)
   intREQR = PullStatus(TpxLink, "TOAR *MENU-REQR", "REQR", "REV", False, shSoftReject)
   
    'Only check Hards once a day
    If Not bPullToday Then
        intREJ = PullStatus(TpxLink, "TOAR *MENU-REJ", "REJ", "SEND", False, shHardReject)
       
        'Adjust hard reject count by removing redundant rejects
        intREJ = intREJ - ExtractRedundantHardRejects(intRHRD)
    End If
   
    PullNewRecords = 0
   
End Function
 
'PullStatus
'
'Function will copy records from the given screen, with the given status. sPostStat
'can be used to indicate a status that always comes after the desired status, thus
'halting the function.
'
'With sStat and sPostStat blank, and bValidBlank True, the function will
'grab all records from sScreen where VALID is blank
'
'bValidBlank = True, means VALID field MUST be blank
'
'Warning: If sPostStat is not present on BETA, there will be an infinite loop
'
'Returns: Number of new records found
Private Function PullStatus(TpxLink, sScreen As String, sStat As String, _
    sPostStat As String, bValidBlank As Boolean, shOutput As Worksheet) As Integer
   
    Dim intToaRow As Integer
    Dim intLastPullSheetRow As Integer
    Dim intNewSheetRow As Integer
    Dim sToaC As String
    Dim sToaStat As String
    Dim sPullTime As String
   
    With shOutput
       
        'Find next empty row on shOutput, header on first row
        intLastPullSheetRow = .Cells(.Rows.Count, TOAR_CONTROL_COL).End(xlUp).row
        intNewSheetRow = intLastPullSheetRow + 1
       
        'Clear screen and go to sScreen
        TpxModule.HostOnline TpxLink, sScreen
       
        'Add Pull Time
        sPullTime = TpxLink.GetScreen2(1, 55, 11) & ":00 ET"
       
        If .Range("A:A").Find(sPullTime) Is Nothing Then
            .Range("A" & intNewSheetRow).Value = sPullTime
        End If
       
        'Loop until sPostStat is found in Stat
        Do While Trim(TpxLink.GetScreen2(5, 18, 4)) <> sPostStat
            For intToaRow = 5 To 21
               
                sToaC = Trim(TpxLink.GetScreen2(intToaRow, 16, 1))
                sToaStat = Trim(TpxLink.GetScreen2(intToaRow, 18, 4))
               
                'Check if there are no more records on sScreen
                If sToaStat = "" Then
                   Exit Do
               
                'Ignore if incorrect VALID
                ElseIf bValidBlank And _
                    Trim(TpxLink.GetScreen2(intToaRow, 69, 8)) <> "" Then
               
                'Check for correct status
                ElseIf (sStat = "" Or sToaStat = sStat) Then
                   
                    'TODO: Remove account 0000-0000
                    'If TpxLink.GetScreen2(intToaRow, 6, 9) <> "0000-0000" Then
                        ImportRecord TpxLink, intToaRow, .Range(.Cells(intNewSheetRow, 1), _
                            .Cells(intNewSheetRow, TOAR_CONTROL_COL + 1))
                       
                        intNewSheetRow = intNewSheetRow + 1
                    'End If
                End If
               
            Next intToaRow
           
            TpxModule.HostEnter TpxLink
           
            'Check if all pages on sScreen have been viewed
            If TpxLink.GetScreen2(3, 78, 3) = "001" Then
                Exit Do
            End If
        Loop
       
        'Remove duplicates based on CONTROL
        .Range(.Cells(2, 2), .Cells(intNewSheetRow, TOAR_CONTROL_COL + 1)).RemoveDuplicates _
            Columns:=Array(TOAR_CONTROL_COL - 1)
        
        'Count only new records
        PullStatus = .Cells(.Rows.Count, TOAR_CONTROL_COL).End(xlUp).row - intLastPullSheetRow
       
        'Clear formatting on cells that had duplicates
        .Range(.Cells(intLastPullSheetRow + PullStatus + 1, 1), _
            .Cells(intNewSheetRow, 1)).EntireRow.Delete
       
    End With
   
End Function
 
'PullArex
'
'Find the TOAR *MENU-REV records corresponding to the AREX accounts
'found on TOAR *REX. Will ignore records no longer in Review status.
'
'Returns: Number of new AREX control numbers found
Private Function PullArex(TpxLink) As Integer
   
    Dim intArexRemaining As Integer
    Dim intArexIndex As Integer
    Dim intToaRow As Integer
    Dim intLastPullSheetRow As Integer
    Dim intNewSheetRow As Integer
    Dim sVBA As String
    Dim sToaC As String
    Dim sToaStat As String
    Dim sPullTime As String
   
    'Grab unique AREX account numbers from TOAR *REX
    intArexRemaining = GrabArexAccounts(TpxLink)
   
    'Exit if error occurred
    If intArexRemaining < 0 Then
        Exit Function
    End If
   
    'Exit if no AREX accounts
    If intArexRemaining = 0 Then
        PullArex = 0
        Exit Function
    End If
   
    With shArex
       
        'Find next empty row on shArex, header on first row
        intLastPullSheetRow = .Cells(.Rows.Count, AREX_CONTROL_COL).End(xlUp).row
        intNewSheetRow = intLastPullSheetRow + 1
       
        'Clear screen and go to sScreen
        TpxModule.HostOnline TpxLink, "TOAR *MENU-REV"
       
        'Add Pull Time
        sPullTime = TpxLink.GetScreen2(1, 55, 11) & ":00 ET"
       
        If .Range("A:A").Find(sPullTime) Is Nothing Then
            .Range("A" & intNewSheetRow).Value = sPullTime
        End If
       
        sToaStat = Trim(TpxLink.GetScreen2(5, 18, 4))
       
        Do While (sToaStat = "REV") Or (sToaStat = "RVAD")
            For intToaRow = 5 To 21
               
                sToaC = Trim(TpxLink.GetScreen2(intToaRow, 16, 1))
                sToaStat = Trim(TpxLink.GetScreen2(intToaRow, 18, 4))
               
                'Ignore if incorrect C-field
                'Check if there are no more records on sScreen
                If (sToaC <> "E") Or (sToaStat = "") Then
               
                'Ignore if VALID not blank, and not PTF/PTR
                ElseIf Trim(TpxLink.GetScreen2(intToaRow, 69, 8)) <> "" And _
                    Trim(TpxLink.GetScreen2(intToaRow, 36, 3)) <> "PTF" And _
                    Trim(TpxLink.GetScreen2(intToaRow, 36, 3)) <> "PTR" Then
               
                'Check for correct status
                ElseIf (sToaStat = "REV") Or (sToaStat = "RVAD") Then
                   
                    sVBA = TpxLink.GetScreen2(intToaRow, 6, 9)
                   
                    'Check if VBA was on TOAR *REX
                    For intArexIndex = 0 To (intArexCount - 1)
                       
                        If sArexAccts(intArexIndex) = sVBA Then
                       
                            'Format as text
                            .Range(.Cells(intNewSheetRow, 1), _
                                .Cells(intNewSheetRow, AREX_CONTROL_COL)).NumberFormat = "@"
                       
                            'Store Account
                            .Cells(intNewSheetRow, AREX_CONTROL_COL - 2).Value = sVBA
                           
                            'Text for robot
                            .Cells(intNewSheetRow, AREX_CONTROL_COL - 1).Value = "2 MUTUAL FUND"
                           
                            'Store CONTROL
                            .Cells(intNewSheetRow, AREX_CONTROL_COL).Value = _
                                TpxLink.GetScreen2(intToaRow, 45, 14)
                           
                            intNewSheetRow = intNewSheetRow + 1
                           
                            Exit For
                           
                        End If
                       
                    Next intArexIndex
                   
                End If
               
            Next intToaRow
           
            TpxModule.HostEnter TpxLink
           
            'Check if all pages on sScreen have been viewed
            If TpxLink.GetScreen2(3, 78, 3) = "001" Then
                Exit Do
            End If
        Loop
       
        'Remove duplicates based on ACCOUNT and CONTROL number
        .Range(.Cells(2, 2), .Cells(intNewSheetRow, AREX_CONTROL_COL)).RemoveDuplicates _
            Columns:=Array(AREX_CONTROL_COL - 3, AREX_CONTROL_COL - 1)
       
        'Count only new records
        PullArex = .Cells(.Rows.Count, AREX_CONTROL_COL).End(xlUp).row - intLastPullSheetRow
       
        'Clear formatting on cells that had duplicates
        .Range(.Cells(intLastPullSheetRow + PullArex + 1, 1), _
            .Cells(intNewSheetRow, 1)).EntireRow.Delete
       
    End With
   
End Function
 
'PullArex2 - Discontinued
'
'Find the TOAR *MENU-REV records corresponding to the AREX accounts
'found on TOAR *REX. Will ignore records no longer in Review status.
'
'Returns: Number of new AREX control numbers found
Private Function PullArex2(TpxLink) As Integer
   
    Dim intArexRemaining As Integer
    Dim intAcctIndex As Integer
    Dim intToarRow As Integer
    Dim sPrevLastToarRow As String
    Dim sLastToarRow As String
    Dim bAcctFound As Boolean
    Dim intLastPullSheetRow As Integer
    Dim intNewSheetRow As Integer
    Dim sPullTime As String
   
    'Grab unique AREX account numbers from TOAR *REX
    intArexRemaining = GrabArexAccounts(TpxLink)
   
    'Boolean array defaults to Falses
    ReDim bControlFound(intArexCount)
   
    'Exit if error occurred
    If intArexRemaining < 0 Then
        Exit Function
    End If
   
    'Exit if no AREX accounts
    If intArexRemaining = 0 Then
        PullArex2 = 0
        Exit Function
    End If
   
    With shArex
       
        'Find next empty row on shArex, header on first row
        intLastPullSheetRow = .Cells(.Rows.Count, AREX_CONTROL_COL).End(xlUp).row
        intNewSheetRow = intLastPullSheetRow + 1
       
        'Clear screen and go to TOAR *MENU-REV
        TpxModule.HostOnline TpxLink, "TOAR *MENU-REV"
       
        'Add Pull Time
        sPullTime = TpxLink.GetScreen2(1, 55, 11) & ":00 ET"
       
        If .Range("A:A").Find(sPullTime) Is Nothing Then
            .Range("A" & intNewSheetRow).Value = sPullTime
        End If
       
        bAcctFound = False
        sLastToarRow = TpxLink.GetScreen2(21, 6, 9)
       
        'Loop through all AREX account numbers
        For intAcctIndex = 0 To (intArexCount - 1)
           
            'Simple check to prevent infinite loop
            Do While TpxLink.GetScreen2(5, 18, 3) = "REV"
               
                'Check to see if last account on page is less than current AREX account
                If StrComp(sLastToarRow, sArexAccts(intAcctIndex)) < 0 Then
                   
                    TpxModule.HostEnter TpxLink
                    sPrevLastToarRow = sLastToarRow
                    sLastToarRow = TpxLink.GetScreen2(21, 6, 9)
                   
                    'Make sure account numbers continue to rise, otherwise the final
                    'page of exceptions has been reached
                    If StrComp(sLastToarRow, sPrevLastToarRow) < 0 Then
                        Exit Do
                    End If
                   
                Else
                    Exit Do
                   
                End If
                sLastToarRow = TpxLink.GetScreen2(21, 6, 9)
            Loop
           
            'Either correct page or last page was found
            For intToarRow = 5 To 21
               
                'Check VBA, C="E" and VALID is blank
                If TpxLink.GetScreen2(intToarRow, 6, 9) = sArexAccts(intAcctIndex) And _
                    TpxLink.GetScreen2(intToarRow, 16, 1) = "E" And _
                    Trim(TpxLink.GetScreen2(intToarRow, 69, 8)) = "" Then
                   
                    'Format as text
                    .Range(.Cells(intNewSheetRow, 1), _
                        .Cells(intNewSheetRow, AREX_CONTROL_COL)).NumberFormat = "@"
                   
                    'Store Account
                    .Cells(intNewSheetRow, AREX_CONTROL_COL - 2).Value = _
                        TpxLink.GetScreen2(intToarRow, 6, 9)
                   
                    'Text for robot
                    .Cells(intNewSheetRow, AREX_CONTROL_COL - 1).Value = "2 MUTUAL FUND"
                   
                    'Store CONTROL
                    .Cells(intNewSheetRow, AREX_CONTROL_COL).Value = _
                        TpxLink.GetScreen2(intToarRow, 45, 14)
                   
                    intNewSheetRow = intNewSheetRow + 1
                    bAcctFound = True
                    bControlFound(intAcctIndex) = True
                    intArexRemaining = intArexRemaining - 1
                Else
                    If bAcctFound Then
                        bAcctFound = False
                        Exit For
                    End If
                End If
               
                If intToarRow = 21 And bAcctFound Then
                    TpxModule.HostEnter TpxLink
                    intToarRow = 4
                End If
            Next intToarRow
           
        Next intAcctIndex
       
        'Check TOAR *MENU-RVAD for un-validated, AREX transfers
       
        'Clear screen and go to TOAR *MENU-RVAD
        TpxModule.HostOnline TpxLink, "TOAR *MENU-RVAD"
       
        For intToarRow = 5 To 21
           
            'Only continue for RVAD status
            If TpxLink.GetScreen2(intToarRow, 18, 4) <> "RVAD" Then
                Exit For
            End If
               
            'Check C="E" and VALID is blank
            If TpxLink.GetScreen2(intToarRow, 16, 1) = "E" And _
                Trim(TpxLink.GetScreen2(intToarRow, 69, 8)) = "" Then
               
                'Format as text
                .Range(.Cells(intNewSheetRow, 1), _
                    .Cells(intNewSheetRow, AREX_CONTROL_COL)).NumberFormat = "@"
               
                'Store Account
                .Cells(intNewSheetRow, AREX_CONTROL_COL - 2).Value = _
                    TpxLink.GetScreen2(intToarRow, 6, 9)
               
                'Text for robot
                .Cells(intNewSheetRow, AREX_CONTROL_COL - 1).Value = "2 MUTUAL FUND"
               
                'Store CONTROL
                .Cells(intNewSheetRow, AREX_CONTROL_COL).Value = _
                    TpxLink.GetScreen2(intToarRow, 45, 14)
               
                intNewSheetRow = intNewSheetRow + 1
            End If
           
            If intToarRow = 21 Then
                TpxModule.HostEnter TpxLink
                intToarRow = 4
            End If
        Next intToarRow
       
        'Remove duplicates based on ACCOUNT and CONTROL number
        .Range(.Cells(2, 2), .Cells(intNewSheetRow, AREX_CONTROL_COL)).RemoveDuplicates _
            Columns:=Array(AREX_CONTROL_COL - 3, AREX_CONTROL_COL - 1)
       
        'Count only new records
        PullArex2 = .Cells(.Rows.Count, AREX_CONTROL_COL).End(xlUp).row - intLastPullSheetRow
       
        'Clear formatting on cells that had duplicates
        .Range(.Cells(intLastPullSheetRow + PullArex2 + 1, 1), _
            .Cells(intNewSheetRow, 1)).EntireRow.Delete
       
        'If the page scrolling went past the "REV" status, add any remaining AREX accounts
        'whose control numbers were not found. In this rare situation, there may be
        'accounts listed whose transfers were actually validated previously.
        If TpxLink.GetScreen2(5, 18, 3) <> "REV" And intArexRemaining > 0 Then
           
            'Find next empty row on shArex, header on first row
            intNewSheetRow = .Cells(.Rows.Count, AREX_CONTROL_COL).End(xlUp).row + 1
           
            For intAcctIndex = (intArexCount - 1) To (intArexCount - intArexRemaining) Step -1
               
                'Once we come to an account whose control number WAS found, we no longer
                'need to list accounts we thought were missing
                If bControlFound(intAcctIndex) Then
                    Exit For
                End If
               
                'Format as text
                .Range(.Cells(intNewSheetRow, 1), _
                    .Cells(intNewSheetRow, AREX_CONTROL_COL)).NumberFormat = "@"
               
                'Store Account
                .Cells(intNewSheetRow, AREX_CONTROL_COL - 2).Value = sArexAccts(intAcctIndex)
               
                'Text for robot
                .Cells(intNewSheetRow, AREX_CONTROL_COL - 1).Value = "2 MUTUAL FUND"
               
                'Store CONTROL
                .Cells(intNewSheetRow, AREX_CONTROL_COL).Value = "??????????????"
               
                intNewSheetRow = intNewSheetRow + 1
               
            Next intAcctIndex
        End If
       
    End With
   
End Function
 
'GrabArexAccounts
'
'Grab unique AREX account numbers from TOAR *REX
'
'Update: Capture cusips for Fund Access
'
'Returns: Number of accounts found
Private Function GrabArexAccounts(TpxLink) As Integer
   
    Dim intToarRow As Integer
    Dim sPrevAcct As String
    Dim sNewAcct As String
    Dim sDate As String
    Dim intNewFASheetRow As Integer 'Row on Fund Access sheet
   
    intArexSize = AREX_BUCKET_SIZE
    ReDim sArexAccts(intArexSize)
    intArexCount = 0
   
    'Clear screen and go to TOAR *REX
    TpxModule.HostOnline TpxLink, "TOAR *REX"
   
    'Check if no exceptions exist
    If TpxLink.GetScreen2(24, 2, 19) = "NO EXCEPTIONS FOUND" Then
        GrabArexAccounts = 0
        Exit Function
    End If
   
    sPrevAcct = ""
   
    sDate = TpxLink.GetScreen2(1, 55, 8)
   
    'Find next empty row on shFundAccess, header on first row
    intNewFASheetRow = shFundAccess.Cells(shFundAccess.Rows.Count, 1).End(xlUp).row + 1
   
    'Scan each line, save account number if new
    Do While True
        For intToarRow = 5 To 21
            'Dash indicates line with account number
            If TpxLink.GetScreen2(intToarRow, 10, 1) = "-" Then
                sNewAcct = TpxLink.GetScreen2(intToarRow, 6, 9)
               
                'Format as text
                shFundAccess.Range(shFundAccess.Cells(intNewFASheetRow, 1), _
                    shFundAccess.Cells(intNewFASheetRow, 4)).NumberFormat = "@"
               
                'Store Date
                shFundAccess.Cells(intNewFASheetRow, 1).Value = sDate
               
                'Store Account
                shFundAccess.Cells(intNewFASheetRow, 2).Value = Replace(sNewAcct, "-", "")
               
                'Store Description
                shFundAccess.Cells(intNewFASheetRow, 3).Value = Trim(Replace(TpxLink.GetScreen2(intToarRow, 39, 24), "¬", ""))
               
                'Store cusip
                shFundAccess.Cells(intNewFASheetRow, 4).Value = TpxLink.GetScreen2(intToarRow, 64, 9)
               
                intNewFASheetRow = intNewFASheetRow + 1
               
                'Is this a new account number?
                If sNewAcct <> sPrevAcct Then
                    sArexAccts(intArexCount) = sNewAcct
                    intArexCount = intArexCount + 1
                    sPrevAcct = sNewAcct
                End If
            End If
        Next intToarRow
       
        'Check if this is the last page
        If TpxLink.GetScreen2(24, 2, 12) = "END OF FILE." Then
            GrabArexAccounts = intArexCount
            Exit Do
       End If
       
        'Increase array size if necessary
        If (intArexSize - intArexCount) < 20 Then
            intArexSize = intArexSize + AREX_BUCKET_SIZE
            ReDim Preserve sArexAccts(intArexSize)
        End If
       
        TpxModule.HostEnter TpxLink
       
        'Check if no exceptions exist, may occur if the final AREX
        'record was on the very last line of the previous page
        If TpxLink.GetScreen2(24, 2, 19) = "NO EXCEPTIONS FOUND" Then
            GrabArexAccounts = intArexCount
            Exit Do
        End If
   
    Loop
   
    With shFundAccess
       
        'Remove duplicate info for the past 100 rows
        If intNewFASheetRow > 100 Then
            .Range(.Cells(intNewFASheetRow - 100, 1), .Cells(intNewFASheetRow - 1, 4)).RemoveDuplicates _
                Columns:=Array(1, 3, 4)
        Else
            .Range(.Cells(1, 1), .Cells(intNewFASheetRow, 4)).RemoveDuplicates _
                Columns:=Array(1, 3, 4)
        End If
    End With
   
End Function
 
'ImportRecord
'
'Copy row from BETA and paste onto output range
Private Sub ImportRecord(TpxLink, intRow As Integer, ByRef rngOutput As Range)
   
    With rngOutput
       
        If (.Columns.Count < (TOAR_CONTROL_COL + 1)) Or (intRow < 5) Or (intRow > 21) _
            Or TpxLink.GetScreen2(4, 48, 7) <> "CONTROL" Then
            Exit Sub
        End If
       
        'Format as text
        .NumberFormat = "@"
       
        'Store Account
        .Cells(1, TOAR_CONTROL_COL - 8).Value = TpxLink.GetScreen2(intRow, 6, 9)
       
        'Store C
        .Cells(1, TOAR_CONTROL_COL - 7).Value = TpxLink.GetScreen2(intRow, 16, 1)
       
        'Store STAT
        .Cells(1, TOAR_CONTROL_COL - 6).Value = TpxLink.GetScreen2(intRow, 18, 4)
       
        'Store DY
        .Cells(1, TOAR_CONTROL_COL - 5).Value = TpxLink.GetScreen2(intRow, 23, 2)
       
        'Store BRKR
        .Cells(1, TOAR_CONTROL_COL - 4).Value = TpxLink.GetScreen2(intRow, 26, 4)
       
        'Store OCC
        .Cells(1, TOAR_CONTROL_COL - 3).Value = TpxLink.GetScreen2(intRow, 31, 4)
       
        'Store TRN
        .Cells(1, TOAR_CONTROL_COL - 2).Value = TpxLink.GetScreen2(intRow, 36, 3)
       
        'Store REP
        .Cells(1, TOAR_CONTROL_COL - 1).Value = TpxLink.GetScreen2(intRow, 40, 4)
       
        'Store CONTROL
        .Cells(1, TOAR_CONTROL_COL).Value = TpxLink.GetScreen2(intRow, 45, 14)
       
        'Store OPEN
        .Cells(1, TOAR_CONTROL_COL + 1).Value = TpxLink.GetScreen2(intRow, 60, 8)
   
    End With
       
End Sub
 
'ExtractRedundantHardRejects
'
'Filter hard reject records that have a control number found on the AREX or
'soft reject worksheet, or rep code 'Z001'
'
'Returns the number of redundant rejects deleted from the Hard Reject sheet,
'while intNewRedundantHards records the number of new additions to the
'Redundant Hard sheet
Public Function ExtractRedundantHardRejects(ByRef intNewRedundantHards As Integer) As Integer
   
    Dim intLastRejectRow As Integer     'On Hard Reject sheet
    Dim intRejectRow As Integer         'On Hard Reject sheet
    Dim intLastPullSheetRow As Integer  'On Redundant Hard Reject sheet
    Dim intNextOutputRow As Integer     'On Redundant Hard Reject sheet
    Dim bExtractReject As Boolean       'If Hard Reject is redundant
    Dim sPullTime As String
   
    Dim rngArexControls As Range
    Dim rngSoftControls As Range
   
    ExtractRedundantHardRejects = 0
   
    If intREJ <= 0 Then
        Exit Function
    End If
   
    intLastPullSheetRow = 1     'When sheet is blank, will provide correct count
   
    intNextOutputRow = shRedundantHard.Cells(shRedundantHard.Rows.Count, 2).End(xlUp).row + 1
   
    If intNextOutputRow > 2 Then
        intLastPullSheetRow = intNextOutputRow - 1
    End If
   
    'Grab the Control number columns from AREX and SOFT
    Set rngArexControls = shArex.Columns(AREX_CONTROL_COL)
    Set rngSoftControls = shSoftReject.Columns(TOAR_CONTROL_COL)
   
    With shHardReject
       
        sPullTime = .Cells(.Cells(.Rows.Count, 1).End(xlUp).row, 1).Value
       
        intLastRejectRow = .Cells(.Rows.Count, TOAR_CONTROL_COL).End(xlUp).row
        bExtractReject = False
       
        For intRejectRow = intLastRejectRow To (intLastRejectRow - intREJ + 1) Step -1
           
            If .Cells(intRejectRow, TOAR_CONTROL_COL - 1) = "Z001" Then
                bExtractReject = True
           
            ElseIf Not rngSoftControls.Find(.Cells(intRejectRow, TOAR_CONTROL_COL)) Is Nothing Then
                bExtractReject = True
               
            ElseIf Not rngArexControls.Find(.Cells(intRejectRow, TOAR_CONTROL_COL)) Is Nothing Then
                bExtractReject = True
               
            End If
           
            If bExtractReject Then
           
                'Copy row to Redundant Hard reject sheet
                .Range(.Cells(intRejectRow, 2), .Cells(intRejectRow, TOAR_CONTROL_COL + 1)).Copy _
                shRedundantHard.Range(shRedundantHard.Cells(intNextOutputRow, 2), _
                shRedundantHard.Cells(intNextOutputRow, TOAR_CONTROL_COL + 1))
               
                'Delete current row
                .Range(.Cells(intRejectRow, 2), _
                    .Cells(intRejectRow, TOAR_CONTROL_COL + 1)).Delete Shift:=xlShiftUp
               
                ExtractRedundantHardRejects = ExtractRedundantHardRejects + 1
                intNextOutputRow = intNextOutputRow + 1
               
            End If
           
            bExtractReject = False
           
        Next intRejectRow
       
    End With
   
    With shRedundantHard
       
        'Remove duplicates based on CONTROL number
        .Range(.Cells(2, 2), .Cells(intNextOutputRow, TOAR_CONTROL_COL + 1)).RemoveDuplicates _
            Columns:=Array(TOAR_CONTROL_COL - 1)
       
        'Count only new records
        intNewRedundantHards = .Cells(.Rows.Count, TOAR_CONTROL_COL).End(xlUp).row _
            - intLastPullSheetRow
       
        If intNewRedundantHards > 0 Then
            If .Range("A:A").Find(sPullTime) Is Nothing Then
                .Range("A" & (intLastPullSheetRow + 1)).Value = sPullTime
            End If
        End If
       
        'Clear formatting on cells that had duplicates
        .Range(.Cells(intLastPullSheetRow + intNewRedundantHards + 1, 1), _
            .Cells(intNextOutputRow, 1)).EntireRow.Delete
       
    End With
   
End Function
 
'RemoveOldRecords
'
'Delete records from shSheet that are older than intDays
'
'Return Whether this sheet has had a pull today
Private Function RemoveOldRecords(shSheet As Worksheet, intControlCol As Integer, intDays As Integer) As Boolean
   
    Dim intLastPullSheetRow As Integer
    Dim intLastOldRow As Integer
    Dim sPullTime As String
    Dim sLastPullTime As String
   
    RemoveOldRecords = False    'Default
   
    With shSheet
       
        intLastPullSheetRow = .Cells(.Rows.Count, intControlCol).End(xlUp).row
       
        If intLastPullSheetRow = 1 Then
            Exit Function
        End If
       
        sLastPullTime = .Cells(.Cells(.Rows.Count, 1).End(xlUp).row, 1).Value
       
        If sLastPullTime <> "Pull Time" Then
            If DateDiff("d", Format(Left(sLastPullTime, 8), "mm/dd/yy"), Now) = 0 Then
                RemoveOldRecords = True
            End If
        End If
       
        For intLastOldRow = 2 To intLastPullSheetRow
           
            sPullTime = .Range("A" & intLastOldRow).Value
            If sPullTime <> "" Then
                If DateDiff("d", Format(Left(sPullTime, 8), "mm/dd/yy"), Now) <= intDays Then
                    Exit For
                End If
            End If
           
        Next intLastOldRow
       
        If intLastOldRow > 2 Then
            .Range("A2:A" & (intLastOldRow - 1)).EntireRow.Delete
        End If
       
    End With
End Function
 
 
 
'TpxModule
'
'Author: Brendan Horan
'Last Update: 02/26/2020
 
Option Explicit
 
'CheckLoginStatus
'
'Test to determine if BETA is currently logged in
'
'Return values less than zero indicate BETA is not logged in
Public Function CheckLoginStatus(TpxLink) As Integer
   
    CheckLoginStatus = -1   'Default error code
   
    'Clear screen and go to TOAR, check if "CONTROL" is visible
    HostOnline TpxLink, "TOAR"
   
    If TpxLink.GetScreen2(4, 48, 7) = "CONTROL" Then
        CheckLoginStatus = 0
    End If
   
    HostOnline TpxLink, ""
   
End Function
 
'HostEnter
'
'Send Enter key push
Public Sub HostEnter(TpxLink)
 
    Dim sStatus As String
    Dim strStat As String
    Dim strGI As String
   
    TpxLink.SetFocus
    TpxLink.SendKeys "@E"   'Enter key
 
    ' Obtain current Host status
    sStatus = TpxLink.GetStatusLine(strStat, strGI)
   
    ' While Host status is not zero (Zero = Ready for new info), wait and re-check
    Do While sStatus <> 0
        sStatus = TpxLink.GetStatusLine(strStat, strGI)
        DoEvents
    Loop
 
End Sub
 
'HostLogoff
'
'Logoff of BETA
Public Sub HostLogoff(TpxLink)
   
    'Check user is still logged in
    If CheckLoginStatus(TpxLink) = 0 Then
        HostOnline TpxLink, "CESF LOGOFF"
        TpxLink.SendKeys "k"    'kill session
        Pause 1.5
        TpxLink.SendKeys "@E"   'Enter key
    End If
   
End Sub
 
'HostOnline
'
'Direct BETA to the screen provided, ex. "TOAR *MENU-REJ".
'If sScreen blank, Sub clears screen
Public Sub HostOnline(TpxLink, sScreen As String)
 
    Dim sStatus As String
   Dim strStat As String
    Dim strGI As String
   
    TpxLink.SetFocus
    TpxLink.SendKeys "@R"   'Reset, i.e. Ctrl
    Pause 0.75
    TpxLink.SendKeys "@C"   'Clear
   
    ' Obtain current Host status
    sStatus = TpxLink.GetStatusLine(strStat, strGI)
   
    ' While Host status is not zero (Zero = Ready for new info), wait and re-check
    Do While sStatus <> 0
        sStatus = TpxLink.GetStatusLine(strStat, strGI)
        DoEvents
    Loop
   
    If sScreen = "" Then
        Exit Sub
    End If
   
    TpxLink.PutScreen2 sScreen, 1, 1
    TpxLink.SendKeys "@E"   'Enter key
   
    ' Obtain current Host status
    sStatus = TpxLink.GetStatusLine(strStat, strGI)
   
    ' While Host status is not zero (Zero = Ready for new info), wait and re-check
    Do While sStatus <> 0
        sStatus = TpxLink.GetStatusLine(strStat, strGI)
        DoEvents
    Loop
   
End Sub
 
'Pause
'
'Pause macro for the given amount of seconds
Public Sub Pause(dblWaitTime As Double)
   
    Dim start: start = Timer
   
    Do While Timer < start + dblWaitTime
        DoEvents
    Loop
   
End Sub
