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
