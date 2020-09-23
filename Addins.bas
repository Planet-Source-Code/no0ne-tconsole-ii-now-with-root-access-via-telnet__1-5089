Attribute VB_Name = "AddIns"
'/////////////////
'// This .bas file is a compulation of various code
'// from the internet. ( it may have been redone to
'// suit my needs. ) Else it was left, as is.
'///


Sub StayOnTop(frm As Form)

SetWinOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Sub
Sub NOTOnTop(frm As Form)

SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

End Sub

Public Sub RebootComputer()
Dim intRetVal As Long
       intRetVal = ExitWindows(EW_REBOOTSYSTEM, 0)

 End Sub
 Public Sub CloseWindows()
Dim intRetVal As Long
       intRetVal = ExitWindows(EW_EXITWINDOWS, 0)

 End Sub
   Public Sub Hide_Program_In_CTRL_ALT_Delete()


       Dim pid As Long
       Dim reserv As Long
       pid = GetCurrentProcessId()
       regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
   End Sub



   Public Sub Show_Program_In_CTRL_ALT_DELETE()


       Dim pid As Long
       Dim reserv As Long
       pid = GetCurrentProcessId()
       regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
   End Sub
Public Function Append2File(CmdLine As String, CmdFile As String)

       Dim fnum As Long
       Dim Txt As String
       fnum = FreeFile
       Open CmdFile For Append As fnum
       Txt = CmdLine
       Print #fnum, Txt
       Close fnum
End Function



   Public Function GetFromINI(Section As String, ByVal Key As String, Directory As String) As String
   'Call WriteToINI("Header", "Color", "Black", "C:/Windows/whatever
   'Stuff$ = GetFromINI("Header", "Color", "C:/Windows/whatever.ini")
   

       Dim strBuffer As String
       strBuffer = String(750, Chr(0))
       Key$ = LCase$(Key$)
       GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
   End Function



   Public Sub WriteToINI(Section As String, ByVal Key As String, ByVal KeyValue As String, Directory As String)


       Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
   End Sub


Function GetUser() As String

'gsUserId = ClipNull(GetUser())
       Dim lpUserID As String
       Dim nBuffer As Long
       Dim Ret As Long
       lpUserID = String(25, 0)
       nBuffer = 25
       Ret = GetUserName(lpUserID, nBuffer)


       If Ret Then
           GetUser$ = lpUserID$
       End If


   End Function

   Function ClipNull(InString As String) As String

       Dim intpos As Integer

       If Len(InString) Then
           intpos = InStr(InString, vbNullChar)

           If intpos > 0 Then
               ClipNull = Left(InString, intpos - 1)
           Else
               ClipNull = InString
           End If

       End If

   End Function



