VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTconsole 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tconsole [ Telnet Control Server ]"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Control.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton button 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdkill 
      Caption         =   "Kill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   30
      Picture         =   "Control.frx":08CA
      ScaleHeight     =   675
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   2250
      Width           =   3495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lsTpanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C000&
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6495
   End
   Begin VB.Menu mnuControler 
      Caption         =   "Options"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconfig 
         Caption         =   "Setup Controler"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLoging 
      Caption         =   "Logging"
      Begin VB.Menu mnuClrList 
         Caption         =   "Clear Log"
      End
      Begin VB.Menu mnuSavelog 
         Caption         =   "Save Log"
      End
   End
End
Attribute VB_Name = "frmTconsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim FaiLer As Integer

Public Sub ResetConsole()

    Winsock1.Close
    Winsock1.LocalPort = LOCAL_PORT
    Winsock1.Listen
    
    LogedIn = False
    accept = False
    nameEntered = False
    logIn = ""
    pasS = ""
    FaiLer = 1
    
End Sub
 
Public Sub ClearLogwindow()

     lsTpanel.Clear

End Sub


Public Sub Button_Click()
'///////////////
'// This sub is called every time Tconsole
'// needs to be reset.
'///

    Dim tme As String
    tme = Now
    
If CheckIfInfoWasEntered = False Then
    frmConfig.Show
    Exit Sub
End If
 
   If button.Caption = "Start" Then
        lsTpanel.Clear
        lsTpanel.AddItem "Tconsole Started -- " & tme
        lsTpanel.AddItem "  -- On Local Host - " & Winsock1.LocalHostName & " (" & Winsock1.LocalIP & ")"
        lsTpanel.AddItem "  -- Ready for Telnet services on port " & LOCAL_PORT
        button.Caption = "Reset"
    Else
        lsTpanel.AddItem "Services Restarted -- " & tme
        
    End If
    
Call ResetConsole

End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdkill_Click()

    Call kilL
    
End Sub

Private Sub Form_Load()

    'StayOnTop Me
    Call Button_Click

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub mnuClrList_Click()

    lsTpanel.Clear

End Sub

Private Sub mnuconfig_Click()

    frmConfig.Show

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuKill_Click()

    Call kilL

End Sub



Private Sub mnuSavelog_Click()

    Call Savelog(lsTpanel)
    MsgBox "Log file has been saved.", vbOKOnly, "Log File"
    
End Sub

Private Sub mnuStart_Click()

    Call Button_Click

End Sub


Private Sub Winsock1_Close()

    lsTpanel.AddItem "-- Tconsole succesfuly shut down. "
    lsTpanel.AddItem "-- Restarting. . . . "
    Call ResetConsole
        
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim tme As String
Dim log As String

On Error GoTo shutDownSock

    If Winsock1.State <> sckClosed Then Winsock1.Close
        
    Winsock1.accept requestID
    
        tme$ = Now
        RemoteIP = Winsock1.RemoteHostIP
        log$ = "Remote User IP: " & Winsock1.RemoteHostIP & " -- " & tme$ & vbCrLf
        
    Winsock1.SendData BannerMsg$()
    Winsock1.SendData log$
    Winsock1.SendData "Login: "
    
        lsTpanel.AddItem "Access Attemp from " & RemoteIP
        lsTpanel.AddItem "--- " & tme$
        
Exit Sub
shutDownSock:
Call ResetConsole

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

'//////////////////////
'// Only three things can happen at this console
'// A.) Login B.) DOS Mode C.) Sending a command
'//
'//
'// If the login flag hasn't been set then
'// this first part will control the login
'// process.
'///
Dim commandSent As String
Dim tme As String
tme$ = Time

Winsock1.GetData rfC

If Not LogedIn Then
        
    ' -- When name has been entered the
    ' -- accept flag is then set to true,
    ' -- This is for redirecting then Name in
    ' -- one buffer and the Pass in another.
    
        If accept = True And rfC$ = vbCrLf Then
            If logIn$ = LOGIN_NAME And pasS$ = LOGIN_PASS Then
                
                '// -- The Famus UNIX bash prompt! --
                prompT$ = logIn$ & "@" & Winsock1.RemoteHostIP & ":/#"
                
                ' -- Send Welcome Screen
                Winsock1.SendData vbCrLf & " --------------- [ Welcome! ] ---------------- "
                Winsock1.SendData vbCrLf & " --     Type in ? or help for the menu      -- "
                Winsock1.SendData vbCrLf & " --------------------------------------------- "
                Winsock1.SendData vbCrLf & " --    Type in dos to switch to DOS Mode    -- "
                Winsock1.SendData vbCrLf & " --------------------------------------------- " & vbCrLf & vbCrLf
                
                Winsock1.SendData prompT
                
                ' -- Set Main IF Flag and log Da Login
                LogedIn = True
                lsTpanel.AddItem "  Login OK"
                
                ' -- We did the do now exit and
                ' -- await new key press
                Exit Sub
                
            Else
            
                ' -- Handle for failed Login --
                Winsock1.SendData vbCrLf & "Login Failed" & vbCrLf & vbCrLf & "Login: "
                
                Call MarkFailedAtempt ' -- Log ALL failed attempts
                                
                If BLOCK_ON = True Then ' -- If MAX then Disconnect
                    FaiLer = FaiLer + 1
                    If FaiLer = (MAX_ATTEMPTS + 1) Then
                        Call Button_Click
                    End If
                End If
                
                ' -- Reset for new connection
                accept = False
                logIn = ""
                pasS = ""
                Exit Sub
            End If
            
        ElseIf accept = False And rfC$ = vbCrLf Then
            accept = True
            Winsock1.SendData vbCrLf & "Pass: "
            Exit Sub
        End If
        
        If accept = False Then
            logIn = logIn & rfC$
        Else
            pasS = pasS & rfC
        End If
        
        Winsock1.SendData rfC$
        
'----------------------[ End Login ]-------------------------
'////////////////////////////////////////////////////////////
'// Start DOS Mode:
'////////////////////////////
'//
'// This is ONLY used after DOS_mode Flag has been set
'// the actual switching is called in the Main Menu. [Below]
'// This is the same process as the Main, But tooned
'// and worded for the DOS shell effect.
'//
'////


' -- Is DOS flag set --
ElseIf DOS_Mode Then
    
    If rfC = vbCrLf Then
        

        Dim Lngth As Long, DirString As String

        Select Case CommanD
                    
        Case "help"
            Call DosMenu
            CommanD = ""
            Exit Sub
                
         Case "?"
            Call DosMenu
            CommanD = ""
            Exit Sub
            
        Case "exit"
            DOS_Mode = False
            CommanD = ""
            prompT$ = logIn$ & "@" & Winsock1.RemoteHostIP & ":/#"
            Winsock1.SendData vbCrLf & vbCrLf & prompT
            lsTpanel.AddItem " -- Exiting DOS Mode"
            Exit Sub
    
        Case "dir"
            Lngth = Len(CPrompt(Cpointer))
            DirString$ = Left$(CPrompt(Cpointer), Lngth - 1) ' -- takes the  ">" out
            
            If Cpointer > 0 Then
                DirString$ = Left$(CPrompt(Cpointer), Lngth - 1) & "\"
            End If
            
            Call DirList(DirString$)
            CommanD = ""
            Exit Sub
        
        ' -- ls is UNIX/{Linux} version of dir.. keep in out of habit of using
        Case "ls"
            Lngth = Len(CPrompt(Cpointer))
            DirString$ = Left$(CPrompt(Cpointer), Lngth - 1)
            
            If Cpointer > 0 Then
                DirString$ = Left$(CPrompt(Cpointer), Lngth - 1) & "\"
            End If
            
            Call DirList(DirString$)
            CommanD = ""
            Exit Sub

            
    ' ---- Directory surfing control ------
        
        Case Else
        
        Dim Cd As String
        Cd = Left$(CommanD, 2)
        If Cd = "cd" Then
            If CommanD = "cd.." Or CommanD = "cd .." Then
                If Cpointer = 0 Then Cpointer = 1
                Cpointer = Cpointer - 1
                Winsock1.SendData vbCrLf & vbCrLf & CPrompt(Cpointer)
                CommanD = ""
                Exit Sub
            End If
            
            If CommanD = "cd\" Or CommanD = "cd \" Then
                Cpointer = 0
                Winsock1.SendData vbCrLf & vbCrLf & CPrompt(Cpointer)
                CommanD = ""
                Exit Sub
            End If
            
            If Len(CommanD) > 3 Then
                Call DirUp(CommanD)
                CommanD = ""
                Exit Sub
            End If
        End If
        
        ' -- If more then just entered was pressed
        If Len(CommanD) > 0 Then
            Winsock1.SendData vbCrLf & "Syntax Error." & vbCrLf
        End If
        
        ' -- else drop a line
        Winsock1.SendData vbCrLf & CPrompt(Cpointer)
        CommanD = ""
        Exit Sub
        
    End Select
         
    End If
    
    
' -- If not vbCrLf then keep filling buffer

    CommanD = CommanD & rfC
     
    ' -- Because Telnet is a {Dumb Terminal}
    ' -- We have to send the key that was
    ' -- pressed back so it shows up on
    ' -- the screen...
    
    Winsock1.SendData rfC
    Exit Sub

' -- End DOS Mode --


'-----------------------[ End DOS Mode ]-------------------------
'////////////////////////////////////////////////////////////////
'// Else we're in standerd Telnet Mode
'//
Else
    
    ' -- was key {ENTER}
    If rfC = vbCrLf Then
        
    ' -- then compare the info in CommanD
    ' -- for a mathing string. The computer
    ' -- looks threw these in order , then
    ' -- when command found then Exit Sub.
    
        Select Case CommanD
            
            Case "help"
                Call sendMenu
                CommanD = ""
                Exit Sub
                
            Case "?"
                Call sendMenu
                CommanD = ""
                Exit Sub
                
            Case "exit"
                lsTpanel.AddItem " -- " & LOGIN_NAME & " Logged Out at " & tme
                Call Button_Click
                CommanD = ""
                Exit Sub
                     
           Case "reboot"
                Winsock1.SendData vbCrLf & "Host now restarting......"
                Call RebootComputer
                CommanD = ""
                Exit Sub
                                
           Case "viewlog"
                Call SendLogInfo
                CommanD = ""
                Exit Sub
                                             
           Case "restart"
                lsTpanel.AddItem " -- Tconsole restart requested by " & LOGIN_NAME & " at " & tme
                Call Button_Click
                CommanD = ""
                Exit Sub
    
           Case "sysinfo"
                lsTpanel.AddItem " -- System Information requested by " & LOGIN_NAME & " at " & tme
                Call SendSystemInfo
                CommanD = ""
                Exit Sub
                
           Case "testprint"
                lsTpanel.AddItem " -- Tconsole restart requested by " & LOGIN_NAME & " at " & tme
                Call PrintTestPage
                CommanD = ""
                Exit Sub
                
           Case "spawn /?"
                Call SpawnMenu
                CommanD = ""
                Exit Sub
                
          Case "spawn"
                Call SpawnMenu
                CommanD = ""
                Exit Sub
                
           Case "get /?"
                Call getMenu
                CommanD = ""
                Exit Sub
                
           Case "get"
                Call getMenu
                CommanD = ""
                Exit Sub
            
            
       ' -- Hook for C: Mode   --
       
           Case "dos"
                
                ' -- this keeps track of the directory level
                Cpointer = 0
                CPrompt(Cpointer) = "C:\>"
                
                lsTpanel.AddItem " -- DOS Mode Started"
                Call DosMenu
                DOS_Mode = True '------ Booleans, so small, yet have so much control
                CommanD = ""
                Exit Sub
            
           Case Else
            
       '-- hook for spawning process --
       
            Dim IsSpawn As String
            
            IsSpawn$ = Left$(CommanD, 5)
            If IsSpawn = "spawn" Then
                Call SpawnFile(CommanD$)
                DoEvents
                
                frmTconsole.SetFocus
                frmTconsole.Winsock1.SendData vbCrLf & prompT
                CommanD = ""
                Exit Sub
            End If
            
            
       '-- hook for the get process --
       
            Dim Getfile As String
            
            Getfile$ = Left$(CommanD, 3)
            If Getfile$ = "get" Then
                Call SendFileInfo(CommanD$)
                DoEvents
                
                frmTconsole.Winsock1.SendData vbCrLf & prompT
                CommanD = ""
                Exit Sub
            End If
            
            
    ' -- Now handle unKnown commands --
        
        '// DOS style error :b
            If Len(CommanD) > 0 Then
                Winsock1.SendData vbCrLf & "Syntax Error. [Check your spelling]" & vbCrLf & "For help type ? or help for a menu." & vbCrLf
            End If
                
        '// Drop a line if when {ENTER} is pressed
            Winsock1.SendData vbCrLf & prompT
            CommanD = ""
            Exit Sub
                
        End Select
    End If
    
    
' -- If not vbCrLf then keep filling buffer

    CommanD = CommanD & rfC
     
    ' -- Because Telnet is a {Dumb Terminal}
    ' -- We have to send the key that was
    ' -- pressed back so it shows up on
    ' -- the screen...
    
    Winsock1.SendData rfC
     
End If

End Sub

Private Sub MarkFailedAtempt()
    
    lsTpanel.AddItem "-Failed Atempt (User: " & logIn$ & " - Pass: " & pasS & ")"

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Winsock1.Close
    Call ResetConsole

End Sub
Public Function CheckIfInfoWasEntered() As Boolean

'// preset return of function
    CheckIfInfoWasEntered = True
    
    If LOGIN_NAME = "" Or LOGIN_PASS = "" Then CheckIfInfoWasEntered = False
    
    ' -- name and pass is realy all we need to know
    ' -- any errors in other information [i.e. port]
    ' -- the computer will correct.
    
End Function
Private Sub kilL()
    
    '-- close sock
    Winsock1.Close
    
    ' -- tell pannel for logging
    lsTpanel.AddItem "- *Services Halted* -"
    
    ' -- Reset button
    ' -- Button_Click() uses "Start" as a pointer
    button.Caption = "Start"


End Sub
