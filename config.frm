VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H80000004&
   Caption         =   "Tconsole Configuration Manager"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "config.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Log File Setup"
      Height          =   975
      Left            =   1560
      TabIndex        =   15
      Top             =   2760
      Width           =   4695
      Begin VB.TextBox txTLogFileName 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Text            =   "log.txt"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "( Saved to the programs current Directory )"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label label4 
         Caption         =   "Log  File Name"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   0
      Picture         =   "config.frx":08CA
      ScaleHeight     =   4635
      ScaleWidth      =   1395
      TabIndex        =   14
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Apply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login Setup   (Case Sensitive)"
      Height          =   2655
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txTport 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Text            =   "10101"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chKPort 
         Caption         =   "Use Non-standard Telnet Port"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txTAttempts 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Text            =   "4"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chKmonitor 
         Caption         =   "Monitor Failed Login attempts"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txTpass 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "linux"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txTname 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "root"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Use Port:"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Disconect after"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Login Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Apply_Click()

'// set varibles
LOGIN_NAME = txTname.Text
LOGIN_PASS = txTpass.Text

'// Check 4 user erorrs
If LOGIN_NAME = "" And LOGIN_PASS = "" Then
    MsgBox "Name and Login have to be suplied.", vbInformation, "Configure error"
    txTname.SetFocus
    Exit Sub
ElseIf LOGIN_NAME = "" Then
    MsgBox "You must supply a user name!", vbInformation, "Username error"
    txTname.SetFocus
    Exit Sub
ElseIf LOGIN_PASS = "" Then
    MsgBox "You must supply a password!", vbInformation, "Password error"
    txTpass.SetFocus
    Exit Sub
End If
'-- End basic input error check


'// Disconect after number of failed attempts
If chKmonitor.Value = 1 Then
    BLOCK_ON = True
    If txTAttempts.Text = "" Then
        MAX_ATTEMPTS = 4
    Else
        MAX_ATTEMPTS = txTAttempts.Text
    End If
Else
    BLOCK_ON = False
End If

'// check for entered port number
If chKPort.Value = 1 Then
    
    '// is port number to high or below 0
    If Val(txTport.Text) > 64500 Or Val(txTport.Text) <= 0 Then
        MsgBox "Invalid Port number.", vbInformation, "Port error"
        txTport.SetFocus
        Exit Sub
    End If
    
    '// is port number even entered
    If txTport.Text = "" Then
        LOCAL_PORT = 23
        txTport.Text = "23"
    Else
        LOCAL_PORT = txTport.Text
    End If
    PORT_ON = True
Else
    LOCAL_PORT = 23
    PORT_ON = False
End If


If txTLogFileName.Text = "" Then
    LOG_FILE_NAME = "log.txt"
    txTLogFileName.Text = "log.txt"
Else
    LOG_FILE_NAME = txTLogFileName.Text
End If

    

'////////////////////////////
'// save in INI file for later checking
'// Program loads settings from .ini on start up
'// if varibles not set then show config screen
'///
Dim HeadeR As String

HeadeR = "Tconsole"

Call WriteToINI(HeadeR$, "LoginName", LOGIN_NAME, INIFileName$)
Call WriteToINI(HeadeR$, "PassWord", LOGIN_PASS, INIFileName$)
Call WriteToINI(HeadeR$, "MonitorOnOff", BLOCK_ON, INIFileName$)
Call WriteToINI(HeadeR$, "portOnOff", PORT_ON, INIFileName$)
Call WriteToINI(HeadeR$, "LocalPort", LOCAL_PORT, INIFileName$)
Call WriteToINI(HeadeR$, "LogFileName", LOG_FILE_NAME, INIFileName$)

If FirstTime = True Then Call frmTconsole.Button_Click

End Sub


Private Sub chKmonitor_Click()
If chKmonitor.Value <> 1 Then
    txTAttempts.BackColor = &H80000004
    txTAttempts.Enabled = False
Else
   txTAttempts.BackColor = &H80000005
   txTAttempts.Enabled = True
End If

End Sub

Private Sub chKPort_Click()
If chKPort.Value <> 1 Then
    txTport.BackColor = &H80000004
    txTport.Enabled = False
Else
   txTport.BackColor = &H80000005
   txTport.Enabled = True
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command4_Click()

Call Apply_Click
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next

If FirstTime = True Then GoTo Skip

   
        txTname = GetFromINI("Tconsole", "LoginName", "Tconsole.ini")
        txTpass.Text = GetFromINI("Tconsole", "password", "Tconsole.ini")
        BLOCK = GetFromINI("Tconsole", "MonitorOnOff", "Tconsole.ini")
        PORT = GetFromINI("Tconsole", "portOnOff", "Tconsole.ini")
                
        txTport.Text = GetFromINI("Tconsole", "localport", "Tconsole.ini")
        txTLogFileName.Text = GetFromINI("Tconsole", "LogFileName", "Tconsole.ini")
        
        If BLOCK Then
            'chKmonitor.Value = 1
            Call chKmonitor_Click
        End If
        
        If PORT Then
            'chKPort.Value = 1
            Call chKPort_Click
        End If
  
Skip:

StayOnTop Me


End Sub
