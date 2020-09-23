Attribute VB_Name = "Tconsole"
'-- advapi32 --
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-- kernel32 --
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

'-- user32 --
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
                                                 



Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        wProcessorLevel As Integer
        wProcessorRevision As Integer
End Type

Public Type OSVERSIONINFO ' 148 bytes
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Const VER_INFO_SIZE& = 148
Public Const VER_PLATFORM_WIN32_NT& = 2
Public Const VER_PLATFORM_WIN32_WINDOWS& = 1


Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Global myVer As OSVERSIONINFO


Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

'-------------------------------------------------

Public LogedIn As Boolean
Public HrdKill As Boolean
Public Attempt As Integer
Public logIn As String
Public pasS As String
Public RemoteIP As String
Public rfC As String
Public LocalPrt As Long
Public accept As Boolean
Public nameEntered As Boolean

Public LOG_FILE_NAME As String
Global LOGIN_NAME As String
Global LOGIN_PASS As String
Global BLOCK_ON As Boolean
Global PORT_ON As Boolean
Global LOCAL_PORT As Long
Global MAX_ATTEMPTS As Long
Global DOS_Mode As Boolean
Global Const INIFileName = "Tconsole.ini"

Public ReTurnString As String
'Public CPrompt As String

Public prompT As String
Public CommanD As String
Public FirstTime As Boolean
'-- used for INI write/read
'-- for Converts on Boolean and Long
Global GETPORT As String
Global BLOCK As String
Global PORT As String
'-- End


'//////////////////////////////////////////
'// START: Sub Main()
'/////////////////////////////////////////

Sub Main()
On Error Resume Next
 
        
    '// check to see if program has been used before
    '// If so read previous settings
    
        LOGIN_NAME = GetFromINI("Tconsole", "LoginName", "Tconsole.ini")
        LOGIN_PASS = GetFromINI("Tconsole", "password", "Tconsole.ini")
        BLOCK = GetFromINI("Tconsole", "MonitorOnOff", "Tconsole.ini")
        GETPORT = GetFromINI("Tconsole", "localport", "Tconsole.ini")
        LOG_FILE_NAME = GetFromINI("Tconsole", "LogFileName", "Tconsole.ini")

        LOCAL_PORT = Val(GETPORT)

   '// Set default to block multible faild atempts
    
        If BLOCK = "false" Then
             BLOCK_ON = False
        Else
             BLOCK_ON = True
        End If
    '// if no information was found then
    '// show config window
    
        FirstTime = False
        If LOGIN_NAME = "" Or LOGIN_PASS = "" Or LOG_FILE_NAME = "" Then
            FirstTime = True
            frmConfig.Show
        End If
        
        frmTconsole.Show

End Sub

Public Sub sendMenu()

frmTconsole.Winsock1.SendData vbCrLf & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*                Tconsole Menu                    *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*  1. ? or help  ---------------------- This Menu *" & vbCrLf
frmTconsole.Winsock1.SendData "*  2. viewlog --------------- View Tconsole's Log *" & vbCrLf
frmTconsole.Winsock1.SendData "*  3. sysinfo --------- Get system info on server *" & vbCrLf
frmTconsole.Winsock1.SendData "*  4. restart ------------------ Restart Tconsole *" & vbCrLf
frmTconsole.Winsock1.SendData "*  5. reboot ----------- Restart Tconsole Machine *" & vbCrLf
frmTconsole.Winsock1.SendData "*  6. testprint ----------------- Print test page *" & vbCrLf
frmTconsole.Winsock1.SendData "*  7. spawn /? ---------- For options on spawning *" & vbCrLf
frmTconsole.Winsock1.SendData "*  8. get /? ------------ For info on get command *" & vbCrLf
frmTconsole.Winsock1.SendData "*  9. dos ----------------- to switch to DOS Mode *" & vbCrLf
frmTconsole.Winsock1.SendData "*  10. exit ------------- disconect from Tconsole *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData vbCrLf & prompT

End Sub


Public Function BannerMsg() As String

         intro$ = "----------------------------------------------------" & vbCrLf
intro$ = intro$ & "--          Tconsole Utility Server.              --" & vbCrLf
intro$ = intro$ & "--     All actions on this server are loged       --" & vbCrLf
intro$ = intro$ & "----------------------------------------------------" & vbCrLf & vbCrLf

BannerMsg$ = intro$
End Function

Public Sub getMenu()
frmTconsole.Winsock1.SendData vbCrLf & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*                Get Instructions                 *" & vbCrLf
frmTconsole.Winsock1.SendData "*           Syntax: get FilePath\Filename         *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*        Retrieves contents of an ASCII file      *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "* Exaimple:                                       *" & vbCrLf
frmTconsole.Winsock1.SendData "*      get c:\windows\win.ini                     *" & vbCrLf
frmTconsole.Winsock1.SendData "*      Gets a listing of win.ini's contents       *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "*  [ Read the readme.txt for complete detailz ]   *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData vbCrLf & prompT

End Sub

Public Sub SpawnMenu()
frmTconsole.Winsock1.SendData vbCrLf & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*              Spawning Instructions              *" & vbCrLf
frmTconsole.Winsock1.SendData "*         Syntax: spawn FilePath\Filename         *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "* Exaimple:                                       *" & vbCrLf
frmTconsole.Winsock1.SendData "*      spawn c:\windows\notepad.exe               *" & vbCrLf
frmTconsole.Winsock1.SendData "*      opens notepad                              *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "*  [ Read the readme.txt for complete detailz ]   *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData vbCrLf & prompT
End Sub
Public Sub Savelog(TheList As ListBox)

Dim SaveList As Long
Dim tme As String
    tme = Time
    
    Open LOG_FILE_NAME For Append As #1
            
    Print #1, vbCrLf & " ----- " & tme & " ----- "
        
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    
    Close #1
    
End Sub
Public Sub SendSystemInfo()


Dim flagnum As Long
Dim My As Long, s As String
    
Dim verNum As Long, verWord As Integer
Dim mySys As SYSTEM_INFO

        
        myVer.dwOSVersionInfoSize = VER_INFO_SIZE
        My& = GetVersionEx&(myVer)
        
            If myVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
                s$ = " Windows95/98 "
            ElseIf myVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
                s$ = " Windows NT "
            End If
        
        GetSystemInfo mySys
        
        frmTconsole.Winsock1.SendData vbCrLf & vbCrLf & s$ & myVer.dwMajorVersion & "." & myVer.dwMinorVersion & " Build " & (myVer.dwBuildNumber And &HFFFF&) & vbCrLf & vbCrLf
        frmTconsole.Winsock1.SendData " Page size is " & mySys.dwPageSize & " bytes" & vbCrLf
        frmTconsole.Winsock1.SendData " Lowest memory address: &H" & Hex$(mySys.lpMinimumApplicationAddress) & vbCrLf
        frmTconsole.Winsock1.SendData " Highest memory address: &H" & Hex$(mySys.lpMaximumApplicationAddress) & vbCrLf
        frmTconsole.Winsock1.SendData " Number of processors found: " & mySys.dwNumberOfProcessors & vbCrLf
        frmTconsole.Winsock1.SendData " Processor Type: "
        
        Select Case mySys.dwProcessorType
            Case PROCESSOR_INTEL_386
                    frmTconsole.Winsock1.SendData "Intel 386"
            Case PROCESSOR_INTEL_486
                    frmTconsole.Winsock1.SendData "Intel 486"
            Case PROCESSOR_INTEL_PENTIUM
                    frmTconsole.Winsock1.SendData "Intel Pentium"
            Case PROCESSOR_MIPS_R4000
                    frmTconsole.Winsock1.SendData "MIPS R4000"
            Case PROCESSOR_ALPHA_21064
                    frmTconsole.Winsock1.SendData "Alpha 21064"
        End Select
        
        frmTconsole.Winsock1.SendData vbCrLf & vbCrLf & prompT
        
End Sub


'- this is needed for earlier versions of Vbasic
'Attribute VB_Name = "Shell_Execute"


Public Sub SpawnFile(CommD As String)

Dim StrLen As Integer
Dim FilenameForSpawn As String

StrLen = Len(CommD$)

         If CommD = "spawn /?" Then
                Call SpawnMenu
                CommanD = ""
                Exit Sub
        End If



'/////////////////////////////////////////////
'// this pulls the path/filename from spawned command
'// then seperates the actual programed to be spawned
'// from the directory path. wich is needed for the shell
'// command.
'///
On Error Resume Next

Dim sTrng As String, sTemp1 As String
Dim sTemp As String, File As String
Dim filePath As String


sTrng$ = CommD '-------- save the command
lenstr = Len(sTrng) ' -- Get length of total command

' -- Now we find the first space 'spawn c:\*.*'
sTemp1$ = Mid(sTrng, InStr(1, sTrng, " "), lenstr)

' -- now we move to secondary buffer for the work.
sTemp$ = sTemp1$

iStart = 1

    Do
        lenstr = Len(sTemp)
        iPos = InStr(iStart, sTemp, "\")
        
        
        If iPos <> 0 Then
            sTemp$ = Mid(sTemp, iPos + 1, lenstr)
            stmp$ = Mid(sTemp, InStr(1, sTemp, "\"), lenstr)
            iStart = iPos + 1
        End If

    Loop Until iPos = 0
       
       
' -- C:\path\filename.ext

stmp$ = Mid(stmp, 2, lenstr)
File$ = stmp$
'FileName = filename.ext

lncommand = Len(stmp)
filePath$ = Mid(sTemp1, 1, Len(sTemp1) - lncommand)
'FilePath = C:\path\
 
       Dim temp, Msg As String
       Dim X
       temp = GetActiveWindow()
       
       X = ShellExecute(temp, "Open", File$, "", filePath$, 1)
       
        If X = 2 Then
            frmTconsole.Winsock1.SendData vbCrLf & "Error spawning proccess" & vbCrLf & " 'File note found' check your spelling ;)" & vbCrLf
        ElseIf X > 32 Then
            Dim Tellconsole As String, tme As String
            tme = Time
            Tellconsole$ = " -- " & filePath$ & File$ & " was spawned at " & tme
            frmTconsole.lsTpanel.AddItem Tellconsole$
            frmTconsole.Winsock1.SendData vbCrLf & "Proccess has been spawned" & vbCrLf
        End If
        
   End Sub



Public Sub SendFileInfo(ByVal CommD As String)
       
On Error GoTo Over

Dim sTrng As String, sTemp1 As String
Dim sTemp As String, File As String
Dim filePath As String

sTrng$ = CommD
lenstr = Len(sTrng)
sTemp1$ = Mid(sTrng, InStr(1, sTrng, " "), lenstr)
strfilePath$ = Trim(sTemp1$)
'// strFilePath$ now equals C:\path\filename.ext



       Dim i As Integer
       Dim X As Integer
       X = FreeFile
       Open strfilePath$ For Input As #X


       Do While Not EOF(X)
           Line Input #X, sData
       frmTconsole.Winsock1.SendData vbCrLf & sData
  
       Loop


       Close #X
       frmTconsole.Winsock1.SendData vbCrLf
       Tellconsole$ = " -- " & strfilePath$ & " was opened for read at " & tme
       
Over:
End Sub


Public Sub PrintTestPage()

Shell "rundll32 msprint2.dll,RUNDLL_PrintTestPage"
 frmTconsole.Winsock1.SendData vbCrLf & "test page printed"
    frmTconsole.Winsock1.SendData vbCrLf & vbCrLf & prompT
    
End Sub
Public Sub SendLogInfo()


Dim a As Long
Dim SendLog As String

    frmTconsole.Winsock1.SendData vbCrLf

    For a& = 0 To frmTconsole.lsTpanel.ListCount - 1
        SendLog = frmTconsole.lsTpanel.List(a&)
        frmTconsole.Winsock1.SendData SendLog & vbCrLf
    Next a
    
    frmTconsole.Winsock1.SendData vbCrLf & prompT


End Sub

