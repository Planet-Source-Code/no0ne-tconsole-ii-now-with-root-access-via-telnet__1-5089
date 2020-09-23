Attribute VB_Name = "DOSMode"
'//////////////////////////////////////////
'// After given this some thought. I didn't want to
'// re-create DOS, BUT the Prompt that DOS does use
'// is it's directory structure. So I stuck to it
'// in the name of not passing a bunch a varibles
'// around trying to turn DOS in to something else.
'//
'// SO, here it is . DOS Via Telnet
'///

Global CPrompt(32) As String ' -- Directories deep
Global Cpointer As Integer ' ---- pointer to each array space

Public Sub DirList(DirLst As String)

Dim FileNumber As Long
Dim FileName As String

On Error Resume Next

frmTconsole.Winsock1.SendData vbCrLf

FileName$ = Dir(DirLst, vbDirectory) ' ----- a pointer to the file names in the directory
FileNumber = GetAttr(DirLst & FileName) ' -- the number of the position of each file

    Do While FileName$ <> "" ' ------------- go through all the files
        
        '-- now run through the list and spit out each name in the dir acording to its number
        If FileNumber = vbDirectory Then
            frmTconsole.Winsock1.SendData FileName & vbCrLf
        End If
        
        ' -- Dir is realy a pointer to a system array for holding all the names in a directory.
        FileName = Dir
    Loop
    
    frmTconsole.Winsock1.SendData vbCrLf & CPrompt(Cpointer)
End Sub

Public Sub DirUp(cmmD As String)
Dim Lngth As Long, temp2 As String
Dim DirFolder As String, working As String
Dim sTrng As String, sTemp1 As String
Dim strfilePath As String, temp3 As String

    ' -- pic the prompt apart ( for our directory struckure )
    Lngth = Len(CPrompt(Cpointer))
    temp2 = Left(CPrompt(Cpointer), Lngth - 1)
    
    ' - now pull the directory to change to, from the command
    sTrng$ = cmmD
    Lngth& = Len(sTrng)
    sTemp1$ = Mid(sTrng$, InStr(1, sTrng$, " "), Lngth&)
    strfilePath$ = Trim(sTemp1$)
    ' -- strFilePath now is the directory we are jumping up to

        
        '-- now check for If folder exist --
        If Cpointer <> 0 Then ' --- are we at c:\> or c:\windows>
            temp3 = temp2 & "\" & strfilePath$ & "\"
        Else
            temp3 = temp2 & strfilePath$ & "\"
        End If
        
        ' -- Is it a valid directory
        DirFolder = Dir(temp3$, vbDirectory)
    
        If DirFolder <> "" Then ' -- if valid folder then
        
            '-- change the directory pointer up one level
            Cpointer = Cpointer + 1
            If Cpointer > 1 Then
                CPrompt(Cpointer) = temp2 & "\" & strfilePath$ & ">"
            Else
                CPrompt(Cpointer) = temp2 & strfilePath$ & ">"
            End If
            
            ' -- CPrompt() holds the key for our directory structure
            frmTconsole.Winsock1.SendData vbCrLf & CPrompt(Cpointer)


        Else ' -- it is an invalid directory
        
            frmTconsole.Winsock1.SendData vbCrLf & "Invalid Directory." & vbCrLf
            frmTconsole.Winsock1.SendData vbCrLf & CPrompt(Cpointer)
            Exit Sub
       
        End If
        
End Sub
Public Sub DosMenu()
frmTconsole.Winsock1.SendData vbCrLf & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "*          DOS Mode Listing of commands           *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData "* 1. dir or ls ----- listing of files and folders *" & vbCrLf
frmTconsole.Winsock1.SendData "*                    in the current directory.    *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "* 2. cd ------------ Change directory.            *" & vbCrLf
frmTconsole.Winsock1.SendData "*        a. C:\>cd windows = C:\windows>          *" & vbCrLf
frmTconsole.Winsock1.SendData "*        b. C:\windows>cd.. = C:\>                *" & vbCrLf
frmTconsole.Winsock1.SendData "*        c. C:\windows\system>cd\ = C:\>          *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "* 3. ? or help ----- For this menu.               *" & vbCrLf
frmTconsole.Winsock1.SendData "*                                                 *" & vbCrLf
frmTconsole.Winsock1.SendData "* 4. exit ---------- Quit DOS Mode                *" & vbCrLf
frmTconsole.Winsock1.SendData "***************************************************" & vbCrLf
frmTconsole.Winsock1.SendData vbCrLf & CPrompt(Cpointer)

End Sub
