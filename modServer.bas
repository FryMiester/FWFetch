Attribute VB_Name = "modServer"
Option Explicit


Public Sub doConnectString()
    On Error GoTo doConnectString_Error

    With frmGetMain
        .cmdConnect.BackColor = vbYellow
        .cmdConnect.Caption = "CONNECTING"
        sendToIRC "USER " & MyNick & " " & MyNick & " __A__ " & MyNick & " " & "FalconWorks Software" & vbCrLf & "NICK " & MyNick
        Pause 5
        If MyNickServPass > "" Then
            sendToIRC "PRIVMSG " & "Nickserv Identify " & MyNickServPass
        End If
        Pause 5
        sendToIRC "JOIN " & NulTrim(MyChannel)
    End With
    SysLog "Connected to " & NulTrim(MyChannel) & " on " & NulTrim(MyNetwork), vbGreen
    AmConnected = True
    On Error GoTo 0
    Exit Sub

doConnectString_Error:

'    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure doConnectString of Module modServer"
End Sub

Public Sub sendToIRC(ByVal strData As String)
    On Error GoTo sendToIRC_Error

    If frmGetMain.sckServer.State = sckConnected Then
        frmGetMain.sckServer.SendData strData & vbCrLf
'        If DeBugger = True Then
            Debug.Print "SENDING.." & strData
'        End If
    End If

    On Error GoTo 0
    Exit Sub

sendToIRC_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure sendToIRC of Module modServer"
End Sub

Public Sub IRCNotice(NickName As String, strInfo As String)
sendToIRC "NOTICE " & NickName & " " & strInfo
End Sub

Public Sub IRCMsg(NickName As String, strInfo As String)
sendToIRC "PRIVMSG " & NickName & " " & strInfo
End Sub

Public Sub IRCSay(Channel As String, strInfo As String)
sendToIRC "PRIVMSG " & Channel & " " & strInfo
PaintScreen Format(Now, "HH:MM") & " < " & MyNick & " > " & strInfo, vbGreen
End Sub

Public Sub IRC_CTCP_Notice(NickName As String, strInfo As String)
sendToIRC "NOTICE " & NickName & " : " & strInfo & " "
End Sub

Public Sub IRC_CTCP_MSG(NickName As String, strInfo As String)
sendToIRC "PRIVMSG " & NickName & " : " & strInfo & " "
End Sub

Public Sub SysLog(TextString As String, Optional txtcolor As Long)
Dim OutString As String
Dim LogFile As String
Dim Handle As Integer


Dim DatePortion As String
   On Error GoTo SysLog_Error

DatePortion = Format(Now, "DD/MM  HH:MM:SS")
OutString = "< " & DatePortion & " > " & TextString & vbNewLine

If txtcolor = 0 Then txtcolor = vbBlack
frmGetMain.sysInfo.SelStart = Len(frmGetMain.sysInfo.Text)
frmGetMain.sysInfo.SelColor = txtcolor
frmGetMain.sysInfo.SelText = frmGetMain.sysInfo.SelText & OutString

OutString = Replace(OutString, vbCrLf, "")

'LogFile = App.Path & "\SystemLog.txt"
'Handle = FreeFile
'Open LogFile For Append As Handle
'Print #Handle, OutString
'Close Handle

   On Error GoTo 0
   Exit Sub

SysLog_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure SysLog of Module modMisc"
End Sub

Public Function WhoisOnLine(InputStr As String) As Boolean
Dim NickName As String

If InStr(InputStr, " 311 ") <> 0 Then
    WhoisOnLine = True
Else
    WhoisOnLine = False
End If

End Function

Public Sub CmndProcess(ByVal StrInPut As String)
Dim Cmnd As String
Dim strOption As String
Dim tmpText
Dim Modestring As String

Cmnd = UCase(nWord(StrInPut, " "))
strOption = StrInPut

Select Case Cmnd
      Case "/ME"
            tmpText = "ACTION " & strOption & ""
            sendToIRC "PRIVMSG" & " " & MyChannel & " " & tmpText
            tmpText = " ** " & strOption & " **"
            Modestring = Format(Now, "HH:MM") & " " & MyNick & "  " & tmpText & vbCrLf
            PaintScreen Modestring, vbGreen
      Case "/MSG"
      Case "/NOTICE"
      
      
      Case Else
End Select


End Sub

