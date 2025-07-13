Attribute VB_Name = "modPanel_0"
Option Explicit
Public Type UserPackage
    NickName As String
    RealName As String
    LoggedAs As String
    IP_Info As String
    IsRegUser As Boolean
    ChanList As String
    IsElevated As Integer
    Registered As Boolean
    TransferLimit As Boolean
    FileName As String
    FileSize As Long
    VersionInfo As String
End Type

Global LastScanTime As Long

Private Const Panel = 0

Public Sub DoBuff()
    Dim SVR As ListServerInfo
    Dim InputString As String
    Dim ModeString As String
    Dim tTempstr As String
    Dim sTempStr As String
    Dim ThisNick As String
    Dim sTrigger As String
    Dim Success As Long
    Dim TestChr As Integer
    Dim Garbage As String
    Dim MyTrigger As String
    Dim LogStr As String
    Dim StatString As String
    Dim IPAddr As String
    Dim RecordNum As Long
    Dim HasSlots As Boolean
    Dim TempInfo As String
    Dim a As Long, b As Long
    Dim tTime As Date
    Dim Cnt As Integer
    Dim UserName As String
    Dim L As Long
    Dim cColor As Long
    Dim UserInfo As UserPackage
    Dim SearchStr As String
    
    Dim MD5 As CMD5
    Set MD5 = New CMD5
    
    Dim tmpListServer As clsListServer
    Set tmpListServer = New clsListServer
    
    On Error GoTo DoBuff_Error
    
    ININame = MyServersFolder & "Servers.ini"
    
    If frmGetMain.lstDialogBuff.ListCount = 0 Then Exit Sub
    
    InputString = frmGetMain.lstDialogBuff.List(0)
    
    StatString = InputString
    sTempStr = InputString
    TempInfo = InputString
    ThisNick = nWord(TempInfo, "!")
    Garbage = nWord(TempInfo, "@")
    IPAddr = nWord(TempInfo, " ")
    Garbage = nWord(TempInfo, ":")
        
    If InStr(sTempStr, " 311 ") <> 0 Or InStr(sTempStr, " 401 ") <> 0 Then
        RawWHO UserInfo, sTempStr
    End If
    
    
    
    frmGetMain.lstDialogBuff.RemoveItem (0)
    If NulTrim(InputString) = "" Then Exit Sub
    
    If InStr(InputString, "NOTICE") <> 0 And InStr(InputString, MyNick) <> 0 Then
        If InStr(UCase(InputString), "DCC SEND") <> 0 Then Exit Sub
        TempInfo = InputString
        TempInfo = Right(TempInfo, Len(TempInfo) - 1)
        ThisNick = nWord(TempInfo, "!")
        Garbage = nWord(TempInfo, "@")
        IPAddr = nWord(TempInfo, " ")
        Garbage = nWord(TempInfo, ":")
        TempInfo = StripColor(TempInfo)
        
        ModeString = Format(Now, "HH:MM") & " < " & ThisNick & " NOTICE > " & TempInfo
        PaintScreen ModeString, vbRed
        Exit Sub
    End If
    
    If InStr(InputString, " PRIVMSG ") <> 0 And InStr(InputString, MyNick) <> 0 Then
        TempInfo = InputString
        TempInfo = Right(TempInfo, Len(TempInfo) - 1)
        ThisNick = nWord(TempInfo, "!")
        If frmGetMain.FramePrivateChat.Caption = "" Then
            For L = 0 To 5
                frmGetMain.cmdFrame(L).BackColor = &H8000000F
                frmGetMain.FrameMain(L).Visible = False
            Next
            frmGetMain.cmdFrame(5).BackColor = &H80000016
            frmGetMain.FrameMain(5).Visible = True
        End If
        frmGetMain.FramePrivateChat.Caption = ThisNick
        Garbage = nWord(TempInfo, "@")
        IPAddr = nWord(TempInfo, " ")
        Garbage = nWord(TempInfo, ":")
        ModeString = Format(Now, "HH:MM") & " <" & ThisNick & "> " & TempInfo & vbCrLf
        ColorText ModeString, vbBlue
        Exit Sub
    End If
    
    
    ModeString = InputString

    ' Get the NickName of the poster
    
    InputString = Right(InputString, Len(InputString) - 1)
    ThisNick = nWord(InputString, "!")



    InputString = NulTrim(InputString)
    If InStr(InputString, "SLOTS ") <> 0 And InStr(InputString, " 999 ") <> 0 Then
      If ServersPaused = True Then Exit Sub
      If AmPaused = True Then Exit Sub
            SysLog "Slots Msg from: " & ThisNick
            SearchStr = MD5.HaxMD5(ThisNick)
            ProcessSlotMsg ModeString

            Exit Sub
    End If
   
   If InStr(UCase(InputString), "TYPE") <> 0 Then
        If InStr(UCase(InputString), "@" & UCase(ThisNick)) <> 0 Then
            If AmPaused = True Then Exit Sub
            If ServersPaused = True Then Exit Sub
            'SysLog "Non-Slots Msg from " & ThisNick
            SearchStr = MD5.HaxMD5(ThisNick)
            RecordNum = GetServerByHash(SearchStr)
            If RecordNum = -1 Then
               ProcessAdvert ModeString
               Exit Sub
            Else
               If LServers(RecordNum).HasSlots = True Then
                   SysLog "Non-Slots Msg from " & ThisNick & " - Ignored"
                   Exit Sub
               End If
            
            End If
            Exit Sub
        End If
    End If
    
    'DebugLog StatString
    ModeString = ""
    ModeString = Format(Now, "HH:MM") & " < "
        TempInfo = StatString
        TempInfo = StripColor(TempInfo)
        Garbage = nWord(TempInfo, ":")
        ThisNick = nWord(TempInfo, "!")
        Garbage = nWord(TempInfo, ":")
        If frmGetMain.chkFilterRequests.Value = 1 Then
            If Left(TempInfo, 1) = "!" Then
                ModeString = ""
                GoTo Bottom
            Else
                ModeString = Format(Now, "HH:MM") & " < " & ThisNick & " > " & TempInfo & vbCrLf
            End If
        Else
            ModeString = Format(Now, "HH:MM") & " < " & ThisNick & " > " & TempInfo & vbCrLf
        End If
        
        If frmGetMain.chkFilterSearches.Value = 1 Then
            If Left(TempInfo, 1) = "@" Then
                ModeString = ""
                GoTo Bottom
            Else
                ModeString = Format(Now, "HH:MM") & " < " & ThisNick & " > " & TempInfo & vbCrLf
            End If
        Else
            ModeString = Format(Now, "HH:MM") & " < " & ThisNick & " > " & TempInfo & vbCrLf
        End If
        
' ================ Check Ignore List ===================
' Username = name from ignore list
Cnt = frmGetMain.lstIgnores.ListCount
If Cnt = 0 Then GoTo Bottom
For L = 0 To Cnt - 1
    UserName = UCase(frmGetMain.lstIgnores.List(L))
    If InStr(UCase(ModeString), UserName) <> 0 Then '          UserName = UCase(ThisNick) Then
        ModeString = ""
        GoTo Bottom
    End If
Next
        
Bottom:
                
    If ModeString > "" Then
         cColor = vbRed
         If Left(TempInfo, 1) = "!" Then cColor = vbBlue
         If Left(TempInfo, 1) = "@" Then cColor = vbGreen
        
        PaintScreen ModeString, cColor
    End If
    On Error GoTo 0
    Exit Sub

DoBuff_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure DoBuff of Module modPanel_0"

End Sub

Public Sub RawWHO(UserInfo As UserPackage, InString As String)


End Sub



Public Sub RAWWhoIs(InString As String)
Dim WhoInfo() As String
Dim Channels As String
Dim ThisChan As String
Dim Garbage As String
Dim Params() As String
Dim tUser As String
Dim Code As String
Dim Compare As String
Dim Elevated As Integer
Dim L As Long
Dim Z As Long
Dim j As Long
Dim OutText As String
Dim cCount As Integer
Dim Userpkg As UserPackage
Dim Elevation As String
Dim ThisNick As String

Static Buffer As String

   On Error GoTo RAWWhoIs_Error
'DebugLog InString
WhoInfo = Split(InString, vbCrLf)
cCount = UBound(WhoInfo)
If cCount = 0 Then Exit Sub

ReDim Params(cCount)

For L = 0 To cCount
    Garbage = nWord(WhoInfo(L), " ")
    Code = nWord(WhoInfo(L), " ")
    Garbage = nWord(WhoInfo(L), " ")
    tUser = nWord(WhoInfo(L), " ")
    
    Select Case Code
        Case "352"
        ':zathras.mo.us.irchighway.net 352 AstroDawg #ebooks boomboxnati ihw-p0q.uv7.116.199.IP *.irchighway.net boomboxnation H :0 boomboxnation
            Userpkg.NickName = nWord(WhoInfo(L), " ")
            Userpkg.IP_Info = nWord(WhoInfo(L), " ")
            Garbage = nWord(WhoInfo(L), " ")
            Garbage = nWord(WhoInfo(L), " ")
            Elevation = nWord(WhoInfo(L), " ")
            If NulTrim(Elevation) = "H" Then
                  Elevation = ""
                  ThisNick = NulTrim(Userpkg.NickName)
            Else
                  ThisNick = Left(Elevation, 1) & NulTrim(Userpkg.NickName)
             End If
             DoUserList ThisNick
             
            Userpkg.RealName = NulTrim(WhoInfo(L))
            Garbage = nWord(WhoInfo(L), ":")
            Params(0) = Userpkg.NickName & "@" & Userpkg.IP_Info
            Params(1) = Userpkg.RealName
            If InStr(Userpkg.RealName, "Falcon") <> 0 Then
                DebugLog Userpkg.NickName & "  " & Userpkg.RealName
            End If
            
        Case "319"
            WhoInfo(L) = Right(WhoInfo(L), Len(WhoInfo(L)) - 1)
            Params(2) = WhoInfo(L)
            Channels = UCase(Params(2))
            Userpkg.ChanList = Params(2)
        Case "312"
            Params(3) = WhoInfo(L)
            Debug.Print Params(3)
        Case "330"
            Params(4) = "Is Logged In As: " & nWord(WhoInfo(L), " ")
            Userpkg.LoggedAs = Params(4)
        Case "307"
            Userpkg.Registered = True
            Params(5) = "Is a Registered User"
        Case "671"
            Params(6) = "Is using a secure connection"
        Case "276"
            For j = 1 To 4
                Garbage = nWord(WhoInfo(L), " ")
            Next
            Params(6) = "Client Fingerprint: " & WhoInfo(L)
        Case "318"
    End Select
    
    
    If Params(L) > "" Then
        OutText = OutText & tUser & " : " & Params(L) & vbCrLf
    End If

Next


   On Error GoTo 0
   Exit Sub

RAWWhoIs_Error:

    'DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure RAWWhoIs of Module modDoBuffer"
End Sub

Public Sub DoUserList(UserName As String)
Dim Cnt As Integer
Dim L As Integer
Dim Sample As String


Dim Matched As Boolean
With frmGetMain.lstUsers
      Cnt = .ListCount
      For L = 0 To Cnt - 1
            Sample = .List(L)
            If Sample = UserName Then
                  Matched = True
                  Exit For
            End If
      Next
      If Matched = False Then
            .AddItem UserName
      End If
       frmGetMain.lblUserCnt.Caption = CStr(.ListCount)
End With
End Sub

Public Sub DoUserQuits(UserName As String)
Dim Cnt As Integer
Dim L As Integer
Dim Sample As String

Dim Matched As Boolean
With frmGetMain.lstUsers
      Cnt = .ListCount
      For L = 0 To Cnt - 1
            Sample = .List(L)
            If Sample = UserName Then
                  Matched = True
                  .RemoveItem (L)
                  Exit For
            End If
      Next
      frmGetMain.lblUserCnt.Caption = CStr(.ListCount)
End With
End Sub

