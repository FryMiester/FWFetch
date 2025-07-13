Attribute VB_Name = "modDefines"
Option Explicit


Type myFileBlock
    FileName      As String * 255
    OutPath       As String * 125
    TriggerStr    As String * 255
    ServerName    As String * 50
    HashString    As String * 12
    FileSize      As Long
    Position      As Long
    fHandle       As Integer
    IPAddr        As String * 55
    PortNum       As String
    TotalBytes    As Long
    AmActive      As Boolean
End Type

Type SearchHeader
    NickName As String
    ListFolder As String
    OutputName As String
    ResultFolder As String
    MaxResults As Integer
    FindCount As Integer
    SearchTerms As String
    Quoted As Boolean
    Completed As Boolean
End Type


Public Type IRCChannel
    ChannelName         As String * 50
    NetworkName         As String * 75
    NetworkAddress      As String * 75
    NetworkPort         As String * 10
    BotNickName         As String * 45
    NickServPass        As String * 45
    PortRangeStart      As String * 6
    PortRangeEnd        As String * 6
    ListPath            As String * 255
    AutoConnect         As Boolean
    DLFolder            As String * 255
    ChkMark(10)         As Boolean
    IsScheduled         As String * 1
    SchedTime           As String * 5
    SchedStop           As String * 5
    UpDateDays          As String * 2
    SendToTray          As String * 1
    SpareStringSpace    As String * 461
End Type

Public Type IRCNetWork
    ChannelInfo As IRCChannel
End Type

Type ServerInfo
    NickName        As String * 75
    isDisabled      As Boolean
    ListNeeded      As Boolean
    ISOnline        As Boolean
    Deleted         As Boolean
End Type

'Type ListServerInfo
'    IPAddress       As String * 50
'    SpaceSave       As String * 21
'    ListDate        As Long
'    NextUpDate      As Long
'    NickName        As String * 46
'    FailCount       As Integer
'    HasSlots        As Boolean
'    FilesSize       As Double
'    FilesCount      As Long
'    isDisabled      As Boolean
'    ListNeeded      As Boolean
'    ISOnline        As Boolean
'    Deleted         As Boolean
'    RecordNum       As Long
'    Storage         As String * 150
'End Type


Type SearchBlock
    Index           As Integer
    ListFolder      As String * 255
    ResultFolder    As String * 255
    ZipResults      As Boolean
    NickName        As String * 45
    SearchTerms     As String * 150
    Quoted          As Boolean
    MaxResults      As Integer
End Type

Type FileListBlk
    Connected       As Boolean
    Index           As Integer
    NickName        As String
    RecordNum       As Long
    Outfile         As String
    FileSize        As Long
    Handle          As Integer
    TotalBytes      As Long
    Failed          As Boolean
    Active          As Boolean
End Type

Global ReqString As New Collection
Global SearchBots As New Collection
Global SearchBotCnt As Integer
Global Tree_View As Boolean
Global SVRPause As Boolean
Global GetString(2) As String
Global DLFileBlock(4) As myFileBlock
Global Const Grayed = &H8000000F
Global IRCConfig As IRCNetWork
Global MyVersion As String
Global FileListBlock As FileListBlk
Global MyNick As String
Global MyNetwork As String
Global MyNetworkAddr As String
Global MyNetworkPort As String
Global MyChannel As String
Global MyNickServPass As String
Global MyListFolder As String
Global MyLogFolder As String
Global MyAdditionalFolder As String
Global MySearchResultFolder As String
Global MyServersFolder As String
Global MyIPAddress As String
Global ServerFile As String
Global MyPathHeader As String
Global MyDownloadFolder As String
Global MyProcessFolder As String
Global MyInComingFolder As String
Global Cmnd As String
Global GoodCount As Integer
Global BadCount As Integer
Global Garbage As String
Global ProcessFile(2) As String
Global Screenit(2) As Integer
Global AmConnected As Boolean
Global PathHeader As String
Global SVRRecNum As String
Global ININame As String
Global GetListSaveFile  As String
Global TransferLogFile As String
Global AmPaused As Boolean
Global ServersPaused As Boolean
Global dccTracker(4) As String
Global NewConfig As Boolean
Global INSchedule As Boolean
Global UseSchedule As Boolean
Global LastServer As String
Public ListServer As ListServerInfo
Global SearchAbort As Boolean
Global UpDateDays As Integer


Public Sub ReadConfigs()
Dim ControlFile As String
Dim ControlHandle As Integer
ControlFile = App.Path & "\FWFetch.dat"
ControlHandle = FreeFile
Open ControlFile For Random As ControlHandle Len = Len(IRCConfig)
If LOF(ControlHandle) > 1 Then
    Get #ControlHandle, 1, IRCConfig
End If
Close ControlHandle
ReDim LServers(0)
End Sub

Public Sub WriteConfigs()
Dim ControlFile As String
Dim ControlHandle As Integer
ControlFile = App.Path & "\FWFetch.dat"
ControlHandle = FreeFile
Open ControlFile For Random As ControlHandle Len = Len(IRCConfig)
Put #ControlHandle, 1, IRCConfig
Close ControlHandle
NewConfig = True

End Sub



Public Sub ServerPoll()
Dim SVR As ListServerInfo
Dim ServerCnt As Integer
Dim L As Long
Dim Z As Long
Dim TMPStr As String
Dim Matched As Boolean
Dim IsThere As Boolean
Dim OnLine As Boolean
Dim DisAbled As Boolean
Dim IPAddr As String
Dim RecordNum As Long
Dim Header As String
Dim SvrNum As String
Dim NickName As String
Dim ListCnt As Integer
Dim ServerFile As String
Dim UpDate As Boolean

frmGetMain.LVServers.ListItems.Clear
frmGetMain.LVServers.View = lvwReport

ListCnt = frmGetMain.File1.ListCount

ServerFile = MyServersFolder & "Servers.ini"

ServerCnt = CInt(Val(ReadINI("SERVERS", "SVRCNT", ServerFile))) 'UBound(LServers)
If ServerCnt = 0 Then Exit Sub
ReDim LServers(ServerCnt)
For L = 1 To ServerCnt
   SvrNum = "SVR_" & Format(L, "00#")
   Header = ReadINI("SERVERS", SvrNum, ServerFile)
   With LServers(L)
      .CurrListDate = CLng(Val(ReadINI(Header, "CURRLISTDATE", ServerFile)))
      .FailCount = CInt(Val(ReadINI(Header, "FAILCOUNT", ServerFile)))
      .FilesCount = CDbl(Val(ReadINI(Header, "FILESCOUNT", ServerFile)))
      .FilesSize = CDbl(Val(ReadINI(Header, "FILEBYTES", ServerFile)))
      .HasSlots = Word2Bool(ReadINI(Header, "HASSLOTS", ServerFile))
      .IPAddress = ReadINI(Header, "SVRIPADDR", ServerFile)
      .isDisabled = Word2Bool(ReadINI(Header, "DISABLED", ServerFile))
      .ServerName = ReadINI(Header, "SVRNAME", ServerFile)
      NickName = NulTrim(.ServerName)
      .ISOnline = ISOnline(frmGetMain.lstUsers, NickName)
      .NextUpDate = CLng(Val(ReadINI(Header, "NEXTPOLLDATE", ServerFile)))
      If .NextUpDate < 1000 Then
         .NextUpDate = Date2Epoch(Now) - 31536000
      End If
      If Date2Epoch(Now) > .NextUpDate Then
         If .isDisabled = False Then
            .ListNeeded = True
            UpDate = True
         End If
      Else
         If .isDisabled = False Then
            .ListNeeded = Word2Bool(ReadINI(Header, "LISTNEEDED", ServerFile))
         End If
      End If
      .SVRHash = Header
      .Trigger = ReadINI(Header, "TRIGGER", ServerFile)
   End With
   If UpDate = True Then
      SaveServerINI CInt(L)
      UpDate = False
   End If
Next

'BOOKMARK

frmGetMain.LVServers.ListItems.Clear
If AmConnected = False Then OnLine = False

For L = 1 To ServerCnt
   With SVR
      .IPAddress = LServers(L).IPAddress
      .FailCount = LServers(L).FailCount
      .FilesCount = LServers(L).FilesCount
      .FilesSize = LServers(L).FilesSize
      .HasSlots = LServers(L).HasSlots
      .ISOnline = LServers(L).ISOnline
      .isDisabled = LServers(L).isDisabled
      .ListDate = LServers(L).CurrListDate
      .ListNeeded = LServers(L).ListNeeded
      .NextUpDate = LServers(L).NextUpDate
      .NickName = LServers(L).ServerName
      .SVRHash = LServers(L).SVRHash
      .RecordNum = L
      If .ListNeeded = True Then
         If .isDisabled = False Then
            ADDToGetList CInt(L)
         End If
      End If
    End With
    AddLV frmGetMain.LVServers, SVR
    
Next

frmGetMain.frameServers.Caption = "Servers -- " & frmGetMain.LVServers.ListItems.Count
frmGetMain.File1.Path = MyListFolder
frmGetMain.File1.Refresh
frmGetMain.lblListCount(1).Caption = "Currently " & CStr(frmGetMain.File1.ListCount) & " File Lists For Searching"


End Sub

Public Sub LoadAllServers()
Dim SVR As ListServerInfo
Dim ServerCnt As Integer
Dim L As Long
Dim Z As Long
Dim TMPStr As String
Dim Matched As Boolean
Dim IsThere As Boolean
Dim OnLine As Boolean
Dim DisAbled As Boolean
Dim IPAddr As String
Dim RecordNum As Long
Dim Header As String
Dim SvrNum As String
Dim NickName As String

frmGetMain.LVServers.ListItems.Clear
frmGetMain.LVServers.View = lvwReport

ServerFile = MyServersFolder & "Servers.ini"

ServerCnt = CInt(Val(ReadINI("SERVERS", "SVRCNT", ServerFile))) 'UBound(LServers)
If ServerCnt = 0 Then Exit Sub
ReDim LServers(ServerCnt)
For L = 1 To ServerCnt
   SvrNum = "SVR_" & Format(L, "00#")
   Header = ReadINI("SERVERS", SvrNum, ServerFile)
   With LServers(L)
      .CurrListDate = CLng(Val(ReadINI(Header, "CURRLISTDATE", ServerFile)))
      .FailCount = 0
      .FilesCount = CDbl(Val(ReadINI(Header, "FILESCOUNT", ServerFile)))
      .FilesSize = CDbl(Val(ReadINI(Header, "FILEBYTES", ServerFile)))
      .HasSlots = Word2Bool(ReadINI(Header, "HASSLOTS", ServerFile))
      .IPAddress = ReadINI(Header, "SVRIPADDR", ServerFile)
      .isDisabled = Word2Bool(ReadINI(Header, "DISABLED", ServerFile))
      .ServerName = ReadINI(Header, "SVRNAME", ServerFile)
      NickName = NulTrim(.ServerName)
      .ISOnline = ISOnline(frmGetMain.lstUsers, NickName)
      .ListNeeded = Word2Bool(ReadINI(Header, "LISTNEEDED", ServerFile))
      .NextUpDate = CLng(Val(ReadINI(Header, "NEXTPOLLDATE", ServerFile)))
      .SVRHash = Header
      .Trigger = ReadINI(Header, "TRIGGER", ServerFile)
   End With
Next

'BOOKMARK

frmGetMain.LVServers.ListItems.Clear
If AmConnected = False Then OnLine = False

For L = 1 To ServerCnt
   With SVR
      .IPAddress = LServers(L).IPAddress
      .FailCount = LServers(L).FailCount
      .FilesCount = LServers(L).FilesCount
      .FilesSize = LServers(L).FilesSize
      .HasSlots = LServers(L).HasSlots
      .ISOnline = LServers(L).ISOnline
      .isDisabled = LServers(L).isDisabled
      .ListDate = LServers(L).CurrListDate
      .ListNeeded = LServers(L).ListNeeded
      .NextUpDate = LServers(L).NextUpDate
      .NickName = LServers(L).ServerName
      .SVRHash = LServers(L).SVRHash
    End With
    AddLV frmGetMain.LVServers, SVR
Next

frmGetMain.frameServers.Caption = "Servers -- " & frmGetMain.LVServers.ListItems.Count
frmGetMain.File1.Path = MyListFolder
frmGetMain.lblListCount(1).Caption = "Currently " & CStr(frmGetMain.File1.ListCount) & " File Lists For Searching"


End Sub



Public Sub RAWNamesList(InString As String)
    Dim EachLine() As String
    
    Dim nList() As String
    Dim LineCount As Integer
    Dim L As Long
    Dim TempInfo As String
    Dim X As Long
    Dim Garbage As String
    
    EachLine = Split(InString, Chr(10))
    LineCount = UBound(EachLine)
    For L = 0 To LineCount
      If InStr(EachLine(L), " 353 ") <> 0 Then
            Garbage = nWord(EachLine(L), ":")
            Garbage = nWord(EachLine(L), ":")
      ElseIf InStr(EachLine(L), " 366 ") <> 0 Then
            EachLine(L) = ""
            Exit For
      End If
      TempInfo = TempInfo + EachLine(L) + " "
      Debug.Print EachLine(L)
    Next
    TempInfo = Replace(TempInfo, "  ", " ")
    Debug.Print TempInfo
    nList = Split(TempInfo, " ")
    'frmGetMain.lstUsers.Clear
    
      For X = 0 To UBound(nList)
          nList(X) = Replace(nList(X), vbCr, "")
          If NulTrim(nList(X)) > "" Then
              frmGetMain.lstUsers.AddItem nList(X)
          End If
      Next
      frmGetMain.lblUserCnt.Caption = CStr(frmGetMain.lstUsers.ListCount)
    
    On Error GoTo 0
    Exit Sub

RAWNamesList_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure RAWNamesList of Module modDoBuffer"

End Sub
Public Function ISOnline(lbox As ListBox, Who As String) As Boolean
Dim cCount As Integer
Dim Matched As Boolean
Dim Compare As String
Dim L As Long
'sendToIRC "WHOIS " & Who


cCount = lbox.ListCount
For L = 0 To cCount - 1
    Compare = lbox.List(L)
    Compare = Replace(Compare, "~", "")
    Compare = Replace(Compare, "&", "")
    Compare = Replace(Compare, "@", "")
    Compare = Replace(Compare, "+", "")
    Compare = Replace(Compare, "%", "")

    If UCase(Compare) = UCase(Who) Then
        Matched = True
        ISOnline = True
        Exit Function
    End If
Next
'If Matched = False Then
'    ISOnline = False
'    PaintScreen Who & " is reported OFLINE", vbRed
'End If

End Function

Public Sub GetServerByClass(TravelServer As clsListServer, RecordNum As Long)
Dim L As Long
Dim Handle As Integer
Dim SvrCnt As Integer
Dim SrvrFile As String
Dim Matched As Boolean
Dim Compare As String
Dim RecordNumber As Long
Dim TMPSvr As ListServerInfo

SrvrFile = MyServersFolder & "FetchServers.dat"
RecordNumber = TravelServer.RecordNum
If RecordNum < 1 Then Exit Sub
Handle = FreeFile
Open SrvrFile For Random Shared As Handle Len = Len(TMPSvr)
SvrCnt = LOF(Handle) \ Len(TMPSvr)
Get #Handle, RecordNum, TMPSvr
Close Handle
    
    TravelServer.Deleted = TMPSvr.Deleted
    TravelServer.FilesCount = TMPSvr.FilesCount
    TravelServer.FailCount = TMPSvr.FailCount
    TravelServer.FilesSize = TMPSvr.FilesSize
    TravelServer.HasSlots = TMPSvr.HasSlots
    TravelServer.IPAddress = TMPSvr.IPAddress
    TravelServer.isDisabled = TMPSvr.isDisabled
    TravelServer.ISOnline = TMPSvr.ISOnline
    TravelServer.ListDate = TMPSvr.ListDate
    TravelServer.ListNeeded = TMPSvr.ListNeeded
    TravelServer.LSNickName = TMPSvr.NickName
    TravelServer.NextUpDate = TMPSvr.NextUpDate
    TravelServer.RecordNum = TMPSvr.RecordNum
    
End Sub


Public Function GetServerByHash(LookHash As String) As Long
    ' Return Record Number for Given Nickname Hash

Dim L As Long
Dim SvrCnt As Integer
Dim Matched As Boolean

   On Error GoTo GetServerByHash_Error

SvrCnt = UBound(LServers)
If SvrCnt = 0 Then
    GetServerByHash = -1
    Exit Function
End If

For L = 1 To SvrCnt
    If NulTrim(LServers(L).SVRHash) = LookHash Then
        GetServerByHash = L
        Matched = True
        Exit For
    End If
Next
If Matched = False Then
    GetServerByHash = -1
End If

   On Error GoTo 0
   Exit Function

GetServerByHash_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure GetServerByHash of Module modDefines", "Debuglog.txt"
End Function

Public Sub GetServerByRecord(RecordNum As Long, sServer As ListServerInfo)
Dim Handle As Integer
Dim SvrCnt As Integer
Dim SrvrFile As String

If RecordNum = 0 Then Exit Sub
SrvrFile = MyServersFolder & "FetchServers.dat"
Handle = FreeFile
Open SrvrFile For Random Shared As Handle Len = Len(sServer)
SvrCnt = LOF(Handle) \ Len(sServer)
If SvrCnt = 0 Then Exit Sub
Get #Handle, RecordNum, sServer
Close Handle
End Sub

Public Sub PutServerByClass(tmpServer As clsListServer, RecordNum As Long)
Dim ThisServer As ListServerInfo
Dim TMPSvr As ListServerInfo
Dim NickName As String
Dim L As Long
Dim Record As Long
Dim Handle As Integer
Dim SvrCnt As Integer
Dim SrvrFile As String
Dim Matched As Boolean
Dim Compare As String
Dim Compare1 As String

Exit Sub

With ThisServer
    .Deleted = tmpServer.Deleted
    .FailCount = tmpServer.FailCount
    .FilesCount = tmpServer.FilesCount
    .FilesSize = tmpServer.FilesSize
    .HasSlots = tmpServer.HasSlots
    .IPAddress = tmpServer.IPAddress
    .isDisabled = tmpServer.isDisabled
    .ISOnline = tmpServer.ISOnline
    .ListDate = tmpServer.ListDate
    .ListNeeded = tmpServer.ListNeeded
    .NickName = tmpServer.LSNickName
    .NextUpDate = tmpServer.NextUpDate
    .RecordNum = tmpServer.RecordNum
    If RecordNum <> .RecordNum Then .RecordNum = RecordNum
    Record = .RecordNum
    NickName = .NickName
    
End With

SrvrFile = MyServersFolder & "FetchServers.dat"
Handle = FreeFile
Open SrvrFile For Random Shared As Handle Len = Len(ThisServer)
SvrCnt = LOF(Handle) \ Len(ThisServer)

If Record <> -1 Then
    Put #Handle, Record, ThisServer
    Close Handle
    Exit Sub
End If

If SvrCnt = 0 Then
    With ThisServer
        .Deleted = False
        .FilesCount = 0
        .FilesSize = 0
        .IPAddress = tmpServer.IPAddress
        .isDisabled = False
        .ISOnline = True
        .ListNeeded = True
        .NickName = tmpServer.LSNickName
        .RecordNum = 1
        RecordNum = .RecordNum
    End With
    Put #Handle, 1, ThisServer
    Close Handle
    Exit Sub
End If

End Sub

Public Function PutNewServer(SVR As ListServerInfo) As Long
Dim L As Long
Dim Record As Long
Dim sSVR As ListServerInfo
Dim Handle As Integer
Dim SvrCnt As Integer
Dim SrvrFile As String
Dim Matched As Boolean


SrvrFile = MyServersFolder & "FetchServers.dat"
Handle = FreeFile
Open SrvrFile For Random Shared As Handle Len = Len(SVR)
SvrCnt = LOF(Handle) \ Len(SVR)
For L = 1 To SvrCnt
    Get #Handle, L, sSVR
    If sSVR.Deleted = True Then
        Matched = True
        Record = L
        Exit For
    End If
Next

If Matched = True Then
    SVR.RecordNum = Record
    Put #Handle, Record, SVR
Else
    Record = SvrCnt + 1
    SVR.RecordNum = Record
    Put #Handle, Record, SVR
End If
PutNewServer = Record
Close Handle
End Function

Public Sub PutServerByRecord(RecordNum As Long, SVR As ListServerInfo)
Dim Handle As Integer
Dim SvrCnt As Integer
Dim SrvrFile As String

SrvrFile = MyServersFolder & "FetchServers.dat"
Handle = FreeFile
Open SrvrFile For Random Shared As Handle Len = Len(SVR)
SvrCnt = LOF(Handle) \ Len(SVR)
SVR.RecordNum = RecordNum
Put #Handle, RecordNum, SVR
Close Handle
End Sub
