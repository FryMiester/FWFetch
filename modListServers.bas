Attribute VB_Name = "modListServers"
Option Explicit

Type ListServerInfo
    IPAddress       As String * 50
    SVRHash         As String * 12
    SpaceSave       As String * 21
    ListDate        As Long
    NextUpDate      As Long
    NickName        As String * 46
    FailCount       As Integer
    HasSlots        As Boolean
    FilesSize       As Double
    FilesCount      As Long
    isDisabled      As Boolean
    ListNeeded      As Boolean
    ISOnline        As Boolean
    Deleted         As Boolean
    RecordNum       As Long
    Storage         As String * 150
End Type

Type ListServer
    ServerName      As String * 50
    Trigger         As String * 50
    IPAddress       As String * 50
    SVRHash         As String * 12
    CurrListDate    As Long
    NextUpDate      As Long
    FailCount       As Integer
    HasSlots        As Boolean
    FilesSize       As Double
    FilesCount      As Double
    isDisabled      As Boolean
    ListNeeded      As Boolean
    ISOnline        As Boolean
End Type

Global LServers() As ListServer

Public Sub ProcessSlotMsg(SlotMsg As String)
' The whole string has been sent, process it
Dim MD5 As CMD5
Dim Hash As String
Dim Garbage As String
Dim Nick As String
Dim SVRIP As String
Dim fCount As String
Dim fSize As String
Dim SlotListDate As Long
Dim SlotMsgBAK As String
Dim Work As String
Dim L As Long
Dim sCount As Integer
Dim Matched As Boolean
Dim LNeeded As Boolean
Dim ServNum As Integer
Set MD5 = New CMD5

SlotMsgBAK = SlotMsg
Work = SlotMsg
'19:22:05 12/11/2023 :peapod!peapod@ihw-lqi.c2r.40.8.IP PRIVMSG #ebooks :SLOTS 20 20 NOW 0 999 0 609824 1707772659361 1 1702215341 2282862 FWServer ver 2071

sCount = UBound(LServers)

If sCount = 0 Then
   ' Add First Server
   AddServerINI SlotMsgBAK
   ServerPoll
   Exit Sub
End If

Work = Right(Work, Len(Work) - 1)
Nick = nWord(Work, "!")
Hash = MD5.HaxMD5(Nick)
Garbage = nWord(Work, "@")
SVRIP = nWord(Work, " ")
Garbage = nWord(Work, ":")
For L = 1 To 7
   Garbage = nWord(Work, " ")
Next
fCount = nWord(Work, " ")
fSize = nWord(Work, " ")
Garbage = nWord(Work, " ")
SlotListDate = CLng(Val(nWord(Work, " ")))

' Get Record if it exists and comepare to see if a list is needed
For L = 1 To sCount
   If NulTrim(LServers(L).SVRHash) = Hash Then
      ServNum = L
      Matched = True
      Exit For
   End If
Next

If Matched = False Then

   AddServerINI SlotMsgBAK
   'poll for list
   'ServerPoll
   Exit Sub
End If

With LServers(ServNum)
   If .CurrListDate < SlotListDate Then LNeeded = True
   If .FilesCount < CDbl(Val(fCount)) Then LNeeded = True
   If .FilesSize < CDbl(Val(fSize)) Then LNeeded = True
End With
SysLog "Processing SLOTS Message from " & Nick & " - List Needed = " & LNeeded, vbRed

   ' Post Changes - Request New List
   LServers(ServNum).CurrListDate = SlotListDate
   LServers(ServNum).FilesCount = CDbl(fCount)
   LServers(ServNum).FilesSize = CDbl(fSize)
   LServers(ServNum).HasSlots = True
   SaveServerINI CInt(ServNum)


End Sub

Public Sub AddServerINI(SlotMsg As String, Optional NotASLOT As Boolean, Optional Rtn As Integer)
Dim MD5 As CMD5
Dim Hash As String
Dim Garbage As String
Dim Nick As String
Dim SVRIP As String
Dim fCount As String
Dim fSize As String
Dim SlotListDate As Long
Dim SlotMsgBAK As String
Dim Work As String
Dim L As Long
Dim sCount As Integer
Dim INISrvCnt As Integer
Dim Header As String
Dim SrvHead As String
Dim InIFile As String
Dim HasSlots As Boolean
Dim FailCount As Integer

Set MD5 = New CMD5

SlotMsgBAK = SlotMsg
Work = SlotMsg
'19:22:05 12/11/2023 :peapod!peapod@ihw-lqi.c2r.40.8.IP PRIVMSG #ebooks :SLOTS 20 20 NOW 0 999 0 609824 1707772659361 1 1702215341 2282862 FWServer ver 2071
sCount = UBound(LServers)

sCount = sCount + 1
ReDim Preserve LServers(sCount)
InIFile = MyServersFolder & "Servers.ini"

Work = Right(Work, Len(Work) - 1)
Nick = nWord(Work, "!")
Hash = MD5.HaxMD5(Nick)
Garbage = nWord(Work, "@")
SVRIP = nWord(Work, " ")
Garbage = nWord(Work, ":")
For L = 1 To 7
   Garbage = nWord(Work, " ")
Next

If NotASLOT = False Then
   fCount = nWord(Work, " ")
   fSize = nWord(Work, " ")
   Garbage = nWord(Work, " ")
   SlotListDate = CLng(Val(nWord(Work, " ")))
   HasSlots = True
Else
   fCount = "0"
   fSize = "0"
   SlotListDate = Date2Epoch(Now)
   Rtn = sCount
   HasSlots = False
   
End If

With LServers(sCount)
   .CurrListDate = SlotListDate
   .FailCount = 0
   .FilesCount = CDbl(Val(fCount))
   .FilesSize = CDbl(Val(fSize))
   .HasSlots = HasSlots
   .IPAddress = SVRIP
   .SVRHash = Hash
   .isDisabled = False
   .ISOnline = True
   .ListNeeded = True
   .ServerName = Nick
   .Trigger = Nick
   .FailCount = 0
End With

' Prepare INI stuff
' We already know the record does not exist
INISrvCnt = CInt(Val(ReadINI("SERVERS", "SVRCNT", InIFile)))
INISrvCnt = INISrvCnt + 1

Header = "SVR_" & Format(INISrvCnt, "00#")
WriteINI "SERVERS", "SVRCNT", CStr(INISrvCnt), InIFile
WriteINI "SERVERS", Header, Hash, InIFile

WriteINI Hash, "SVRNAME", Nick, InIFile
WriteINI Hash, "TRIGGER", "@" & Nick, InIFile
WriteINI Hash, "SVRIPADDR", SVRIP, InIFile
WriteINI Hash, "CURRLISTDATE", CStr(SlotListDate), InIFile
WriteINI Hash, "FAILCOUNT", CStr(FailCount), InIFile
WriteINI Hash, "FILEBYTES", fSize, InIFile
WriteINI Hash, "FILESCOUNT", fCount, InIFile
WriteINI Hash, "NEXTPOLLDATE", "-1", InIFile
WriteINI Hash, "HASSLOTS", Bool2Word(HasSlots), InIFile
WriteINI Hash, "DISABLED", "FALSE", InIFile
WriteINI Hash, "LISTNEEDED", "TRUE", InIFile

End Sub


Public Sub SaveServerINI(ServerNum As Integer)
Dim MD5 As CMD5
Dim Hash As String
Dim Garbage As String
Dim Nick As String
Dim SVRIP As String
Dim fCount As String
Dim fSize As String
Dim SlotListDate As Long
Dim SlotMsgBAK As String
Dim Work As String
Dim L As Long
Dim sCount As Integer
Dim INISrvCnt As Integer
Dim Header As String
Dim SrvHead As String
Dim InIFile As String
Dim LNeeded As Boolean
Dim UpDateDay As Long
Dim DisAbled As Boolean
Dim HasSlots As Boolean
Dim FailCount As Integer

InIFile = MyServersFolder & "Servers.ini"

With LServers(ServerNum)
   SlotListDate = .CurrListDate
   fCount = CStr(.FilesCount)
   fSize = CStr(.FilesSize)
   Hash = NulTrim(.SVRHash)
   LNeeded = .ListNeeded
   UpDateDay = .NextUpDate
   DisAbled = .isDisabled
   HasSlots = .HasSlots
   FailCount = .FailCount
End With

WriteINI Hash, "CURRLISTDATE", CStr(SlotListDate), InIFile
WriteINI Hash, "NEXTPOLLDATE", CStr(UpDateDay), InIFile

WriteINI Hash, "FAILCOUNT", CStr(FailCount), InIFile
WriteINI Hash, "FILEBYTES", fSize, InIFile
WriteINI Hash, "FILESCOUNT", fCount, InIFile
WriteINI Hash, "LISTNEEDED", Bool2Word(LNeeded), InIFile
WriteINI Hash, "DISABLED", Bool2Word(DisAbled), InIFile
WriteINI Hash, "HASSLOTS", Bool2Word(HasSlots), InIFile

End Sub

Public Sub LoadServerINI(HashCode As String, Index As Integer)
Dim MD5 As CMD5
Dim Hash As String
Dim Garbage As String
Dim Nick As String
Dim SVRIP As String
Dim fCount As String
Dim fSize As String
Dim SlotListDate As Long
Dim SlotMsgBAK As String
Dim Work As String
Dim L As Long
Dim sCount As Integer
Dim INISrvCnt As Integer
Dim Header As String
Dim SrvHead As String
Dim InIFile As String

InIFile = MyServersFolder & "Servers.ini"
Header = HashCode

With LServers(Index)
   .CurrListDate = CDbl(Val(ReadINI(Header, "CURRLISTDATE", InIFile)))
   .FilesCount = ReadINI(Header, "FILESCOUNT", InIFile)
   .FilesSize = ReadINI(Header, "FILEBYTES", InIFile)
   .FailCount = CInt(Val(ReadINI(Header, "FAILCOUNT", InIFile)))
   
   .HasSlots = Word2Bool(ReadINI(Header, "HASSLOTS", InIFile))
   .IPAddress = ReadINI(Header, "SVRIPADDR", InIFile)
   .isDisabled = Word2Bool(ReadINI(Header, "DISABLED", InIFile))
   .ListNeeded = Word2Bool(ReadINI(Header, "LISTNEEDED", InIFile))
   .NextUpDate = CLng(ReadINI(Header, "NEXTPOLLDATE", InIFile))
   .ServerName = ReadINI(Header, "SVRNAME", InIFile)
   If AmConnected = True Then
          .ISOnline = ISOnline(frmGetMain.lstUsers, NulTrim(.ServerName))
      Else
          .ISOnline = False
      End If
   .Trigger = ReadINI(Header, "TRIGGER", InIFile)
End With

End Sub


Public Sub ProcessAdvert(AdvertMsg As String)
Dim intLength As Integer
Dim Garbage As String
Dim FileListCnt As Long
Dim FileListDate As String
Dim TheYear As String
Dim ListDate As Date
Dim DateLong As Long
Dim ListCntLong As Long
Dim Marker As Integer
Dim Marker0 As Integer
Dim NeedList As Boolean
Dim Success As Boolean
Dim TMPStr As String
Dim fDate As Long
Dim fCount As Long
Dim fSize As Double
Dim IPAddr As String
Dim HasSlots As Boolean
Dim NickName As String
Dim Hash As String
Dim L As Long
Dim Matched As Boolean
Dim AdvertBAK As String
Dim SrvCnt As Integer
Dim SrvNum As Integer

Dim MD5 As CMD5
Set MD5 = New CMD5

Dim SvrInfo As clsListServer
Set SvrInfo = New clsListServer


   'On Error GoTo ProcessAdvert_Error

'AdvertMsg = Replace(AdvertMsg, "  ", " &")

' Converting to using LServers(index) in memory
' If Server is new it has to be treated differently than with slots
' I would prefer to only deal with slots

SrvCnt = UBound(LServers)
AdvertBAK = AdvertMsg

TMPStr = Right(AdvertBAK, Len(AdvertBAK) - 1)
NickName = nWord(TMPStr, "!")
Garbage = nWord(TMPStr, "@")
IPAddr = nWord(TMPStr, " ")
Hash = MD5.HaxMD5(NickName)

For L = 1 To SrvCnt
   If NulTrim(LServers(L).SVRHash) = Hash Then
      SrvNum = L
      Matched = True
      Exit For
   End If
Next
If Matched = False Then
' gotta create a record
   AddServerINI AdvertMsg, True, SrvNum
   
End If



IPAddr = NulTrim(LServers(SrvNum).IPAddress)
NickName = NulTrim(LServers(SrvNum).ServerName)
fDate = LServers(SrvNum).CurrListDate
fCount = LServers(SrvNum).FilesCount
fSize = LServers(SrvNum).FilesSize
HasSlots = LServers(SrvNum).HasSlots

If InStr(UCase(AdvertMsg), "TYPE:") = 0 Then Exit Sub
If HasSlots = True Then Exit Sub

If Date2Epoch(Now) < LServers(SrvNum).NextUpDate Then
    Exit Sub
End If

':peapod!peapod@ihw-lqi.c2r.40.8.IP PRIVMSG #ebooks :04,0014,00• 12,00Type:04,00 @peapod 12,00For My List Of:04,00 608,743 12,00Files 14,00• 12,00Slots:04,00 20 14,00• 12,00Queued:04,00 0 14,00• 12,00Speed:04,00 472 KB/Sec 14,00• 12,00Next: 04,00NOW 14,00• 12,00Served:04,00 195,764 14,00• 12,00List: 04,0012-02-2023 14,00• 12,00Search: 04,00ON 14,00• 12,00Mode: Ser FWServerBot

' find file list count info:

Marker = InStr(UCase(AdvertMsg), "LIST OF:")
Marker0 = InStr(UCase(AdvertMsg), "FILES")
If Marker > 0 Then
    TMPStr = Mid(AdvertMsg, Marker, Marker0 - Marker)
    Garbage = nWord(TMPStr, " ")
    Garbage = nWord(TMPStr, " ")
    TMPStr = nWord(TMPStr, " ")
    TMPStr = Replace(TMPStr, ",", "")
    AdvertMsg = Right(AdvertMsg, Len(AdvertMsg) - Marker0)
    FileListCnt = CLng(Val(TMPStr))
End If

' find file list DATE info:
Marker = InStr(UCase(AdvertMsg), "LIST:")
If Marker = 0 Then
    Marker = InStr(UCase(AdvertMsg), "UPDATED:")
End If

If Marker > 0 Then
    AdvertMsg = NulTrim(Right(AdvertMsg, Len(AdvertMsg) - Marker))
    Garbage = nWord(AdvertMsg, " ")
    TMPStr = nWord(AdvertMsg, " ")
    If NulTrim(TMPStr) = "" Then
        TMPStr = nWord(AdvertMsg, " ")
    End If
    TMPStr = NulTrim(TMPStr)
    If Len(TMPStr) = 8 Then
        If Mid(TMPStr, 3, 1) = "," Then
            TMPStr = Right(TMPStr, 3)
        End If
    End If
    
    Select Case Len(TMPStr)
    Case 3
        Select Case LCase(TMPStr)
            Case "jan"
                FileListDate = "01-"
            Case "feb"
                FileListDate = "02-"
            Case "mar"
                FileListDate = "03-"
            Case "apr"
                FileListDate = "04-"
            Case "may"
                FileListDate = "05-"
            Case "jun"
                FileListDate = "06-"
            Case "jul"
                FileListDate = "07-"
            Case "aug"
                FileListDate = "08-"
            Case "sep"
                FileListDate = "09-"
            Case "oct"
                FileListDate = "10-"
            Case "nov"
                FileListDate = "11-"
            Case "dec"
                FileListDate = "12-"
        End Select
        TMPStr = nWord(AdvertMsg, " ")
        TMPStr = Left(TMPStr, Len(TMPStr) - 2)
        If Len(TMPStr) = 1 Then TMPStr = "0" & TMPStr
        TheYear = Format(Now, "YYYY")
        If CDate(FileListDate & TMPStr & "-" & TheYear) > Now Then
            TheYear = CStr(CInt(TheYear) - 1)
        End If
        FileListDate = FileListDate & TMPStr & "-" & TheYear
        
    Case 10
        FileListDate = TMPStr
    Case 15
        TMPStr = Right(TMPStr, 10)
        FileListDate = TMPStr
    Case Else
    End Select
End If

LServers(SrvNum).ISOnline = True
If FileListDate = "" Then FileListDate = Now

ListDate = CDate(FileListDate)
DateLong = Date2Epoch(ListDate)
ListCntLong = CLng(FileListCnt)
If ListCntLong > fCount Then NeedList = True
If DateLong > fDate Then NeedList = True

If NeedList = True Then
    LServers(SrvNum).ListNeeded = True
Else
    LServers(SrvNum).ListNeeded = False
End If

LServers(SrvNum).FilesCount = ListCntLong
LServers(SrvNum).CurrListDate = DateLong
LServers(SrvNum).ISOnline = True

'SysLog "Processing " & NickName & " NON-SLOTS Message, List needed = " & Bool2Word(NeedList)

SaveServerINI SrvNum

   On Error GoTo 0
   Exit Sub

ProcessAdvert_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessAdvert of Module modDefines"

End Sub

Public Sub ADDToGetList(SvrNum As Integer)
Dim LstCnt As Integer
Dim L As Long
Dim Compare As String
Dim Item As Long
Dim Matched As Boolean

LstCnt = frmGetMain.lstGetNewList.ListCount
For L = 0 To LstCnt - 1
   Item = frmGetMain.lstGetNewList.ItemData(L)
   If Item = CLng(SvrNum) Then
      Matched = True
      Exit For
   End If
Next
If Matched = True Then Exit Sub
LstCnt = (LstCnt - 1) + 1
frmGetMain.lstGetNewList.AddItem NulTrim(LServers(SvrNum).Trigger), LstCnt
frmGetMain.lstGetNewList.ItemData(LstCnt) = SvrNum

End Sub
