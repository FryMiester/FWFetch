VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ctlSVRDownLoad 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   ScaleHeight     =   840
   ScaleWidth      =   3765
   Begin VB.Timer tmrstartup 
      Enabled         =   0   'False
      Left            =   3000
      Top             =   480
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   480
   End
   Begin VB.Timer tmrPollInfo 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1815
      Top             =   465
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1185
      Top             =   480
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   480
   End
   Begin MSWinsockLib.Winsock sckDCC 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox PB1 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Timer tmrBailOut 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   90
         Top             =   90
      End
   End
End
Attribute VB_Name = "ctlSVRDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DoubleClick()

Dim WaitLoops As Integer
Dim TimeOutLoops As Integer
Dim ServerNum As Integer
Dim BailOut As Integer
Dim TotalBytes As Long

Private udtThisServer As ListServerInfo
Private CurrentServer As clsListServer
Private fBlock As myFileBlock
Private inText As String
Private m_MyIndex As Integer
Private m_AmActive As Boolean
Private m_GetListFile As String
Private m_Filename As String
Private m_FileSize As Long
Private m_IP_Info As String
Private m_UserNick As String
Private m_ServerName As String
Private m_IsFileList As Boolean
Private m_IsOnline As Boolean


Public Sub Activate(clsServer As clsListServer, BuffLine As String)
' Main Program Says 'I Have A File For You' Here's the Specs
'    UpdateDLStatus PB1, 100, 3, "Connecting ", vbGreen
Set CurrentServer = New clsListServer

With CurrentServer
    .Deleted = clsServer.Deleted
    .FailCount = clsServer.FailCount
    .FilesCount = clsServer.FilesCount
    .FilesSize = clsServer.FilesSize
    .HasSlots = clsServer.HasSlots
    .IPAddress = clsServer.IPAddress
    .isDisabled = clsServer.isDisabled
    .ISOnline = clsServer.ISOnline
    .ListDate = clsServer.ListDate
    .ListNeeded = clsServer.ListNeeded
    .LSNickName = clsServer.LSNickName
    .NextUpDate = clsServer.NextUpDate
    .RecordNum = clsServer.RecordNum
End With

ProcessFile CurrentServer, BuffLine

End Sub



Public Property Get MyIndex() As Integer
    MyIndex = m_MyIndex
End Property

Public Property Let MyIndex(ByVal New_MyIndex As Integer)
    m_MyIndex = New_MyIndex
    PropertyChanged "MyIndex"
End Property

Public Property Get AmActive() As Boolean
    AmActive = m_AmActive
End Property

Public Property Let AmActive(ByVal New_AmActive As Boolean)
    m_AmActive = New_AmActive
    PropertyChanged "AmActive"
End Property

Public Property Get ISOnline() As Boolean
    ISOnline = m_IsOnline
End Property

Public Property Let ISOnline(ByVal New_IsOnline As Boolean)
    m_IsOnline = New_IsOnline
    PropertyChanged "IsOnline"
End Property

Public Property Get IsFileList() As Boolean
    IsFileList = m_IsFileList
End Property

Public Property Let IsFileList(ByVal New_IsFileList As Boolean)
    m_IsFileList = New_IsFileList
    PropertyChanged "IsFileList"
End Property

Public Property Get GetListFile() As String
    GetListFile = m_GetListFile
End Property

Public Property Let GetListFile(ByVal New_GetListFile As String)
    m_GetListFile = New_GetListFile
    PropertyChanged "GetListFile"
End Property

Public Property Get FileName() As String
    FileName = m_Filename
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_Filename = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get ServerName() As String
    ServerName = m_ServerName
End Property

Public Property Let ServerName(ByVal New_ServerName As String)
    m_ServerName = New_ServerName
    PropertyChanged "ServerName"
End Property


Public Sub Initialize()
Dim Rando As Integer
Randomize Timer
Rando = rand(3, 30)
tmrstartup.Interval = Rando * 1000
tmrstartup.Enabled = True
End Sub


Public Sub DoubleClick()
FinishUp 3, CurrentServer
End Sub


Private Sub ProcessFile(clsServer As clsListServer, InString As String)
Dim ArchPath As String
Dim TMPStr As String
Dim OutputFile As String
Dim NickName As String
Dim fName As String
Dim IPAddr As String
Dim Port As String
Dim fSize As Long
Dim CurrSize As Long
Dim MyMsg As String

TMPStr = InString
ProcessDCCSend TMPStr, NickName, fName, IPAddr, Port, fSize
inText = NickName
    
    
    fBlock.ServerName = NickName
    ArchPath = MyInComingFolder
    OutputFile = ArchPath & fName
    fBlock.FileName = OutputFile
    fBlock.FileSize = fSize
    fBlock.IPAddr = IPAddr
    fBlock.PortNum = Port

If FileExist(OutputFile) = True Then
    CurrSize = FileLen(OutputFile)
    fBlock.Position = CurrSize
    MyMsg = "DCC RESUME " & QtStr(fName) & " " & Port & " " & CStr(CurrSize)
    IRC_CTCP_MSG NickName, MyMsg
    ' Request Resume
End If

ReceiveFile clsServer, InString

End Sub



Private Sub UpdateDLStatus(pic As PictureBox, ByVal sngPercent As Single, Optional DisplayMode As Integer, Optional strText As String, Optional bColor As Long)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intPercent As Integer


   On Error GoTo UpdateDLStatus_Error

    pic.ForeColor = &HC00000
    pic.BackColor = vbWhite
    
    intPercent = PercentOf(100, sngPercent)
    sngPercent = intPercent * 0.01
    
    Select Case DisplayMode
        Case -1                         '  Reset
            pic.BackColor = &H80000004
            pic.ForeColor = &H80000004
            strPercent = ""
        Case 0
            strPercent = Format$(intPercent) & "%"
        Case 1
            If strText > "" Then
                strPercent = strText
            Else
                strPercent = Format$(intPercent) & "%"
            End If
        Case 2
            strPercent = Format$(intPercent) & " % " & strText
        Case 3
            If bColor <> 0 Then
                pic.ForeColor = bColor
            End If
            strPercent = strText & "  " & Format$(intPercent) & "%"
        Case 10
            pic.FontSize = 10
            pic.BackColor = vbBlue
            pic.ForeColor = vbYellow
            strPercent = strText
        Case Else
            strPercent = Format$(intPercent) & "%"
    End Select
    
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent
    
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If
    
    pic.Refresh
    DoEvents
    
    If DisplayMode = -1 Then
        pic.BackColor = &H80000004
        pic.Refresh
    End If

   On Error GoTo 0
   Exit Sub


   On Error GoTo 0
   Exit Sub

UpdateDLStatus_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateDLStatus of User Control ctlDownLoad", "Debuglog.txt"

 End Sub


Private Sub PB1_DblClick()
FinishUp 3, CurrentServer
End Sub

Private Sub sckDCC_Connect()
Dim Index As Integer
Dim Tmp As String
TotalBytes = 0
tmrTimeOut.Enabled = False

With fBlock
    .fHandle = FreeFile
    Open NulTrim(.FileName) For Binary As .fHandle
    If .Position <> 0 Then
        Seek #.fHandle, .Position + 1
    End If
End With
AmActive = True
tmrWait.Enabled = False

End Sub

Private Sub sckDCC_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Dim Handle As Integer
Dim Success As Boolean
Dim intPercent As Integer
Static Stalled As Integer
Static oldPercent As Integer

   On Error GoTo sckDCC_DataArrival_Error

DoEvents
Handle = fBlock.fHandle
sckDCC.GetData sData, vbString
TotalBytes = TotalBytes + Len(sData)
'InfoLog "Total Bytes : " & CStr(TotalBytes) & " Buffer : " & CStr(Len(sData)), App.Path & "\DCCServerInfo.txt"
 
Put #Handle, , sData
intPercent = PercentOf(CDbl(fBlock.FileSize), CDbl(TotalBytes))
UpdateDLStatus PB1, intPercent, 3, NoPath(NulTrim(fBlock.FileName))
If intPercent > oldPercent Then
   oldPercent = intPercent
   TimeOutLoops = 0
   BailOut = 0
End If


fBlock.TotalBytes = TotalBytes

If sckDCC.State = sckClosing Or sckDCC.State = sckClosed Then
   sckDCC.Close
   Close fBlock.fHandle
   Close Handle
   TotalBytes = 0
   Stalled = 0
   FinishUp 1, CurrentServer
   Exit Sub
End If


If TotalBytes >= fBlock.FileSize Then
    TotalBytes = 0
    sckDCC.Close
    Close Handle
    Stalled = 0
    FinishUp 1, CurrentServer
End If

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

sckDCC_DataArrival_Error:
If Debugger = False Then
   Debugger = True
'   DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure sckDCC_DataArrival of User Control ctlSVRDownLoad", App.Path & "\Debuglog.txt"
   Debugger = False
Else
   DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure sckDCC_DataArrival of User Control ctlSVRDownLoad", App.Path & "\Debuglog.txt"
End If

End Sub

Public Sub GetWhois(InPutStr As String)
If InStr(InPutStr, " 311 ") <> 0 Then
    ISOnline = True
Else
    ISOnline = False
End If
End Sub


Public Sub NewPoll(INNum As Integer, Optional Forced As Boolean)
Dim ReqStr As String
Dim Tmp As String
Dim tmpfile As String
Dim tmpDate As Date
Dim TmpNum As Long
Dim tmpNow As Long
Dim OnLine As Boolean
Dim Trigger As String

Debug.Print "AmActive = " & AmActive

If SVRPause = True Then Exit Sub
If AmActive = True Then Exit Sub
If AmConnected = False Then Exit Sub
If AmPaused = True Then Exit Sub

ServerNum = INNum
With LServers(ServerNum)
   ServerName = NulTrim(LServers(ServerNum).ServerName)
   If .isDisabled Then Exit Sub
   OnLine = LServers(ServerNum).ISOnline
   Trigger = NulTrim(.Trigger)
End With

If LServers(ServerNum).isDisabled = True Then Exit Sub

If Forced = False Then
    If Date2Epoch(Now) < LServers(ServerNum).NextUpDate Then Exit Sub
End If

If ServerName = "" Then Exit Sub

If OnLine = True Then
   SysLog "Requesting " & Trigger
   PaintScreen "Requesting " & Trigger, vbGreen
   tmrBailOut.Enabled = True
   sckDCC.Close
   Close fBlock.fHandle
   ProcessFLRequest Trigger
Else
    SysLog "Server " & ServerName & " is Off-line", vbRed
End If
End Sub


Private Sub ProcessFLRequest(FileInfo As String)
'FileName = NulTrim(MyDownloadFolder) & NulTrim(FileInfo)
IsFileList = True
tmrTimeOut.Enabled = False
tmrWait.Enabled = True
IRCSay MyChannel, FileInfo
AmActive = True
End Sub

Private Sub tmrBailOut_Timer()
BailOut = BailOut + 1

If BailOut = 180 Then
   BailOut = 0
   AmActive = False
   tmrBailOut.Enabled = False
   FinishUp -1, CurrentServer
End If
End Sub

Private Sub tmrTimeOut_Timer()
Dim IntPcnt As Integer
DoEvents
TimeOutLoops = TimeOutLoops + 1
IntPcnt = PercentOf(90, TimeOutLoops)
inText = CStr(60 - TimeOutLoops)
Select Case TimeOutLoops
Case Is < 20
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting..." & inText, vbGreen
Case 20 To 45
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting..." & inText, vbCyan
Case 45 To 69
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting..." & inText, vbYellow
Case 70 To 90
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting..." & inText, vbRed
End Select

If TimeOutLoops = 60 Then
    ' Connection Timed out
    UpdateDLStatus PB1, 100, 3, "TIME OUT", vbRed
    Pause 3
    UpdateDLStatus PB1, 100, -1, ""
    TimeOutLoops = 0
    tmrTimeOut.Enabled = False
    FinishUp -1, CurrentServer
End If
End Sub

Private Sub tmrWait_Timer()
WaitLoops = WaitLoops + 1
If WaitLoops >= 240 Then
    FinishUp 3, CurrentServer
End If
End Sub

Private Sub ReceiveFile(clsServer As clsListServer, FullFileName As String)
Dim fHandle As Integer
Dim IPAddr As String
Dim Port As String

   On Error GoTo ReceiveFile_Error

tmrTimeOut.Enabled = True
TimeOutLoops = 0

IPAddr = fBlock.IPAddr
Port = fBlock.PortNum
fHandle = fBlock.fHandle


If IPAddr = "" Or Port = "" Then
    ' Abort Transfer
    FinishUp -1, clsServer
    Exit Sub
End If

'DebugLog IpAddr & "  " & Port
With sckDCC
    If .State <> sckClosed Then
        .Close
    End If
    .Connect NulTrim(fBlock.IPAddr), CStr(fBlock.PortNum)
End With
Pause 5
Do
DoEvents

Loop While sckDCC.State = sckConnected
Close fHandle



   On Error GoTo 0
   Exit Sub

ReceiveFile_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ReceiveFile of User Control ctlDownLoad", "Debuglog.txt"

End Sub


Private Sub ProcessNewList(InlistName As String, ThisServer As clsListServer)
Dim OutputName As String
Dim RtnVal As Double
Dim CommandLine As String
Dim ArchPath As String
Dim Ext As String
Dim strInput As String
Dim Success As Long
Dim ServerName As String
Dim Saved As Long
Dim ArchiveFile As String
Dim ListCnt As Integer
Dim L As Long
Dim IPAddr As String
Dim NickName As String
Dim AddOn As Long
Dim tName As String

sckDCC.Close
Close fBlock.fHandle
UpdateDLStatus PB1, 100, -1, ""
tmrTimeOut.Enabled = False
TimeOutLoops = 0
tmrBailOut.Enabled = False
BailOut = 0

AddOn = (86400 * CInt(UpDateDays))
'frmGetMain.tmrTransfer.Enabled = False
'ININame = MyServersFolder & "Servers.ini"
ServerName = NulTrim(fBlock.ServerName)

ArchPath = RunPath(InlistName)
'ArchPath = PathHeader & "\Incoming\"
ArchiveFile = InlistName

Ext = FileExten(InlistName)
Select Case LCase(Ext)
    Case ".zip"
        CommandLine = App.Path & "\7za.exe" & " e -y " & Chr(34) & InlistName & Chr(34)
        ChDir ArchPath
        GoSub Extractor
        InlistName = noExt(InlistName) & ".txt"
    Case ".rar"
        SysLog "Processing a RAR File"
        CommandLine = App.Path & "\7za.exe" & " e -y " & Chr(34) & InlistName & Chr(34)
        ChDir ArchPath
        GoSub Extractor
        InlistName = noExt(InlistName) & ".txt"
    Case ".txt"
    Case Else
End Select
Pause 2

OutputName = MyListFolder & "!" & ServerName & ".txt"
For L = 0 To 10
   tName = MyListFolder & "!" & ServerName & CStr(L) & ".txt"
   If FileExist(tName) = True Then Kill tName
Next
   
If FileExist(OutputName) = True Then Kill OutputName

BreakTheFile InlistName, OutputName

LServers(ServerNum).isDisabled = False
LServers(ServerNum).FailCount = 0

LServers(ServerNum).ListNeeded = False
LServers(ServerNum).CurrListDate = Date2Epoch(Now)
LServers(ServerNum).NextUpDate = Date2Epoch(Now) + AddOn

SaveServerINI ServerNum

If FileExist(ArchiveFile) Then
    Kill ArchiveFile
End If

XfrTimer = 0
SysLog "Processed List from " & ServerName
FileListBlock.Active = False
ServerPoll
AmActive = False
Exit Sub
Extractor:
    RtnVal = Shell(CommandLine, vbHide)
Return
End Sub

Private Sub BreakTheFile(ByVal StrInfile As String, strOutFile As String)
Dim sText As String
Dim Handle As Integer
Dim HandleOut As Integer
Dim arLines() As String
Dim i As Long
Dim fSize As Long
Dim Limit As Long
Dim cCounter As Long
Dim FileLOC As Long
Dim ArchPath As String
Dim tFile As String
Dim LineCount As Double
Dim FSO As Object
Dim TSO As Object

tmrTimeOut.Enabled = False
TimeOutLoops = 0


ArchPath = App.Path & "\Incoming\"
Debug.Print strOutFile


tFile = StrInfile
tFile = Replace(tFile, ".txt.txt", ".txt")
StrInfile = Replace(StrInfile, ".txt.txt", ".txt")

If FileExist(tFile) = False Then Exit Sub

HandleOut = FreeFile
Open strOutFile For Output Shared As #HandleOut
  
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TSO = FSO.OpenTextFile(StrInfile)
  Do While Not TSO.AtEndOfStream
      DoEvents
      sText = TSO.ReadLine
      If Left(sText, 1) = "!" Then
        Print #HandleOut, sText & vbCrLf
        LineCount = LineCount + 1
      End If
  Loop
  TSO.Close
Set TSO = Nothing
Set FSO = Nothing

Close HandleOut

Kill StrInfile
Debug.Print strOutFile & "  Total File Linecount = " & CStr(LineCount)



' Send signal to continue

End Sub




Private Sub FinishUp(EndType As Integer, CurrentServer As clsListServer)
Dim Handle As Integer
Dim MyText As String
Dim Record As Long
Dim IPAddr As String
Dim NickName As String

sckDCC.Close
Select Case EndType
Case -1

    SysLog "TIMEOUT: Download of " & NulTrim(fBlock.FileName) & " Failed", vbRed
    LServers(ServerNum).FailCount = LServers(ServerNum).FailCount + 1
    LServers(ServerNum).NextUpDate = Date2Epoch(Now) + 900
    If LServers(ServerNum).FailCount > 4 Then
        LServers(ServerNum).isDisabled = True
    End If
    WaitLoops = 0
    TimeOutLoops = 0
    AmActive = False
    SaveServerINI ServerNum
    ServerPoll
Case 1
    ' Process File List
    DoEvents
    If NulTrim(fBlock.FileName) > "" Then
        ProcessNewList NulTrim(fBlock.FileName), CurrentServer
    End If
    LServers(ServerNum).FailCount = 0
    SaveServerINI ServerNum
    ServerPoll
Case 3
    SysLog "Transfer Never Started"
    LServers(ServerNum).FailCount = LServers(ServerNum).FailCount + 1
    LServers(ServerNum).NextUpDate = Date2Epoch(Now) + 900
    If LServers(ServerNum).FailCount > 4 Then
        LServers(ServerNum).isDisabled = True
    End If
    SaveServerINI ServerNum
    ServerPoll
    
    AmActive = False
    WaitLoops = 0
Case Else
End Select

tmrTimeOut.Enabled = False
fBlock.fHandle = 0
fBlock.FileName = ""
fBlock.FileSize = 0
fBlock.IPAddr = ""
fBlock.PortNum = ""
fBlock.Position = 0
fBlock.TotalBytes = 0
tmrUpdate.Enabled = False
tmrWait.Enabled = False
UpdateDLStatus PB1, 100, -1, ""

End Sub
