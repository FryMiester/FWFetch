VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ctlDownLoad 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   ScaleHeight     =   1050
   ScaleWidth      =   3660
   Begin MSWinsockLib.Winsock sckDCC 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrStartup 
      Enabled         =   0   'False
      Left            =   3120
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
      Left            =   1800
      Top             =   480
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   480
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   480
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
   End
End
Attribute VB_Name = "ctlDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Event NewDownload(BufferStr As String)
Private SearchResults As Boolean
Private SearchFile As String
Private TimerCounter As Integer
Private fBlock As myFileBlock
Private TimeoutTicker As Integer
Private m_MyIndex As Integer
Private m_AmActive As Boolean
Private m_GetListFile As String
Private m_Filename As String
Private m_FileSize As Long
Private m_IP_Info As String
Private m_UserNick As String
Private m_ServerName As String
Private m_IsFileList As Boolean
Dim MD5 As CMD5

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

Public Sub Activate(BuffLine As String)
    UpdateDLStatus PB1, 100, 3, "Connecting", vbGreen
    tmrWait.Enabled = True
    tmrTimeOut.Enabled = True
    ProcessFile BuffLine
End Sub

Private Sub ProcessFile(InString As String)
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
ServerName = NickName

If InStr(UCase(fName), "SEARCH") <> 0 And InStr(UCase(fName), "RESULTS") Then
    ArchPath = MyPathHeader & "\SearchBot_Results\"
    If DirExist(ArchPath) = False Then MkDir ArchPath
    SearchResults = True
Else
    ArchPath = MyDownloadFolder
    SearchResults = False
End If
    
    
    OutputFile = ArchPath & fName
    fBlock.FileName = OutputFile
    fBlock.FileSize = fSize
    fBlock.IPAddr = IPAddr
    fBlock.PortNum = Port

If SearchResults = True Then
    SearchFile = OutputFile
Else
    SearchFile = ""
End If

If FileExist(OutputFile) = True Then
    CurrSize = FileLen(OutputFile)
    If CurrSize > 0 Then
      fBlock.Position = CurrSize
      MyMsg = "DCC RESUME " & QtStr(fName) & " " & Port & " " & CStr(CurrSize)
      IRC_CTCP_MSG NickName, MyMsg
      ' Request Resume
    End If
End If

ReceiveFile InString, NickName


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
                pic.BackColor = bColor
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
FinishUp 4
End Sub

Private Sub sckDCC_Connect()
Dim Index As Integer
Dim Tmp As String

With fBlock
    .fHandle = FreeFile
    Open NulTrim(.FileName) For Binary As .fHandle
    If .Position <> 0 Then
        Seek #.fHandle, .Position + 1
    End If
End With
AmActive = True
tmrWait.Enabled = False
tmrTimeOut.Enabled = False
End Sub

Private Sub sckDCC_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Dim Handle As Integer
Dim Success As Boolean
Static TotalBytes As Long
Dim intPercent As Integer


   'On Error GoTo sckDCC_DataArrival_Error

AmActive = True

Handle = fBlock.fHandle
sckDCC.GetData sData, vbString
TotalBytes = TotalBytes + Len(sData)
InfoLog "Total Bytes : " & CStr(TotalBytes) & " Buffer : " & CStr(Len(sData)), App.Path & "\DCCInfo.txt"

 
Put #Handle, , sData
intPercent = PercentOf(fBlock.FileSize, TotalBytes)
UpdateDLStatus PB1, intPercent, 3, NoPath(NulTrim(fBlock.FileName))

fBlock.TotalBytes = TotalBytes

If TotalBytes >= fBlock.FileSize Then
    TotalBytes = 0
    sckDCC.Close
    Close Handle
    FinishUp 1
End If

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

sckDCC_DataArrival_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure sckDCC_DataArrival of User Control ctlDownLoad", "Debuglog.txt"

End Sub

Public Sub NewPoll()
    tmrPollInfo.Enabled = True
End Sub

Private Sub tmrPollInfo_Timer()
Dim CurrTime As Long
Dim LstCnt As Integer
Dim Rando As Integer
Dim ReqStr As String
Dim Tmp As String
Dim CurrServer As String
Dim Key As String
Dim MD5 As New CMD5



Set MD5 = New CMD5

If AmActive = True Then Exit Sub
If AmConnected = False Then Exit Sub
If AmPaused = True Then Exit Sub

If UseSchedule = True Then
      If INSchedule = False Then
            Exit Sub
      End If
End If



If frmGetMain.lstGetList.ListCount = 0 Then Exit Sub

LstCnt = frmGetMain.lstGetList.ListCount

CurrTime = Date2Epoch(Now)
If LastScanTime + 20 > CurrTime Then
      Exit Sub
End If

LastScanTime = CurrTime

If LstCnt > 0 Then
    'Rando = rand(0, frmGetMain.lstGetList.ListCount - 1)
    ReqStr = frmGetMain.lstGetList.List(0)
    frmGetMain.lstGetList.RemoveItem 0
    
    ' Does File Exists?
    Tmp = ReqStr
    CurrServer = nWord(Tmp, " ")
'    Debug.Print CurrServer & " - " & LastServer
    
    'If ISOnline(frmGetMain.lstUsers, CurrServer) = False Then
    '  frmGetMain.lstOffLine.AddItem ReqStr
    '  Exit Sub
    'End If
    
    LastServer = CurrServer
    If Left(Tmp, 1) = "[" Then
        Garbage = nWord(Tmp, "]")
    ElseIf Left(Tmp, 1) = "%" Then
        Garbage = nWord(Tmp, "%")
        Garbage = nWord(Tmp, "%")
    End If
    
    Tmp = NulTrim(nWord(Tmp, ":"))
    Tmp = NulTrim(MyDownloadFolder) & Tmp
    If FileExist(Tmp) = True Then Exit Sub
    Key = MD5.HaxMD5(Tmp & Date2Epoch(Now))
    ReqString.Add ReqStr, Key
    TimeoutTicker = 0
    tmrPollInfo.Enabled = False
    IsFileList = False
    ' Process request for file
    'PaintScreen "Requesting " & ReqStr, vbGreen
    ProcessRequest ReqStr
End If

End Sub


Private Sub ProcessRequest(FileInfo As String)
Dim Server As String
Dim IPAddr As String
Dim InFileName As String
Dim Tmp As String
Dim fSize As String
Dim Connect As Boolean
Dim OLServer As String
Static LastServer As String

   'On Error GoTo ProcessRequest_Error

Tmp = FileInfo
Server = nWord(Tmp, " ")

OLServer = Right(Server, Len(Server) - 1)
InFileName = nWord(Tmp, ":")

If Left(InFileName, 1) = "[" Then
    Garbage = nWord(InFileName, "]")
ElseIf Left(InFileName, 1) = "%" Then
    Garbage = nWord(InFileName, "%")
    Garbage = nWord(InFileName, "%")
End If

Garbage = nWord(Tmp, " ")
Connect = ISOnline(frmGetMain.lstUsers, OLServer)

'If Connect = False Then
'      Debug.Print OLServer & " Is Offline?"
'   frmGetMain.lstGetList.AddItem FileInfo
'   tmrPollInfo.Enabled = True
'   Exit Sub
'End If
If OLServer = LastServer Then
      Pause 10
End If
LastServer = OLServer

FileName = NulTrim(MyDownloadFolder) & NulTrim(InFileName)

fBlock.FileName = FileName
AmActive = True
tmrTimeOut.Enabled = False
tmrWait.Enabled = True
SysLog "Requesting..." & FileInfo, vbBlue
IRCSay MyChannel, FileInfo

   On Error GoTo 0
   Exit Sub

ProcessRequest_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessRequest of User Control ctlDownLoad", "Debuglog.txt"

End Sub

Private Sub tmrstartup_Timer()
tmrPollInfo.Enabled = True
End Sub

Private Sub tmrTimeOut_Timer()
Dim IntPcnt As Integer
TimeoutTicker = TimeoutTicker + 1
IntPcnt = PercentOf(60, TimeoutTicker)
Select Case TimeoutTicker
Case Is < 15
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting...", vbGreen
Case 16 To 29
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting...", vbCyan
Case 30 To 49
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting...", vbYellow
Case 50 To 60
    UpdateDLStatus PB1, IntPcnt, 3, "Connecting...", vbRed
End Select

If TimeoutTicker = 60 Then
    ' Connection Timed out
    UpdateDLStatus PB1, 100, 3, "TIME OUT", vbRed
    Pause 3
    UpdateDLStatus PB1, 100, -1, ""
    TimeoutTicker = 0
    tmrTimeOut.Enabled = False
    FinishUp -1
End If
End Sub

Private Sub tmrWait_Timer()
Static loops As Integer
Exit Sub
loops = loops + 1
If loops >= 120 Then
    FinishUp 3
End If
End Sub

Private Sub ReceiveFile(FullFileName As String, Optional NickName As String)
Dim fHandle As Integer
Dim IPAddr As String
Dim Port As String

   On Error GoTo ReceiveFile_Error

IPAddr = fBlock.IPAddr
Port = fBlock.PortNum
fHandle = fBlock.fHandle


If IPAddr = "" Or Port = "" Then
    ' Abort Transfer
    FinishUp -1
    Exit Sub
End If

With sckDCC
    If .State <> sckClosed Then
        .Close
    End If
    .Connect NulTrim(fBlock.IPAddr), CStr(fBlock.PortNum)
    TimeoutTicker = 0
    tmrTimeOut.Enabled = True
End With
Pause 5
Do
DoEvents

Loop While sckDCC.State = sckConnected
Close fHandle
'Debug.Print "Handle Closed, what happened"


   On Error GoTo 0
   Exit Sub

ReceiveFile_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ReceiveFile of User Control ctlDownLoad", "Debuglog.txt"

End Sub
Private Sub FinishUp(EndType As Integer)
Dim Handle As Integer
Dim MyText As String

sckDCC.Close
Select Case EndType
Case -1
    SysLog "TIMEOUT: Download of " & NulTrim(fBlock.FileName) & " Failed from " & fBlock.ServerName, vbRed
    BadCount = BadCount + 1
    ' Time Out
        If NulTrim(fBlock.FileName) > "" Then
            GoodCount = GoodCount + 1
            UpdateDLStatus PB1, 100, -1, ""
            SysLog "TIME OUT Connecting For " & NulTrim(fBlock.FileName), vbRed
            MyText = Format(Now, "MM-DD-YYYY  HH:MM ") & " TIMEOUT Failed from " & fBlock.ServerName & " for " & NulTrim(fBlock.FileName)
            Handle = FreeFile
            Open TransferLogFile For Append As Handle
            Print #Handle, MyText
            Close Handle
        End If
    ' log The Transfer
    ' Save for future
Case 1
        If SearchResults = False Then
            If NulTrim(fBlock.FileName) > "" Then
                GoodCount = GoodCount + 1
                UpdateDLStatus PB1, 100, -1, ""
                SysLog "Successful Download of " & NulTrim(fBlock.FileName), vbGreen
                MyText = Format(Now, "MM-DD-YYYY  HH:MM ") & " Successful Download " & NulTrim(fBlock.FileName)
                Handle = FreeFile
                Open TransferLogFile For Append As Handle
                Print #Handle, MyText
                MyText = fBlock.TriggerStr
                MyText = Replace(MyText, vbCrLf, "")
                Print #Handle, MyText
                Close Handle
                frmGetMain.RemoveGETListItem NulTrim(fBlock.FileName), True, ServerName
                FillRequestList
            End If
        Else
            UpdateDLStatus PB1, 100, -1, ""
            Process_Search_File SearchFile, Tree_View
            Exit Sub
        End If
    ' log The Transfer
Case 3
    SysLog "Downloading Error ... ", vbRed
    BadCount = BadCount + 1
        If NulTrim(fBlock.FileName) > "" Then
            GoodCount = GoodCount + 1
            UpdateDLStatus PB1, 100, -1, ""
            SysLog "Downloading ERROR of " & NulTrim(fBlock.FileName), vbRed
            MyText = Format(Now, "MM-DD-YYYY  HH:MM ") & " Downloading ERROR " & NulTrim(fBlock.FileName)
            Handle = FreeFile
            Open TransferLogFile For Append As Handle
            Print #Handle, MyText
            Close Handle
        End If
    
Case 4
    SysLog "User Aborted ... ", vbRed
    BadCount = BadCount + 1
        If NulTrim(fBlock.FileName) > "" Then
            GoodCount = GoodCount + 1
            UpdateDLStatus PB1, 100, -1, ""
            SysLog "User ABORT of " & NulTrim(fBlock.FileName), vbRed
            MyText = Format(Now, "MM-DD-YYYY  HH:MM ") & " USER ABORT " & NulTrim(fBlock.FileName)
            Handle = FreeFile
            Open TransferLogFile For Append As Handle
            Print #Handle, MyText
            Close Handle
        End If
Case Else
End Select

TimeoutTicker = 0
tmrTimeOut.Enabled = False
AmActive = False
fBlock.fHandle = 0
fBlock.FileName = ""
fBlock.FileSize = 0
fBlock.IPAddr = ""
fBlock.PortNum = ""
fBlock.Position = 0
fBlock.TotalBytes = 0
tmrUpdate.Enabled = False
tmrPollInfo.Enabled = True
tmrWait.Enabled = False
UpdateDLStatus PB1, 100, -1, ""

End Sub
