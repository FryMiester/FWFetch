VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Type ListServerInfo
'    IPAddress       As String * 75
'    NickName        As String * 48
'    HasSlots        As Boolean
'    FilesSize       As Double
'    FilesCount      As Long
'    ListDate        As Long
'    isDisabled      As Boolean
'    ListNeeded      As Boolean
'    IsOnline        As Boolean
'    Deleted         As Boolean
'End Type


Public IPAddress As String
Public SVRHash As String
Public LSNickName As String
Public NextUpDate As Long
Public HasSlots As Boolean
Public FilesSize As Double
Public FilesCount As Long
Public ListDate As Long
Public isDisabled As Boolean
Public ListNeeded As Boolean
Public ISOnline As Boolean
Public FailCount As Long
Public Deleted As Boolean
Public RecordNum As Long

