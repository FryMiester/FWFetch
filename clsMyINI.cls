VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type KeyVals
    KeyStr As String
    ValStr As String
End Type

Private Type INI
    HeaderStr As String
    keyVal(100) As KeyVals
End Type




Public Sub WriteMyINI(Header As String, KeyName As String, Value As String, FileName As String)
Dim fHandle As Integer
Dim fName As String
Dim NumHeaders As Integer
Dim NumKeys As Integer
Dim myINI() As INI
Dim Headers() As String, Keys() As String, TheValue() As String

Dim Portions() As String
Dim Testit As String
Dim sText As String
Dim Handle As Integer
Dim L As Long
Dim Z As Long
Dim LineCount As Integer
fName = FileName

If FileExist(fName) = False Then Exit Sub
fHandle = FreeFile
Open fName For Input As #fHandle
sText = Input(LOF(fHandle), #fHandle)
Close #fHandle
Portions() = Split(sText, vbCrLf)
sText = vbNullString
LineCount = UBound(Portions)
For L = 0 To LineCount - 1
    Portions(L) = NulTrim(Portions(L))
Next
' find Header Section
NumHeaders = -1
For L = 0 To LineCount - 1
    sText = Portions(L)
    If Left(Portions(L), 1) = "[" And Right(Portions(L), 1) = "]" Then
        NumHeaders = NumHeaders + 1
        ReDim Preserve myINI(NumHeaders)
        myINI(NumHeaders).HeaderStr = Portions(L)
        NumKeys = 0
        Debug.Print myINI(NumHeaders).HeaderStr
        Debug.Print myINI(0).HeaderStr
    End If
    
    If InStr(Portions(L), "=") <> 0 Then
        myINI(NumHeaders).keyVal(NumKeys).KeyStr = NulTrim(nWord(sText, "="))
        myINI(NumHeaders).keyVal(NumKeys).ValStr = NulTrim(sText)
        If UCase(myINI(NumHeaders).HeaderStr) = "[" & UCase(Header) & "]" Then
            If UCase(myINI(NumHeaders).keyVal(NumKeys).KeyStr) = UCase(KeyName) Then
                myINI(NumHeaders).keyVal(NumKeys).ValStr = Value
            End If
        End If
        NumKeys = NumKeys + 1
    End If
Next

Debug.Print NumHeaders
fHandle = FreeFile
Open fName & ".bak" For Output As #fHandle
For L = 0 To NumHeaders
    Print #fHandle, "-"
    Print #fHandle, NulTrim(myINI(L).HeaderStr)
    For Z = 0 To 100
        sText = NulTrim(myINI(L).keyVal(Z).KeyStr)
        If sText > "" Then
            sText = sText & "=" & NulTrim(myINI(L).keyVal(Z).ValStr)
        End If
        If sText > "=" Then Print #fHandle, sText
    Next
Next
Close fHandle
Debug.Print myINI(0).HeaderStr
End Sub

Public Function ReadMyINI(ByVal Header As String, ByVal Topic As String, ByVal fName As String) As String
Dim Portions() As String
Dim Testit As String
Dim sText As String
Dim Handle As Integer
Dim L As Long
Dim LineCount As Integer

Handle = FreeFile
Open fName For Input As #Handle
sText = Input(LOF(Handle), #Handle)
Close #Handle
Portions() = Split(sText, vbCrLf)
sText = vbNullString
LineCount = UBound(Portions)
For L = 0 To LineCount - 1
    sText = Portions(L)
    Testit = UCase(nWord(sText, "="))
    If Testit = UCase(Topic) Then
        ReadMyINI = NulTrim(sText)
        Exit Function
    End If
Next
ReadMyINI = ""
End Function

