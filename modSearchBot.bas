Attribute VB_Name = "modSearchBot"
Public Const FILE_BEGIN = 0
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Private aUChars(255)        As Byte

Public Type BookInfo
    Title As String
    Series As String
    Author As String
    ISBN As String
End Type



Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
  
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Const INVALID_HANDLE_VALUE = -1
Public Const INVALID_SET_FILE_POINTER = -1

Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Public Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public lngStart As Long

Public Quoted As Boolean

Public Returnfile As String
Public FoundLine() As String
Dim FoundCount As Long
Public SearchLists() As String

Public Function DoSearch(ByRef SearchInfo As SearchHeader) As String
Dim StringToFind As String
Dim SearchTerms As String
Dim Filters As String
Dim UserFilter() As String
Dim Reject As String
Dim L As Long
Dim r As Integer
Dim cCounter As Integer
Dim AA As Long
Dim strFilename As String
Dim ListCount As Integer
Dim strStringToFind As String
Dim OutputName As String
Dim Matched As Boolean
Dim start As Long
Dim Finished As Long
Dim Countup As Integer
Dim Handle As Integer
Dim ReturnHandle As Integer
Dim ReturnStr As String
Dim ZipFile As String
Dim ZipCmnd As String
Dim RtnVal As Long
Dim pPath As String
Dim tPath As String
Dim MaxResults As Integer

   'On Error GoTo DoSearch_Error

ReDim FoundLine(0)
FoundCount = 0
SearchTerms = SearchInfo.SearchTerms
Randomize Timer
Quoted = SearchInfo.Quoted

If NulTrim(SearchTerms) = "" Then Exit Function
Filters = ":/<|>\,*?&" & Chr(34)
For L = 1 To Len(Filters)
    SearchTerms = Replace(SearchTerms, Mid(Filters, L, 1), "")
    SearchTerms = LCase(SearchTerms)
Next

UserFilter = Split(SearchTerms, " ")

For L = 0 To UBound(UserFilter)
    If Left(UserFilter(L), 1) = "-" Then
        Reject = Right(UserFilter(L), Len(UserFilter(L)))
        Exit For
    End If
Next

If Reject > "" Then
    SearchTerms = Replace(SearchTerms, Reject, "")
    Reject = Right(Reject, Len(Reject) - 1)
End If


SysLog "Searching for " & SearchTerms
AddTitleLV frmGetMain.lvTitles, "!Searching For " & SearchTerms

Debug.Print "Searching for " & SearchTerms
strStringToFind = SearchTerms
MaxResults = SearchInfo.MaxResults

frmGetMain.File1.Path = SearchInfo.ListFolder
frmGetMain.File1.Pattern = "*.txt"

ReDim SearchLists(frmGetMain.File1.ListCount)
frmGetMain.File1.Refresh
frmGetMain.lstResults.Clear
ListCount = frmGetMain.File1.ListCount
SysLog "Using " & CStr(ListCount) & " File lists"

Debug.Print "Using " & CStr(ListCount) & " File lists"
Randomize Timer

For L = 0 To ListCount - 1
    SearchLists(L) = frmGetMain.File1.Path & "\" & frmGetMain.File1.List(L)
    
Next

For L = 0 To ListCount - 1
    AA = rand(0, ListCount - 1)
    Swap SearchLists(L), SearchLists(AA)
Next

For L = 0 To ListCount - 1
   DoEvents
   If SearchAbort = True Then GoTo Bottom
    strFilename = SearchLists(L)
    Debug.Print "Using " & strFilename
    Countup = Countup + 1
    AddTitleLV frmGetMain.lvTitles, "!SCANNING " & NoPath(strFilename)
    'ApiReadFile strFilename, strStringToFind
    BlockSearch strFilename, strStringToFind, Quoted, FoundLine
Next
FoundCount = UBound(FoundLine)
If SearchAbort = True Then GoTo Bottom

Debug.Print "DONE " ' & (Finished - start)

'frmSpinner.KillMe

Reject = NulTrim(Reject)
frmGetMain.lvTitles.ListItems.Clear

    If FoundCount > 0 Then
        For L = 0 To FoundCount
        'Debug.Print "-" & Reject & "-"
            If L > MaxResults Then Exit For
            If NulTrim(Reject) > "" Then
                If InStr(UCase(FoundLine(L)), UCase(Reject)) <> 0 Then
                    'AddTitleLV frmGetMain.lvTitles, FoundLine(L)
                Else
                    DoTreeview frmGetMain.TV1, FoundLine(L)
                    AddTitleLV frmGetMain.lvTitles, FoundLine(L)
                    frmGetMain.lstResults.AddItem FoundLine(L)
                End If
            Else
                DoTreeview frmGetMain.TV1, FoundLine(L)
                AddTitleLV frmGetMain.lvTitles, FoundLine(L)
                frmGetMain.lstResults.AddItem FoundLine(L)
            End If
        Next
    Else
        AddTitleLV frmGetMain.lvTitles, "NOTHING FOUND"
    End If


SearchInfo.Completed = True
SearchInfo.FindCount = FoundCount
frmGetMain.MousePointer = vbArrow
    
   On Error GoTo 0
   Exit Function

Bottom:

AddTitleLV frmGetMain.lvTitles, "SEARCH ABORTED"
frmGetMain.MousePointer = vbArrow
frmSpinner.KillMe
Exit Function



DoSearch_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure DoSearch of Module modSearchBot"
    
End Function

Public Sub ApiReadFile(ByVal strFilename As String, ByVal strStringToFind As String)
   Dim hFile As Long, bContent() As Byte
   Dim FileLenght As Long
   Dim Result As Long
   Dim FindPoint As Long
   Dim StartPoint As Long
   Dim EndPoint As Long
   Dim FindLine As String
   Dim SearchPoint As Long
   Dim SearchText As String
   Dim ReadText As String
   Dim cCounter As Long
   Dim lngBlockSize As Long
   Dim lngBlocks As Long
   Dim lngRemaining As Long
   Dim Terms() As String
   Dim Matched As Boolean
   Dim L As Long
    
    
    On Error GoTo ApiReadFile_Error
    StringToFind = LCase(StringToFind)
    
    If Quoted = False Then
        Terms = Split(strStringToFind, " ")
    Else
       ReDim Terms(0)
       Terms(0) = strStringToFind
    End If
    
    lngBlockSize = 2& ^ 20&
    
    
    'DebugLog "Searching " & strFilename
    Debug.Print "Searching " & strFilename
    DoEvents

    hFile = CreateFile(strFilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    FileLenght = GetFileSize(hFile, 0)
    
   If FileLength > lngBlockSize Then
      lngBlocks = FileLength \ lngBlockSize
   Else
      lngBlocks = lngBlockSize
   End If
   lngRemaining = FileLength Mod lngBlockSize
       
    If FileLenght = 0 Then
      Exit Sub
      CloseHandle hFile
    End If
    
    
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    
    ReDim bContent(1 To FileLenght) As Byte
    
    
    
    ReadFile hFile, bContent(1), UBound(bContent), Result, ByVal 0&
    If Result <> UBound(bContent) Then DebugLog "Error reading file ..." & strFilename
    
    CloseHandle hFile
    
    SearchText = LCase(StrConv(bContent, vbUnicode))
    ReadText = StrConv(bContent, vbUnicode)
    SearchPoint = 1
    
    Do
    If SearchAbort = True Then
      CloseHandle hFile
      Exit Sub
    End If
    Matched = True
    FindPoint = InStr(SearchPoint, SearchText, Terms(0))
    If FindPoint >= SearchPoint Then
        StartPoint = InStrRev(SearchText, vbNewLine, FindPoint)
        If StartPoint = 0 Then StartPoint = 1
        EndPoint = InStr(FindPoint + 1, SearchText, vbNewLine)
        FindLine = Mid(ReadText, StartPoint, EndPoint - StartPoint)
        If InStr(FindLine, "!") = 0 Then FindLine = ""
        For L = 1 To UBound(Terms)
            If InStr(LCase(FindLine), Terms(L)) = 0 Then
                Matched = False
                FindLine = ""
                Exit For
            End If
        Next
Skip:
        If FindLine > "" Then
            cCounter = cCounter + 1
            FoundCount = FoundCount + 1
            ReDim Preserve FoundLine(FoundCount)
        End If
    Else
        Exit Do
    End If
    
    If FindLine > "" Then
        FindLine = Replace(FindLine, vbNewLine, "")
        FindLine = Replace(FindLine, vbCr, "")
        FindLine = Replace(FindLine, vbLf, "")
        FoundLine(FoundCount) = FindLine
        'Debug.Print FindLine
    End If
    SearchPoint = EndPoint + 1
    FindPoint = 0
    Loop While SearchPoint < FileLenght
    
    ReDim bContent(0) As Byte
    


   On Error GoTo 0
   Exit Sub

ApiReadFile_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ApiReadFile of Module modSearchBot"
    DebugLog "Error " & strFilename
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ApiReadFile of Module modSearchBot"
    Debug.Print "Error " & strFilename
    
End Sub



Private Sub BlockSearch(InputFile As String, SearchingTerms As String, Quoted As Boolean, Foundlist() As String)

Dim intFilein As Integer
Dim bytData() As Byte
Dim lngBlockSize As Long
Dim lngLength As Long
Dim lngSeek As Long
Dim lngBlocks As Long
Dim lngRemaining As Long
Dim lngJ As Long
Dim lngI As Long
Dim lngLast As Long
Dim lngPos As Long
Dim strSearch As String
Dim strData As String
Dim HoldData As String
Dim StrHold As String
Dim strRecord As String
Dim boEnd As Boolean
Dim boFinished As Boolean
Dim intCount As Integer
Dim Matched As Boolean
Dim L As Long
Dim tTerms() As String
Dim Terms As String
Dim TermCount As Integer

ReDim tTerms(0)

lngBlockSize = 2& ^ 23&
intFilein = FreeFile


Terms = LCase(SearchingTerms)
If Quoted = True Then
   tTerms(0) = Terms
Else
   tTerms = Split(Terms, " ")
End If

TermCount = UBound(tTerms)


strSearch = tTerms(0)

'lngReturn = QueryPerformanceCounter(curThen)

intFilein = FreeFile
Open InputFile For Binary As intFilein

'intFileOut = FreeFile
'Open App.Path & "\LargeSearch.txt" For Output As intFileOut

lngLength = LOF(intFilein)
If lngLength > lngBlockSize Then
   lngBlocks = lngLength \ lngBlockSize
Else
   lngBlocks = lngBlockSize
End If
lngRemaining = lngLength Mod lngBlockSize
ReDim bytData(lngBlockSize - 1)
lngSeek = 1
Do
    If lngBlocks = 0 Then
        '
        ' We've processed all the complete blocks
        ' check if there are any remaining bytes
        ' if ao then set up the Buffer to the correct length
        ' otherwise signal to exit
        '
        If lngRemaining <> 0 Then
            ReDim bytData(lngRemaining - 1)
        Else
            boEnd = True
        End If
    End If
    If boEnd = False Then
        Get intFilein, lngSeek, bytData
        strData = LCase(StrConv(bytData, vbUnicode))
        HoldData = StrConv(bytData, vbUnicode)
        
        '
        ' Find the last Line Feed character
        ' and truncate the string to that position
        ' set the position for the next Get (lngSeek)
        ' to the position after the last Line Feed
        '
        lngLast = InStrRev(strData, vbLf, Len(strData))
        lngSeek = lngSeek + lngLast
        strData = Left$(strData, lngLast)
        lngJ = 1
        Do
            lngPos = InStr(lngJ, strData, strSearch)
            If lngPos > 0 Then
                '
                ' Found what we were looking for
                ' Find the start of this record (lngI)
                ' and the end of this record (lngJ)
                ' extract the entire record and write it
                ' to the results file
                '
                lngI = InStrRev(strData, vbLf, lngPos) + 1
                If lngI = 0 Then lngI = 1
                lngJ = InStr(lngPos, strData, vbLf)
                If lngJ = 0 Then lngJ = Len(strData) + 1
                strRecord = Mid$(strData, lngI, lngJ - lngI)
                StrHold = Mid$(HoldData, lngI, lngJ - lngI)
                Select Case TermCount
                Case 0
                     If Left(strRecord, 1) = "!" Then
                        AddToSTRArray Foundlist, StrHold
                        'Print #intFileOut, strRecord
                        intCount = intCount + 1
                     End If
                Case Else
                     Matched = True
                       For L = 0 To TermCount
                          If InStr(strRecord, tTerms(L)) = 0 Then
                             Matched = False
                          End If
                      Next
                      If Matched = True Then
                          If Left(strRecord, 1) = "!" Then
                              AddToSTRArray Foundlist, StrHold
                              'Print #intFileOut, strRecord
                              intCount = intCount + 1
                          End If
                      End If
                 End Select
                '
                ' Set the start point for the next search
                '
                lngJ = lngJ + 2
            Else
                boFinished = True
            End If
        Loop Until boFinished = True
        '
        ' Decrement the number of full blocks
        ' re-set the 'block completed' flag
        ' and loop until we've processed the complete file
        '
        lngBlocks = lngBlocks - 1
        boFinished = False
    End If
Loop Until boEnd = True Or lngSeek > lngLength
Close intFilein
ReDim bytData(0)
End Sub

Private Sub AddToSTRArray(ByRef strArray() As String, NewToAdd As String)
Dim cCount As Integer
cCount = UBound(strArray)
cCount = cCount + 1
ReDim Preserve strArray(cCount)
strArray(cCount) = NewToAdd

End Sub

