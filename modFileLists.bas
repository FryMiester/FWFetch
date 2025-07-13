Attribute VB_Name = "modFileLists"
Option Explicit
Public XfrTimer As Integer



Public Sub Process_Search_File(InlistName As String, Optional UseTree As Boolean)
Dim Ext As String
Dim CommandLine As String
Dim ArchPath As String
Dim Handle As Integer
Dim InString As String
Dim RtnVal As Long
Dim Pattern() As String
Dim Tmp As String
Dim NewName As String
Dim Ending As String
Dim L As Long
Dim Archive As String
Dim FileMask As String
Dim DirMask As String
Dim LogName As String

   'On Error GoTo Process_Search_File_Error
   
Archive = InlistName
Tmp = InlistName
ArchPath = RunPath(InlistName)
DirMask = ArchPath
FileMask = "search*.*"
frmGetMain.lvTitles.ListItems.Clear
frmGetMain.TV1.Nodes.Clear
frmGetMain.TV1.Nodes.Add , , "root", "Search"

If LCase(FileExten(InlistName)) = ".zip" Then
   CommandLine = App.Path & "\7za.exe" & " e -y " & Chr(34) & Archive & Chr(34)
   ChDir ArchPath
   RtnVal = Shell(CommandLine, vbHide)
   ChDir App.Path
   Pause 1
   Kill Archive
End If
frmGetMain.File1.Path = ArchPath
frmGetMain.File1.Pattern = "*.txt"
frmGetMain.File1.Refresh
NewName = ArchPath & frmGetMain.File1.List(0)
Debug.Print NewName

Handle = FreeFile
Open NewName For Input As Handle

Do Until EOF(Handle)
   DoEvents
    Line Input #Handle, InString
    If Left(InString, 1) = "!" Then
        AddTitleLV frmGetMain.lvTitles, InString
        DoTreeview frmGetMain.TV1, InString
    End If
Loop
Close Handle
LogName = MySearchResultFolder & Format(Now, "YYYYMMDD-hhmmss") & "-" & NoPath(NewName)

Name NewName As LogName

For L = 0 To 5
    frmGetMain.cmdFrame(L).BackColor = &H8000000F
    frmGetMain.FrameMain(L).Visible = False
Next
frmGetMain.cmdFrame(1).BackColor = &H80000016
frmGetMain.FrameMain(1).Visible = True
frmGetMain.LoadCBO



   On Error GoTo 0
   Exit Sub

Process_Search_File_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure Process_Search_File of Module modFileLists", "Debuglog.txt"
End Sub

Public Sub WriteIgnores()
Dim IgnoreFile As String
Dim Cnt As Integer
Dim Handle As Integer
Dim L As Long
Cnt = frmGetMain.lstIgnores.ListCount
If Cnt = 0 Then Exit Sub

IgnoreFile = MyAdditionalFolder & "Ignores.txt"
Handle = FreeFile
Open IgnoreFile For Output As #Handle
For L = 0 To Cnt - 1
    Print #Handle, frmGetMain.lstIgnores.List(L)
Next
Close Handle

End Sub

Public Sub ReadIgnores()
Dim IgnoreFile As String
Dim Handle As Integer
Dim Tmp As String

IgnoreFile = MyAdditionalFolder & "Ignores.txt"
frmGetMain.lstIgnores.Clear
If FileExist(IgnoreFile) = False Then Exit Sub
Handle = FreeFile
Open IgnoreFile For Input As #Handle
Do While Not EOF(Handle)
    Line Input #Handle, Tmp
    frmGetMain.lstIgnores.AddItem Tmp
Loop
Close #Handle

End Sub
