Attribute VB_Name = "modListViewWork"



Public Sub BuildBookListview(LLView As ListView)

   On Error GoTo BuildMyListview_Error

LLView.ColumnHeaders.Clear
LLView.View = lvwReport

If TerminateMe = True Then Exit Sub

With LLView.ColumnHeaders
    
    .Add , , "Server", 1500
    .Add , , "Title", 7000
    .Add , , "Ext.", 1025
    .Add , , "Size", 1225
    .Add , , "Trigger", 0
    .Add , , "OSize", 0
End With
LLView.View = lvwReport

   On Error GoTo 0
   Exit Sub

BuildMyListview_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure BuildMyListview of Module modListViewWork"

End Sub

Public Sub AddTitleLV(LLView As ListView, BookString As String)
    Dim IDX As Integer
    Dim r As Integer
    Dim cCount As Integer
    Dim Title As String
    Dim Exten As String
    Dim fSize As String
    Dim OSize As String
    Dim Tmp As String
    Dim Filter As String
    Dim Server As String
    Dim OnLine As Boolean
    Dim ttmp As String
    
    Tmp = BookString
    ttmp = BookString
    If Tmp = "" Then Exit Sub
    Server = nWord(ttmp, " ")
    Server = Right(Server, Len(Server) - 1)
    'Garbage = nWord(tmp, " ")
    Title = NulTrim(nWord(Tmp, ":"))
    Exten = Right(Title, Len(Title) - Len(noExt(Title)))
    Garbage = nWord(Tmp, ":")
    Garbage = nWord(Tmp, ":")
    Garbage = nWord(Tmp, ":")
    If InStr(Tmp, ":") <> 0 Then
        Garbage = nWord(Tmp, ":")
        Tmp = Garbage
    End If
    If Title = "" Then Exit Sub
    
    fSize = NulTrim(Tmp)
    
    If fSize = "" Then fSize = "Unknown"
    OSize = OSizeUp(fSize)
    OnLine = ISOnline(frmGetMain.lstUsers, Server)
    cCount = LLView.ListItems.Count + 1
    IDX = cCount ' cCount
    If BookString = "NOTHING FOUND" Then
        Server = " "
        Title = BookString
    End If
    
    Set lvr = LLView.ListItems.Add(IDX, , Server)
        If OnLine = False Then
            LLView.ListItems.Item(IDX).ForeColor = vbRed
        Else
            LLView.ListItems.Item(IDX).ForeColor = vbGreen
        End If
        LLView.ListItems.Item(IDX).ListSubItems.Add 1, , Title
        LLView.ListItems.Item(IDX).ListSubItems.Add 2, , Exten
        LLView.ListItems.Item(IDX).ListSubItems.Add 3, , fSize
        LLView.ListItems.Item(IDX).ListSubItems.Add 4, , BookString
        LLView.ListItems.Item(IDX).ListSubItems.Add 5, , OSize
   On Error GoTo 0
   Exit Sub

AddLV_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure AddLV of Module modListViewWork"
End Sub




Public Sub BuildMyListview(LLView As ListView)

   On Error GoTo BuildMyListview_Error

LLView.ColumnHeaders.Clear
LLView.View = lvwReport

If TerminateMe = True Then Exit Sub

With LLView.ColumnHeaders
    .Add , , "ServerName", 1500
    .Add , , "LN", 375
    .Add , , "OL", 375
    .Add , , "RecordNum", 0
    .Add , , "Last UpDate", 1800
    .Add , , "Next UpDate", 1800
    .Add , , "Disabled", 975
    .Add , , "Fails", 675
    .Add , , "IP Address", 0
End With
LLView.View = lvwReport

   On Error GoTo 0
   Exit Sub

BuildMyListview_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure BuildMyListview of Module modListViewWork"

End Sub


Public Sub AddLV(LLView As ListView, User As ListServerInfo)
    Dim IDX As Integer
    Dim cCount As Integer
    Dim CurrColor As Long
    
   'On Error GoTo AddLV_Error

    cCount = LLView.ListItems.Count + 1
    If inPosition = 0 Then
        IDX = cCount ' cCount
    Else
        IDX = inPosition
    End If
    
    Set lvr = LLView.ListItems.Add(IDX, , NulTrim(User.NickName))
        If User.ISOnline = False Then CurrColor = QBColor(7)
        If User.isDisabled = True Then CurrColor = vbMagenta
        LLView.ListItems(IDX).ForeColor = CurrColor
        If User.ListNeeded = True Then
            LLView.ListItems.Item(IDX).ListSubItems.Add 1, , "§"
            LLView.ListItems.Item(IDX).ListSubItems.Item(1).ForeColor = vbRed
        Else
            LLView.ListItems.Item(IDX).ListSubItems.Add 1, , " "
        End If
        If User.ISOnline = True Then
            LLView.ListItems.Item(IDX).ListSubItems.Add 2, , "Ø"
            LLView.ListItems.Item(IDX).ListSubItems.Item(2).ForeColor = vbGreen
        Else
            LLView.ListItems.Item(IDX).ListSubItems.Add 2, , "§"
            LLView.ListItems.Item(IDX).ListSubItems.Item(2).ForeColor = vbRed
        End If
        LLView.ListItems.Item(IDX).ListSubItems.Add 3, , CStr(User.RecordNum)
        LLView.ListItems.Item(IDX).ListSubItems.Add 4, , Format(Epoch2Date(User.ListDate), "MM/DD/YY - hh:mm") '   CStr(User.RecordNum)
        LLView.ListItems.Item(IDX).ListSubItems.Add 5, , Format(Epoch2Date(User.NextUpDate), "MM/DD/YY - hh:mm")  '   CStr(User.RecordNum)
        LLView.ListItems.Item(IDX).ListSubItems.Add 6, , Bool2Word(User.isDisabled)
        LLView.ListItems.Item(IDX).ListSubItems.Add 7, , CStr(User.FailCount)
        LLView.ListItems.Item(IDX).ListSubItems.Add 8, , NulTrim(User.IPAddress)
        LLView.ListItems.Item(IDX).ListSubItems.Item(6).ForeColor = CurrColor
   On Error GoTo 0
   Exit Sub

AddLV_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure AddLV of Module modListViewWork"
End Sub


Public Sub DoTreeview(TV1 As TreeView, InputString As String)
Dim IDX As Integer
    Dim r As Integer
    Dim Title As String
    Dim IsThere As Boolean
    Dim Author As String
    Dim Tmp As String
    Dim Server As String
    Dim ttmp As String
'    Static AuthorName As String
'    Static AuthorNameKey As String
'    Static ServerName As String
'    Static ServerNameKey As String
    
    Dim MD5 As CMD5
    Set MD5 = New CMD5
    
    
    BookString = InputString
    
    Tmp = BookString
    ttmp = BookString
    If Tmp = "" Then Exit Sub
    Server = nWord(ttmp, " ")
    Server = Right(Server, Len(Server) - 1)
    Title = NulTrim(nWord(Tmp, ":"))
    If InStr(Title, "%") <> 0 Then
        Garbage = nWord(Title, "%")
        Garbage = nWord(Title, "%")
    End If
    ttmp = Title
    Author = NulTrim(nWord(ttmp, "-"))
    
    ServerNameKey = MD5.HaxMD5(Server)
    ServerName = Server
    AuthorNameKey = MD5.HaxMD5(Server & Author)
    AuthorName = Author
    
       Dim nodX As Node
    
    On Error Resume Next
    
    If Title = "" Then Exit Sub
    If TV1.Nodes.Count > 63000 Then Exit Sub
    
            Set nodX = TV1.Nodes.Add("root", tvwChild, ServerNameKey, Server)
            TV1.Nodes.Add ServerNameKey, tvwChild, AuthorNameKey, Author
            TV1.Nodes.Add AuthorNameKey, tvwChild, Title, BookString
            nodX.EnsureVisible



End Sub

Public Function NodeExists(ByVal strKey As String) As Boolean
    Dim Node As MSComctlLib.Node
    On Error Resume Next
    Set Node = frmGetMain.TV1.Nodes(strKey)
    Select Case Err.Number
        Case 0
            NodeExists = True
        Case Else
            NodeExists = False
    End Select
End Function

Private Function OSizeUp(ByVal SmallSize As String) As String
Dim Sz As String
Dim BigNum As Double
Dim Temp As String

Temp = "000000000000000000"

If SmallSize = "" Then
      OSizeUp = "0"
      Exit Function
End If
Sz = Right(SmallSize, 2)
SmallSize = NulTrim(Left(SmallSize, Len(SmallSize) - 2))
Select Case Sz
Case "KB"
      BigNum = CDbl(Val(SmallSize)) * 1024
      '1024
Case "MB"
      BigNum = CDbl(Val(SmallSize)) * 1048576
Case "GB"
      BigNum = CDbl(Val(SmallSize)) * 1073741824
End Select
Temp = Temp & CStr(BigNum)
OSizeUp = Right(Temp, 18)
End Function
