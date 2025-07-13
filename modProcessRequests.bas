Attribute VB_Name = "modProcessRequests"
Option Explicit

' User Requests Files .... Add to INI File

Public Sub StoreRequest(ReqStr As String)
Dim ReqININame As String
Dim Header As String
Dim ServerName As String
Dim REQHead As String
Dim SVRHead As String
Dim ExistREQ As String
Dim TempStr As String
Dim SVRName As String
Dim L As Long
Dim cCount As Integer
Dim Matched As Boolean
Dim SvrCount As Integer

ReqININame = MyServersFolder & "Requests.ini"

TempStr = ReqStr
ServerName = nWord(TempStr, " ")
ServerName = Right(ServerName, Len(ServerName) - 1)
SvrCount = CInt(Val(ReadINI("SERVERLIST", "SERVERCOUNT", ReqININame)))
For L = 1 To SvrCount
   SVRHead = "SVR_" & Format(L, "00#")
   SVRName = ReadINI("SERVERLIST", SVRHead, ReqININame)
   If ServerName = SVRName Then
      Matched = True
      Exit For
   End If
Next

If Matched = False Then
   SvrCount = SvrCount + 1
   SVRHead = "SVR_" & Format(SvrCount, "00#")
   WriteINI "SERVERLIST", SVRHead, ServerName, ReqININame
   WriteINI "SERVERLIST", "SERVERCOUNT", CStr(SvrCount), ReqININame
End If
Matched = False

cCount = CInt(Val(ReadINI(ServerName, "REQUESTCOUNT", ReqININame)))

If cCount = 0 Then
   REQHead = "REQ_001"
   WriteINI ServerName, "REQUESTCOUNT", "1", ReqININame
   WriteINI ServerName, REQHead, ReqStr, ReqININame
   Exit Sub
End If

For L = 1 To cCount
   REQHead = "REQ_" & Format(L, "00#")
   If UCase(ReqStr) = UCase(ReadINI(ServerName, REQHead, ReqININame)) Then
      MsgBox ReqStr & vbCrLf & " is already requested from " & ServerName, 3
      Matched = True
      Exit Sub
   End If
Next
   
For L = 1 To cCount
   REQHead = "REQ_" & Format(L, "00#")
   If NulTrim(ReadINI(ServerName, REQHead, ReqININame)) = "" Then
      Matched = True
      WriteINI ServerName, REQHead, ReqStr, ReqININame
      Exit Sub
   End If
Next
      
cCount = cCount + 1
REQHead = "REQ_" & Format(cCount, "00#")
WriteINI ServerName, "REQUESTCOUNT", CStr(cCount), ReqININame
WriteINI ServerName, REQHead, ReqStr, ReqININame

End Sub

Public Sub FillRequestList()
Dim ReqININame As String
Dim Header As String
Dim ServerName As String
Dim REQHead As String
Dim sREQHead As String
Dim ExistREQ As String
Dim TempStr As String
Dim L As Long
Dim a As Long
Dim b As Long
Dim cCount As Integer
Dim sCount As Integer
Dim Matched As Boolean

ReqININame = MyServersFolder & "Requests.ini"
Header = "SERVERLIST"
cCount = CInt(Val(ReadINI(Header, "SERVERCOUNT", ReqININame)))
If cCount = 0 Then Exit Sub
frmGetMain.lstGetList.Clear

For L = 1 To cCount
   sCount = 0
   REQHead = "SVR_" & Format(L, "00#")
   ServerName = ReadINI(Header, REQHead, ReqININame)
   sCount = CInt(Val(ReadINI(ServerName, "REQUESTCOUNT", ReqININame)))
   If sCount > 0 Then
      For a = 1 To sCount
         sREQHead = "REQ_" & Format(a, "00#")
         TempStr = ReadINI(ServerName, sREQHead, ReqININame)
         Debug.Print TempStr
         If TempStr > "" Then
            frmGetMain.lstGetList.AddItem TempStr
         End If
      Next
   End If
Next
End Sub

