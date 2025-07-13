Attribute VB_Name = "modMisc"
Option Explicit

Public Function longIP2Dotted(ByVal LongIP As String) As String
    Dim i As Long, num As Currency
   On Error GoTo longIP2Dotted_Error
'    If InStr(LongIP, "::") > 0 Then
'        longIP2Dotted = LongIP
'        Exit Function
'    End If
    
    For i = 1 To 4
        num = Int(LongIP / 256 ^ (4 - i))
        LongIP = LongIP - (num * 256 ^ (4 - i))
        If num > 255 Then Err.Raise vbObjectError + 1
        If i = 1 Then
            longIP2Dotted = num
        Else
            longIP2Dotted = longIP2Dotted & "." & num
        End If
    Next

   On Error GoTo 0
   Exit Function

longIP2Dotted_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure longIP2Dotted of Module modMisc"
End Function


Public Function myLongIp(StaticIP As String) As Double

    Dim retVal As Currency, myIPSegments() As String, i As Long
   On Error GoTo myLongIp_Error

    myIPSegments = Split(StaticIP, ".")    'Split(Form1.sckServer.LocalIP, ".")
    myIPSegments = invertArray(myIPSegments)
    For i = 0 To 3
        retVal = retVal + (myIPSegments(i) * (256 ^ i))
    Next
    myLongIp = CDbl(retVal)

   On Error GoTo 0
   Exit Function

myLongIp_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure myLongIp of Module modMisc"
End Function

Public Function invertArray(ByRef arrayName() As String) As String()
    Dim i As Long, tempArr() As String: ReDim tempArr(UBound(arrayName))
   On Error GoTo invertArray_Error

    For i = 0 To UBound(arrayName)
        tempArr(UBound(arrayName) - i) = arrayName(i)
    Next
    invertArray = tempArr

   On Error GoTo 0
   Exit Function

invertArray_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure invertArray of Module modMisc"
End Function


Public Function NoExtFile(ByVal strFile As String) As String
    NoExtFile = Mid(strFile, InStrRev(strFile, "\") + 1)
End Function


Public Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
   On Error GoTo lngMIN_Error

    If L1 < L2 Then
        lngMIN = L1
    Else
        lngMIN = L2
    End If

   On Error GoTo 0
   Exit Function

lngMIN_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure lngMIN of Module modMisc"
End Function

Public Sub ColorText(TextString As String, Optional txtcolor As Long)

    Dim LogFile As String
    Dim BackString As String
   On Error GoTo ColorText_Error

    BackString = TextString
    If txtcolor = 0 Then txtcolor = vbBlack
    'TextString = Replace(TextString, vbNewLine, "")
    frmGetMain.RTBPrivate.SelStart = Len(frmGetMain.RTBPrivate.Text)
    frmGetMain.RTBPrivate.SelColor = txtcolor
    frmGetMain.RTBPrivate.SelText = frmGetMain.RTBPrivate.SelText & TextString
    
    TextString = Replace(TextString, vbCrLf, "")
    'LogFile = App.Path & "\" & ServerChannel & ".log"
    
   On Error GoTo 0
   Exit Sub

ColorText_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ColorText of Module modMisc"
End Sub

Public Sub ChatColorText(TextString As String, Optional txtcolor As Long)

    Dim LogFile As String
    Dim BackString As String
   On Error GoTo ColorText_Error

    BackString = TextString
    If txtcolor = 0 Then txtcolor = vbBlack
    'TextString = Replace(TextString, vbNewLine, "")
    frmGetMain.RTBChat.SelStart = Len(frmGetMain.RTBLog.Text)
    frmGetMain.RTBChat.SelColor = txtcolor
    frmGetMain.RTBChat.SelText = frmGetMain.RTBLog.SelText & TextString
    
    TextString = Replace(TextString, vbCrLf, "")
    'LogFile = App.Path & "\" & ServerChannel & ".log"
    
   On Error GoTo 0
   Exit Sub

ColorText_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ChatColorText of Module modMisc"
End Sub

Public Sub PaintScreen(TextString As String, Optional txtcolor As Long, Optional FontSize As Long)
    
   On Error GoTo PaintScreen_Error

    TextString = Replace(TextString, vbCrLf, "")
    If txtcolor = 0 Then txtcolor = vbBlack
    
    frmGetMain.RTBChat.SelStart = Len(frmGetMain.RTBChat.Text)
    frmGetMain.RTBChat.SelColor = txtcolor
    frmGetMain.RTBChat.SelText = vbCrLf & TextString
    If FontSize <> 0 Then
        frmGetMain.RTBChat.SelFontSize = FontSize
    Else
        frmGetMain.RTBChat.SelFontSize = 12
    End If
    frmGetMain.RTBChat.SelStart = Len(frmGetMain.RTBChat.Text)

   On Error GoTo 0
   Exit Sub

PaintScreen_Error:

    'DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure PaintScreen of Module modMisc"
    
End Sub


Public Sub UpdateDLStatus(pic As PictureBox, ByVal sngPercent As Single, Optional DisplayMode As Integer, Optional strText As String)
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

UpdateDLStatus_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateDLStatus of User Control DCCUnit"

 End Sub

Public Sub ProcessDCCSend(InputStr As String, NickName As String, FileName As String, IPAddress As String, Port As String, FileSize As Long)
Dim Garbage As String
Dim ProcessLine As String
Dim dData() As String
Dim intStart As Integer
Dim intStop As Integer
Dim L As Long
   On Error GoTo ProcessDCCSend_Error

' ""

ProcessLine = Right(InputStr, Len(InputStr) - 1)
NickName = nWord(ProcessLine, "!")
Garbage = nWord(ProcessLine, "")

Garbage = nWord(ProcessLine, " ")
Garbage = nWord(ProcessLine, " ")

intStart = InStr(ProcessLine, Chr(34))
If intStart <> 0 Then
    ProcessLine = Right(ProcessLine, Len(ProcessLine) - 1)
    FileName = nWord(ProcessLine, Chr(34))
Else
    FileName = nWord(ProcessLine, Chr(32))
End If
ProcessLine = NulTrim(ProcessLine)
dData = Split(ProcessLine, " ")


Select Case UBound(dData)
Case 2
    IPAddress = longIP2Dotted(dData(0))
    Port = dData(1)
    dData(2) = Replace(dData(2), Chr(1), "")
    FileSize = CLng(Val(dData(2)))
Case 3
    IPAddress = longIP2Dotted(dData(0))
    Port = dData(2)
    dData(3) = Replace(dData(3), Chr(1), "")
    FileSize = CLng(Val(dData(3)))
Case Else
End Select
' DebugLog IPAddress

   On Error GoTo 0
   Exit Sub

ProcessDCCSend_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessDCCSend of Form frmMain"

End Sub

Public Function QuickRead(fName As String) As String
    Dim i As Integer
    Dim res As String
    Dim L As Long

    i = FreeFile
    L = FileLen(fName)
    res = Space(L)
    Open fName For Binary Access Read As #i
    Get #i, , res
    Close i
    QuickRead = res
End Function

Public Sub GetSearchBots(InString As String)
Dim Cnt As Integer
':Search!Search@ihw-8bf.2bt.25.50.IP PRIVMSG AstroDawg :TRIGGER irchighway #ebooks @search
Garbage = nWord(InString, "!")
Garbage = nWord(InString, ":")
Garbage = nWord(InString, "#")
Garbage = nWord(InString, " ")
InString = Replace(InString, "", "")
InString = Replace(InString, vbCrLf, "")

Cnt = SearchBots.Count
If Cnt = 1 Then
   frmGetMain.cmdIRCSearch.Caption = SearchBots.Item(1)
   Exit Sub
End If


If Cnt < 6 Then
      If Left(InString, 1) = "@" Then
            SearchBots.Add InString
            frmGetMain.cmdIRCSearch.Caption = SearchBots.Item(1)
      End If
      Exit Sub
End If

End Sub


Public Function StripColor(InputTxt As String) As String
' :zathras.mo.us.irchighway.net 332 FWServer #ebooks :4You need DCC, so browser-based IRC clients won't work - Look for books with 1@Search or 1@Searchook - No textbooks here, say 1@TEXTBOOKS for secret site - New books: 1!trainpacks or 1@Oatmeal - Send new books to Oatmeal - 3Using 6m4IR7C3? Type 1@sbclient - 4Don't test scripts here!
Dim Rtn As String
Dim Params() As String
Dim ParamCount As Integer
Dim Backup As String
Dim L As Long
Dim A As Integer
Dim Char As String

Backup = InputTxt
Params = Split(InputTxt, "")
ParamCount = UBound(Params)
For L = 0 To ParamCount
      Params(L) = Replace(Params(L), "", "")
      For A = 0 To ParamCount
            If IsNumeric(Left(Params(L), 1)) = True Then
                  Params(L) = Right(Params(L), Len(Params(L)) - 1)
            End If
            If Left(Params(L), 1) = "," Then
                  Params(L) = Right(Params(L), Len(Params(L)) - 1)
            End If
      Next A
      Rtn = Rtn & Params(L)
Next L
For L = 0 To 31
      Char = Chr(L)
      Rtn = Replace(Rtn, Char, "")
Next
StripColor = NulTrim(Rtn)

End Function

Public Function ConvertChars(InputString As String) As String
' Catch incoming multi-byte characters and convert them
Dim GoodList() As String
Dim BadList() As String
Dim ListGood As String
Dim ListBad As String
Dim L As Long
Dim Upper As Long
Dim OutLine As String
   On Error GoTo ConvertChars_Error

OutLine = InputString

ListGood = "0128|0130|0131|0132|0133|0134|0135|0136|0137|0138|0139|0140|0142|0145|0146|0147|0148|0149|0150|0151|0152|0153|0154|0155|0156|0158|0159|0161|0162|0163|0164|0165|0166|0167|0168|0169|0170|0171|0172|0173|0174|0175|0176|0177|0178|0179|0180|0181|0182|0183|0184|0185|0186|0187|0188|0189|0190|0191|0192|0193|0194|0195|0196|0197|0198|0199|0200|0201|0202|0203|0204|0205|0206|0207|0208|0209|0210|0211|0212|0213|0214|0215|0216|0217|0218|0219|0220|0221|0222|0223|0224|0225|0226|0227|0228|0229|0230|0231|0232|0233|0234|0235|0236|0237|0238|0239|0240|0241|0242|0243|0244|0245|0246|0247|0248|0249|0250|0251|0252|0253|0254|0255"
ListBad = "€|‚|ƒ|„|…|†|‡|ˆ|‰|Š|‹|Œ|Ž|‘|’|“|”|•|–|—|˜|™|š|›|œ|ž|Ÿ|¡|¢|£|¤|¥|¦|§|¨|©|ª|«|¬|­|®|¯|°|±|²|³|´|µ|¶|·|¸|¹|º|»|¼|½|¾|¿|À|Á|Â|Ã|Ä|Å|Æ|Ç|È|É|Ê|Ë|Ì|Í|Î|Ï|Ð|Ñ|Ò|Ó|Ô|Õ|Ö|×|Ø|Ù|Ú|Û|Ü|Ý|Þ|ß|à|ÿ|â|ã|ä|å|æ|ç|è|é|ê|ë|ì|í|î|ï|ð|ñ|ò|ó|ô|õ|ö|÷|ø|ù|ú|û|ü|ý|þ|ÿ"
GoodList = Split(ListGood, "|")
BadList = Split(ListBad, "|")
Upper = UBound(GoodList)
For L = 0 To Upper
    OutLine = Replace(OutLine, BadList(L), Chr(Val(GoodList(L))))
Next
ConvertChars = NulTrim(OutLine)

   On Error GoTo 0
   Exit Function

ConvertChars_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ConvertChars of Module modUnicode"

End Function


