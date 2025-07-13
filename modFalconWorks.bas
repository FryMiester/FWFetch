Attribute VB_Name = "modFalconWorks"
'---------------------------------------------------------------------------------------
' Module    : modFalconWorks
' DateTime  : 4/7/03 17:27
' Author    : Tim Bennett
' Purpose   : Main Procedure code
'---------------------------------------------------------------------------------------
'Code to be placed in module "modMain.bas"
'Option Explicit

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH = 260
Private Const FS_CASE_IS_PRESERVED = 2
Private Const FS_CASE_SENSITIVE = 1
Private Const FS_UNICODE_STORED_ON_DISK = 4
Private Const FS_PERSISTENT_ACLS = 8
Private Const FS_FILE_COMPRESSION = 16
Private Const FS_VOL_IS_COMPRESSED = 32768

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const INVALID_HANDLE_VALUE = -1&

Public Const ARRAY_SIZE = 5000
Public m_lCurIndex As Long
Public Const CB_SHOWDROPDOWN = &H14F    ' Used By Show Dropdown

' =========================================================================
Public Const KB = 1024
Public Const MB = 1048576
Public Const GB = 1073741824
Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&          ' Registry path does not exist
Public Const ERROR_ACCESS_DENIED = 5&           ' Requested permissions not available
Public Const ERROR_INVALID_HANDLE = 6&          ' Invalid handle or top-level key
Public Const ERROR_BAD_NETPATH = 53             ' Network path not found
Public Const ERROR_INVALID_PARAMETER = 87       ' Bad parameter to a Win32 API function
Public Const ERROR_CALL_NOT_IMPLEMENTED = 120&  ' Function valid only in WinNT/2000?XP
Public Const ERROR_INSUFFICIENT_BUFFER = 122    ' Buffer too small to hold data
Public Const ERROR_BAD_PATHNAME = 161           ' Registry path does not exist
Public Const ERROR_NO_MORE_ITEMS = 259&         ' Invalid enumerated value
Public Const ERROR_BADDB = 1009                 ' Corrupted registry
Public Const ERROR_BADKEY = 1010                ' Invalid registry key
Public Const ERROR_CANTOPEN = 1011&             ' Cannot open registry key
Public Const ERROR_CANTREAD = 1012&             ' Cannot read from registry key
Public Const ERROR_CANTWRITE = 1013&            ' Cannot write to registry key
Public Const ERROR_REGISTRY_RECOVERED = 1014&   ' Recovery of part of registry successful
Public Const ERROR_REGISTRY_CORRUPT = 1015&     ' Corrupted registry
Public Const ERROR_REGISTRY_IO_FAILED = 1016&   ' Input/output operation failed
Public Const ERROR_NOT_REGISTRY_FILE = 1017&    ' Input file not in registry file format
Public Const ERROR_KEY_DELETED = 1018&          ' Key already deleted
Public Const ERROR_KEY_HAS_CHILDREN = 1020&     ' Key has subkeys & cannot be deleted

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const REG_CREATED_NEW_KEY = &H1          ' A new key was created
Public Const REG_OPENED_EXISTING_KEY = &H2      ' An existing key was opened

Public Const REG_OPTION_BACKUP_RESTORE = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_OPTION_VOLATILE = 1

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1

Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const STANDARD_RIGHTS_WRITE = &H20000

Public Const SYNCHRONIZE = &H100000

Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_SET_VALUE = &H2
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Type SECURITY_ATTRIBUTES

    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long

End Type

Public Type FILETIME

    dwLowDateTime As Long
    dwHighDateTime As Long

End Type

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Const MB_DEFAULTBEEP As Long = -1   ' the default beep sound
Public Const MB_ERROR As Long = 16        ' for critical errors/problems
Public Const MB_WARNING As Long = 48      ' for conditions that might cause problems in the future
Public Const MB_INFORMATION As Long = 64  ' for informative messages only

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function RegOpenKeyEx _
                         Lib "advapi32.dll" _
                             Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                    ByVal lpSubKey As String, _
                                                    ByVal ulOptions As Long, _
                                                    ByVal samDesired As Long, _
                                                    phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx _
                         Lib "advapi32.dll" _
                             Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                      ByVal lpSubKey As String, _
                                                      ByVal Reserved As Long, _
                                                      ByVal lpClass As String, _
                                                      ByVal dwOptions As Long, _
                                                      ByVal samDesired As Long, _
                                                      lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                      phkResult As Long, _
                                                      lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx _
                         Lib "advapi32.dll" _
                             Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                     ByVal lpValueName As String, _
                                                     ByVal Reserved As Long, _
                                                     ByVal dwType As Long, _
                                                     lpData As Any, _
                                                     ByVal cbData As Long) As Long
Public Declare Function RegQueryInfoKey _
                         Lib "advapi32.dll" _
                             Alias "RegQueryInfoKeyA" (ByVal hKey As Long, _
                                                       ByVal lpClass As String, _
                                                       lpcbClass As Long, _
                                                       ByVal lpReserved As Long, _
                                                       lpcSubKeys As Long, _
                                                       lpcbMaxSubKeyLen As Long, _
                                                       lpcbMaxClassLen As Long, _
                                                       lpcValues As Long, _
                                                       lpcbMaxValueNameLen As Long, _
                                                       lpcbMaxValueLen As Long, _
                                                       lpcbSecurityDescriptor As Long, _
                                                       lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumKeyEx _
                         Lib "advapi32.dll" _
                             Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                                    ByVal dwIndex As Long, _
                                                    ByVal lpName As String, _
                                                    lpcbName As Long, _
                                                    ByVal lpReserved As Long, _
                                                    ByVal lpClass As String, _
                                                    lpcbClass As Long, _
                                                    lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegDeleteKey _
                         Lib "advapi32.dll" _
                             Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                    ByVal lpSubKey As String) As Long

Rem *** The function above will fail if the key to be deleted contains any subkeys. Use the following function in that case ***
Public Declare Function SHDeleteKey _
                         Lib "shlwapi.dll" _
                             Alias "SHDeleteKeyA" (ByVal hKey As Long, _
                                                   ByVal pszSubKey As String) As Long

'--------------------------------------------------------------------------------
' Project    :       RegistrySaveGet
' Procedure  :       Main
' Description:       Start up module to call the form
' Parameters :       None
'------------------------------------------------------------------------

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


'-------------------------------------------------

Global Const GW_HWNDNEXT = 2
Global Const GW_CHILD = 5
Global Const GWW_ID = (-12)
Global Rtn_Value As String
Global Result As Boolean
Global Debugger As Boolean
Global Aborted As Boolean
Global Transfer As String
Global INIFileName As String

' Constants used to detect clicking on the icon
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

' Constants used to control the icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Used as the ID of the call back message
Public Const WM_MOUSEMOVE = &H200

Private Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Type T_FILES
    Path As String
    File As String
End Type

' Used by Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                             Alias "SHGetPathFromIDListA" _
                                             (ByVal pidl As Long, _
                                              ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                           Alias "SHBrowseForFolderA" _
                                           (lpBrowseInfo As BrowseInfo) As Long

Private Declare Function GetSystemDirectory Lib _
                                            "kernel32" Alias "GetSystemDirectoryA" _
                                            (ByVal lpBuffer As String, _
                                             ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib _
                                             "kernel32" Alias "GetWindowsDirectoryA" _
                                             (ByVal lpBuffer As String, _
                                              ByVal nSize As Long) As Long

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                                                   ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
                                     "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
                                                      ByVal lpFile As String, ByVal lpParameters As String, _
                                                      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'create variable of type NOTIFYICONDATA
Public TrayIcon As NOTIFYICONDATA

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, Source As Any, ByVal bytes As Long)

'Find files and Dirs APIs
Private Declare Function FindFirstFile Lib "kernel32" Alias _
                                       "FindFirstFileA" (ByVal lpFileName As String, _
                                                         lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias _
                                      "FindNextFileA" (ByVal hFindFile As Long, _
                                                       lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
                                         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SearchTreeForFile Lib _
                                           "imagehlp.dll" (ByVal lpRootPath As String, ByVal _
                                                                                       lpInputPathName As String, ByVal lpOutputPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                                                                      ByVal lpFileName As String, _
                                                                                      ByVal nSize As Long) _
                                                                                      As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

' GetFileSize

'Private Const INVALID_HANDLE_VALUE = -1

'Private Type FILETIME
'dwLowDateTime As Long
'dwHighDateTime As Long
'End Type

'Private Type WIN32_FIND_DATA
'dwFileAttributes As Long
'ftCreationTime As FILETIME
'ftLastAccessTime As FILETIME
'ftLastWriteTime As FILETIME
'nFileSizeHigh As Long
'nFileSizeLow As Long
'dwReserved0 As Long
'dwReserved1 As Long
'cFileName As String * MAX_PATH
'cAlternate As String * 14
'End Type





' Custom Enum for the return unit type
Public Enum FILE_SIZE_UNIT
FSU_BYTES = 0
FSU_KBYTES = 1
FSU_MBYTES = 2
FSU_GBYTES = 3
End Enum

Global SizeOfFile As Double


' GetFileSize

Public Function GetFileSize(ByVal sFile As String, ByVal eSizeUnit As FILE_SIZE_UNIT) As Double

' Function to get file size which handles files > 4 GB
' Type Declaration for Double is # so we'll force all calculations
' to use Double
'
' eSizeUnit is used to determine how the caller wants the result represented
' (Bytes, KB, MB or GB)
'
' For example if you call the function with FSU_MBYTES then the result is the
' size of the file in Megabytes. If you use FSU_BYTES, the result shoule be
' the same as displayed on the file's property sheet in Explorer

Dim wFD As WIN32_FIND_DATA
Dim hFile As Long
Dim dblFileSize As Currency

hFile = FindFirstFile(sFile, wFD)

' If file doesn't exist, hFile will be INVALID_HANDLE_VALUE
If hFile = INVALID_HANDLE_VALUE Then

dblFileSize = wFD.nFileSizeLow

' Files greater than 4GB return a negative value for nFileSizeLow and also
' populate nFileSizeHigh with some data. Need to do some calculations to
' get the correct size

If wFD.nFileSizeLow < 0 Then
' Add 4GB to nFileSizeLow (the default for dblFileSize at this time)
dblFileSize = dblFileSize + 4294967296#
End If

If wFD.nFileSizeHigh > 0 Then
' Multiply nFileSizeHigh by 4GB
dblFileSize = dblFileSize + (wFD.nFileSizeHigh * 4294967296#)
End If

' Divide the result by the number of bytes in the requested return type
' and Fix() the results to 2 decimal places without rounding. We use Fix()
' because if the number is negative, Int() rounds the wrong direction:
' Int(-99.2) = -100
' Fix(-99.2) = -99

Select Case eSizeUnit
Case FSU_BYTES
' Default (no conversion needed)
GetFileSize = dblFileSize
Case FSU_KBYTES
' KB (1,024 Bytes = 1 KB)
GetFileSize = Fix((dblFileSize / 1024#) * 100) / 100
Case FSU_MBYTES
' MB (1,048,576 Bytes = 1 MB)
GetFileSize = Fix((dblFileSize / 1048576#) * 100) / 100
Case FSU_GBYTES
' GB (1,073,741,824 Bytes = 1 GB)
GetFileSize = Fix((dblFileSize / 1073741824#) * 100) / 100
Case Else
' Unknown
GetFileSize = 0
End Select

End If

' Cleanup
Call FindClose(hFile)

End Function

 
Public Function FileExists(ByRef sFileName As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
End Function

 
Public Function InIDE() As Boolean
Dim S As String
 
    S = Space$(255)
    
    Call GetModuleFileName(GetModuleHandle(vbNullString), S, Len(S))
    
    InIDE = (UCase$(Trim$(S)) Like "*VB6.EXE*")
    
End Function

Public Function BoolToInt(Value As Boolean) As Integer
Dim Rtn As Integer
   On Error GoTo BoolToInt_Error

If Value = True Then
    Rtn = 1
Else
    Rtn = 0
End If
BoolToInt = Rtn

   On Error GoTo 0
   Exit Function

BoolToInt_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure BoolToInt of Module modFalconWorks"
End Function

Public Sub MsgBox(Message As String, Optional DisPlaySeconds As Integer)
On Local Error Resume Next
Dim TimeOut As String
TimeOut = CStr(DisPlaySeconds)
Load frmMsgBox
With frmMsgBox
    .txtDisplayText.text = Message
    .txtTimeOut.text = TimeOut
    .Show
End With
End Sub


Public Function SizeFile(inFile As String, Optional MySize As Double) As String
    Dim fSize As Double
    Dim TheSize As Double
    Dim DoubleBytes As Double
    
   On Error GoTo SizeFile_Error

    On Error Resume Next

    If inFile = "" And MySize > 0 Then
        fSize = MySize
    Else
        fSize = FileLen(inFile)
    End If

    If fSize = 0 Then
        SizeFile = "00KB"
        Exit Function
    End If
    TheSize = fSize

Select Case TheSize
            Case Is >= 1099511627776#
                DoubleBytes = CDbl(TheSize / 1099511627776#) 'TB
                SizeFile = Format(DoubleBytes, "Standard") & " TB"
            Case 1073741824 To 1099511627775#
                DoubleBytes = CDbl(TheSize / 1073741824) 'GB
                SizeFile = Format(DoubleBytes, "Standard") & " GB"
            Case 1048576 To 1073741823
                DoubleBytes = CDbl(TheSize / 1048576) 'MB
                SizeFile = Format(DoubleBytes, "Standard") & " MB"
            Case 1024 To 1048575
                DoubleBytes = CDbl(TheSize / 1024) 'KB
                SizeFile = Format(DoubleBytes, "Standard") & " KB"
            Case 0 To 1023
                DoubleBytes = TheSize ' bytes
                SizeFile = Format(DoubleBytes, "Standard") & " bytes"
            Case Else
                SizeFile = "00KB"
        End Select

   On Error GoTo 0
   Exit Function

SizeFile_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure SizeFile of Module modFalconWorks"
End Function

Public Sub LBoxNoDupes(lbox As ListBox)
    Dim i As Long, x As Long, y As Long

    On Error GoTo LBoxNoDupes_Error

    For i = 0 To lbox.ListCount - 1
        DoEvents
        For x = 0 To lbox.ListCount - 1
            DoEvents
            If x <> i Then
                If UCase(NulTrim(lbox.List(i))) = UCase(NulTrim(lbox.List(x))) Then
                    lbox.RemoveItem x
                    x = x - 1
                    y = y + 1
                End If
            End If
        Next
    Next

    On Error GoTo 0
    Exit Sub

LBoxNoDupes_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure LBoxNoDupes of Module modPlayMasterMain"

End Sub

Public Function CurrLboxData(lbox As ListBox) As Long

' -----------  ERROR HANDLER  ------------------
On Error GoTo CurrLboxData_Error
' -----------  ERROR HANDLER  ------------------

If lbox.ListIndex <> -1 Then
    CurrLboxData = lbox.ItemData(lbox.ListIndex)
Else
CurrLboxData = -1
End If

On Error GoTo 0
Exit Function

CurrLboxData_Error:

End Function



Public Sub AddLboxItem(lboxAdd As ListBox, ByVal sText As String, ByVal lData As Long, Optional OldIndex As Integer)

   On Error GoTo AddLboxItem_Error

    If OldIndex = 0 Then
        lboxAdd.AddItem sText
        lboxAdd.ItemData(lboxAdd.NewIndex) = lData
    Else
        lboxAdd.AddItem sText, OldIndex
        lboxAdd.ItemData(OldIndex) = lData
    End If

   On Error GoTo 0
   Exit Sub


AddLboxItem_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure AddLboxItem of Module modFalconWorks"
End Sub

Public Sub DeleteINISection(Header As String, Optional INIFName As String)
    Dim U As Long
    Dim OpenName As String
    If INIFName = "" Then
        OpenName = INIFileName
    Else
        OpenName = INIFName
    End If
    U = WritePrivateProfileString(Header, vbNullString, vbNullString, OpenName)
    If U = 0 Then
    End If
End Sub

Public Sub DeleteINIValue(Header As String, Section As String, Optional INIFName As String)
    Dim U As Long
    Dim OpenName As String
    If INIFName = "" Then
        OpenName = INIFileName
    Else
        OpenName = INIFName
    End If
    U = WritePrivateProfileString(Header, Section, vbNullString, OpenName)
    If U = 0 Then
    End If
End Sub


Public Sub WriteMyINI(Header As String, keyname As String, Value As String, Optional INIFName As String)
    Dim U As Long
    Dim OpenName As String
    If INIFName = "" Then
        OpenName = INIFileName
    Else
        OpenName = INIFName
    End If
    U = WritePrivateProfileString(Header, keyname, Value, OpenName)
    If U = 0 Then
    End If
End Sub

Public Function ReadINI(Header As String, keyname As String, sINIFileName As String, Optional DefaultValue As String = "") As String
    Dim sRet As String
    Dim Result As String
    
    sRet = String(255, Chr(0))
    Result = Left(sRet, GetPrivateProfileString(Header, ByVal keyname, DefaultValue, sRet, Len(sRet), sINIFileName))
    
    ' If no value found and we have a default, return the default
    If Result = "" And DefaultValue <> "" Then
        ReadINI = DefaultValue
    Else
        ReadINI = Result
    End If
End Function

Public Function NulTrim(text As String) As String
    Dim Z, L As Integer
    Dim r As Integer
    Dim q As Integer
    Dim P As Integer
    Dim Test As String
   On Error GoTo NulTrim_Error

    Z = Len(text)
    r = InStr(text, Chr$(0))
    q = InStr(text, Chr$(255))
    If r = 0 And q = 0 Then
        NulTrim = LTrim$(RTrim$(text))
        Exit Function
    End If
    If r > 0 And q = 0 Then
        P = r
    ElseIf q > 0 And r = 0 Then
        P = q
    ElseIf r > 0 And q > 0 Then
        P = 1
    End If
    For L = P To Z
        Test = Mid$(text, L, 1)
        If Test = Chr$(0) Or Test = Chr$(255) Then
            Mid$(text, L, 1) = " "
        End If
    Next
    NulTrim = RTrim$(text)

   On Error GoTo 0
   Exit Function

NulTrim_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure NulTrim of Module modFalconWorks"
End Function

Public Function NoPath(FullPathName As String) As String
    Dim prom As String
    Dim AA As Integer, A As Integer
   On Error GoTo NoPath_Error

    FullPathName = LTrim(RTrim(FullPathName))
    AA = 0: A = 0
    AA = Len(FullPathName)
    prom = FullPathName
    Do
        A = InStr(prom, "\")
        AA = Len(prom)
        If A > 0 Then
            prom = Right(prom, AA - A)
        End If
    Loop Until A = 0
    NoPath = prom

   On Error GoTo 0
   Exit Function

NoPath_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure NoPath of Module modFalconWorks"
End Function

Public Function nWord(text As String, Which As String) As String
    Dim Z As Long
    Dim A As Long
   On Error GoTo nWord_Error

    If Len(Which) > 1 Then Which = Left(Which, 1)
    Z = Len(text)
    A = InStr(text, Which)
    If A = 0 Then
        nWord = text: text = ""
        Exit Function
    End If
    nWord = Left$(text, A - 1)
    text = Right$(text, Z - A)

   On Error GoTo 0
   Exit Function

nWord_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure nWord of Module modFalconWorks"
End Function

Public Function rand(lowest, highest)
   On Error GoTo rand_Error

    Randomize Timer
    rand = Int((highest - lowest + 1) * Rnd + lowest)

   On Error GoTo 0
   Exit Function

rand_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure rand of Module modFalconWorks"
End Function

Public Function FileExist(FileName As String) As Boolean
    Dim x As Date
    On Error GoTo bad
    
    x = VBA.FileDateTime(FileName)
    If x > "12:00:00 AM" Then
        FileExist = True
    Else
        FileExist = False
    End If
    Exit Function
bad:
    FileExist = False
End Function

Public Sub UnloadAllForms()
    Dim Form As Form
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
End Sub

Public Sub Swap(Item1 As Variant, Item2 As Variant)
    Dim SaveItem As Variant
   On Error GoTo Swap_Error

    SaveItem = Item1
    Item1 = Item2
    Item2 = SaveItem

   On Error GoTo 0
   Exit Sub

Swap_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure Swap of Module modFalconWorks"
End Sub

Public Function RunPath(Optional FullRunPath As String) As String
    Dim Temp As String
    Dim Test As String
    Dim A As Integer

   On Error GoTo RunPath_Error

    If FullRunPath = "" Then
        RunPath = App.Path
        Exit Function
    End If
    Temp = FullRunPath
    Do
        A = InStr(Temp, "\")
        If A > 0 Then
            Test = Test$ + Left$(Temp, A)
            Temp = Right$(Temp, Len(Temp) - A)
        End If
    Loop Until InStr(Temp, "\") = 0
    RunPath = Test

   On Error GoTo 0
   Exit Function

RunPath_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure RunPath of Module modFalconWorks"
End Function

Public Function CRC32(B As String) As Long
Dim i As Long
Dim j As Long
Dim ByteVal As Integer

    Dim Power(0 To 7)
    Dim CRC As Long
    For i = 0 To 7
        Power(i) = 2 ^ i
    Next i
    CRC = 0
    For i = 1 To Len(B)
        ByteVal = Asc(Mid$(B, i, 1))
        For j = 7 To 0 Step -1
            TestBit = ((CRC And 32768) = 32768) Xor ((ByteVal And Power(j)) = Power(j))
            CRC = ((CRC And 32767&) * 2&)
            If TestBit Then CRC = CRC Xor &H8005&
        Next j
    Next i
    CRC32 = CRC
End Function

Private Function Process_Command(TextString As String) As String
    Dim zz As Integer
    Dim L As Long
    Dim Temp As String
    zz = Len(TextString)
    For L = 1 To zz
        If Mid$(TextString, L, 1) <> Chr$(34) Then
            Temp = Temp & Mid$(TextString, L, 1)
        End If
    Next
    Process_Command = Temp
End Function

Public Sub PutOnTop(hWnd As Long)
    Dim OnTop As Long
   On Error GoTo PutOnTop_Error

    OnTop = SetWindowPos(hWnd, -1, 0, 0, 0, 0, 2 Or 1)

   On Error GoTo 0
   Exit Sub

PutOnTop_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure PutOnTop of Module modFalconWorks"
End Sub

Public Sub TakeOffTop(hWnd As Long)
    Dim OnTop As Long
   On Error GoTo TakeOffTop_Error

    OnTop = SetWindowPos(hWnd, 1, 0, 0, 0, 0, 2 Or 1)

   On Error GoTo 0
   Exit Sub

TakeOffTop_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure TakeOffTop of Module modFalconWorks"
End Sub

Public Function PercentOf(BigNumber As Variant, LittleNumber As Variant) As Integer
    On Error Resume Next
    PercentOf = Int((100 / BigNumber) * LittleNumber)
End Function

Public Function ProperCase(InLine As String) As String
    Dim fName As String
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
   On Error GoTo ProperCase_Error

    On Error Resume Next
    fName = InLine
    InLine = LCase(NoPath(InLine))
    Temp = InLine
    For L = 1 To Len(Temp)
        If Mid(Temp, L, 1) = "_" Then Mid$(Temp, L, 1) = " "
    Next
    Do
        Temp1 = nWord(Temp, " ")
        Mid$(Temp1, 1, 1) = UCase(Mid$(Temp1, 1, 1))
        Temp2 = Temp2 & " " & Temp1
    Loop While Len(Temp)
    Temp2 = Trim(Temp2)
    Temp2 = RunPath(fName) & Temp2
    ProperCase = Temp2

   On Error GoTo 0
   Exit Function

ProperCase_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ProperCase of Module modFalconWorks"

End Function

Public Function noExt(text As String) As String
    Dim A As Integer
    Dim Z As Integer
    Dim r As Integer

   On Error GoTo noExt_Error

    A = Len(text)

    For r = A To 1 Step -1
        If Mid$(text, r, 1) = "." Then
            Z = r - 1
            Exit For
        End If
    Next
    If Z > 0 Then
        noExt = Left$(text, Z)
    Else
        noExt = text
    End If

   On Error GoTo 0
   Exit Function

noExt_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure noExt of Module modFalconWorks"
End Function

Public Function GetLongFilename(ByVal sShortName As String) As String
    Dim sLongName As String
    Dim sTemp As String
    Dim iSlashPos As Integer

    'Add \ to short name to prevent Instr from failing
    sShortName = sShortName & "\"
    'Start from 4 to ignore the "[Drive Letter]:\" characters
    iSlashPos = InStr(4, sShortName, "\")
    'Pull out each string between \ character for conversion
    While iSlashPos
        sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
        If sTemp = "" Then            'Error 52 - Bad File Name or Number
            GetLongFilename = ""
            Exit Function
        End If
        sLongName = sLongName & "\" & sTemp
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Wend
    'Prefix with the drive letter
    GetLongFilename = Left$(sShortName, 2) & sLongName

End Function

Public Function DirExist(Directory As String) As Boolean
    Dim x As String
    x = CurDir
    On Error GoTo nope
    ChDir Directory
    ChDir x
    DirExist = True
nope:
End Function

Public Function StripHighBits(InString As String) As String
    Dim TempStr As String
    Dim Hold As String
    Dim TheLength As Integer
    If Len(InString) = 0 Then
        StripHighBits = ""
        Exit Function
    End If
    TheLength = Len(InString)
    TempStr = String(TheLength, " ")

    For L = 1 To TheLength
        Hold = Mid$(InString, L, 1)
        Select Case Asc(Hold)
        Case 13 To 128
            Mid$(TempStr, L, 1) = Hold
        Case Else
            Mid$(TempStr, L, 1) = " "
        End Select
    Next
    TempStr = NulTrim(TempStr)
    StripHighBits = TempStr
End Function

Public Function StripQuotes(InString As String) As String
Dim tmpString As String
    tmpString = InString
    tmpString = Replace(tmpString, Chr(34), "")
    StripQuotes = tmpString
End Function

Public Function SecToMin(TheSeconds As Single) As String
    Dim Minutes As String
    Dim Seconds As String
   
   On Error GoTo SecToMin_Error
    If TheSeconds > 3599 Or TheSeconds < 1 Then
        Exit Function
    End If
    
    Minutes = Right$("00" & CStr(TheSeconds \ 60), 2)
    Seconds = Right$("00" & CStr(TheSeconds Mod 60), 2)
    SecToMin = Minutes & ":" & Seconds

   On Error GoTo 0
   Exit Function

SecToMin_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure SecToMin of Module modFalconWorks"
End Function

Public Function FindFile(ByVal StartingPath As String, ByVal FileName As String) As String
    Dim NullChar As Integer
    Dim x As Long
    Dim Buffer As String
    Buffer = String$(1024, 0)
    x = SearchTreeForFile(StartingPath, FileName, Buffer)
    If Not x Then
        FindFile = ""
    End If
    If x Then
        NullChar = InStr(Buffer, vbNullChar)
        If NullChar Then
            Buffer = Left$(Buffer, NullChar - 1)
        End If
        FindFile = Buffer
    End If
End Function

Public Sub ListAllFilesByExt(ByVal sPath As String, _
                             uFiles() As T_FILES, _
                             Optional ByVal sExt As String, _
                             Optional bSearchSubDirs = True)
'********************************************************
'Searches for files by given path and file extension.   *
'Files found are put in the uFiles() array. Optionally  *
'searches in all subdirs recursively.                   *
'********************************************************
'Input:                                                 *
'       sPath - string, containing the directory to     *
'               search in.                              *
'       sExt - string, optional - allows to search for  *
'              files with given extention.              *
'       bSearchSubDirs - boolean, optional - if TRUE    *
'              will cause to search all subdirs under   *
'              the directory pointed by sPath.          *
'********************************************************
'Return:                                                *
'       uFiles() - array of T_FILES containing the files*
'                  and their paths.                     *
'********************************************************

    Dim sFile As String
    Dim hFind As Long
    Dim wFD As WIN32_FIND_DATA
    Dim nExtLen As Integer
    Dim lRetVal As Long
    Dim sTemp As String
    Dim lPos As Long
    On Error Resume Next
    If sExt <> "" Then
        If Left$(sExt, 1) <> "." Then
            sExt = "." & sExt
        End If

        nExtLen = Len(sExt)
    End If

    If bSearchSubDirs Then
        If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If

        sFile = sPath & "*.*"

        hFind = FindFirstFile(sFile, wFD)

        If hFind <> 0 Then
            Do
                With wFD
                    lPos = InStr(1, .cFileName, Chr$(0))
                    If lPos > 0 Then
                        sTemp = Left$(.cFileName, lPos - 1)
                        If (.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                            'it is another subdir to search
                            If sTemp <> "." And sTemp <> ".." Then
                                'Start searching the new dir recursively
                                ListAllFilesByExt sPath & sTemp, uFiles, sExt
                            End If
                        Else
                            'It is a file
                            If nExtLen > 0 Then
                                'List only files with given extention.
                                If UCase$(Right$(sTemp, nExtLen)) = UCase$(sExt) Then
                                    'Add the file to the list
                                    uFiles(m_lCurIndex).File = sTemp
                                    uFiles(m_lCurIndex).Path = sPath

                                    'Increment index
                                    m_lCurIndex = m_lCurIndex + 1
                                End If
                            Else
                                'List all files

                                'Add the file to the list
                                uFiles(m_lCurIndex).File = sTemp
                                uFiles(m_lCurIndex).Path = sPath

                                'Increment index
                                m_lCurIndex = m_lCurIndex + 1
                            End If

                            'Check if we need to resize buffer
                            If m_lCurIndex >= UBound(uFiles) Then
                                ReDim Preserve uFiles(UBound(uFiles) + ARRAY_SIZE)
                            End If
                        End If  'If (.dwFileAttributes
                    End If  'If lPos > 0
                End With

                'Search for next file
                lRetVal = FindNextFile(hFind, wFD)

            Loop While lRetVal <> 0

            FindClose hFind

        End If  'If hFind <> 0
    Else
        'No Sub Dir search is required.
        If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If

        If nExtLen > 0 Then
            sFile = sPath & "*" & sExt
        Else
            sFile = sPath & "*.*"
        End If

        hFind = FindFirstFile(sFile, wFD)
        If hFind <> 0 Then
            Do
                With wFD
                    lPos = InStr(1, .cFileName, Chr$(0))
                    If lPos > 0 Then
                        sTemp = Left$(.cFileName, lPos - 1)

                        If (.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
                            'It is a file
                            'Add the file to the list

                            uFiles(m_lCurIndex).File = sTemp
                            uFiles(m_lCurIndex).Path = sPath

                            'Increment index
                            m_lCurIndex = m_lCurIndex + 1

                            'Check if we need to resize buffer
                            If m_lCurIndex >= UBound(uFiles) Then
                                ReDim Preserve uFiles(UBound(uFiles) + ARRAY_SIZE)
                            End If
                        End If
                    End If
                End With

                'Search for next file
                lRetVal = FindNextFile(hFind, wFD)
            Loop While lRetVal <> 0

            FindClose hFind
        End If
    End If
End Sub

Public Function PrintItUsing(FormatString As String, Params() As Variant, Optional HideZero As Boolean) As String
    Dim DataCount As Integer
    Dim HasTags As Boolean
    Dim Tags() As String
    Dim TagCounter As Integer
    Dim Marker As Integer
    Dim intLoop As Integer
    Dim ReturnLength As Integer
    Dim FullString As String

    Dim TagLength() As Integer
    Dim StartPos() As Integer
    Dim StopPos() As Integer
    Dim TempString As String

    ' -----------  ERROR HANDLER  ------------------
    'On Error GoTo PrintItUsing_Error
    On Error Resume Next
    ' -----------  ERROR HANDLER  ------------------

    DataCount = UBound(Params)

    ReDim TagLength(DataCount)
    ReDim StartPos(DataCount)
    ReDim StopPos(DataCount)

    TempString = FormatString
    FullString = String(Len(TempString), Chr(32))
    ReturnLength = Len(TempString)
    Dim Index As Integer

    If StringCount(FormatString, "\") Mod 2 <> 0 Then
        GoTo PrintItUsing_Error
    End If

    For intLoop = 1 To ReturnLength
        If Mid(TempString, intLoop, 1) = "\" Then
            Marker = Marker + 1
            If Marker = 3 Then
                Marker = 1
                Index = Index + 1
            End If
            If Marker Mod 2 = 0 Then
                StopPos(Index) = intLoop + 1
                TagLength(Index) = StopPos(Index) - StartPos(Index)
            Else
                StartPos(Index) = intLoop
            End If
        End If
    Next

    DataCount = Index
    ReDim Tags(Index)

    For intLoop = 0 To Index
        Tags(intLoop) = Mid(TempString, StartPos(intLoop), (StopPos(intLoop) - StartPos(intLoop)))
        Tags(intLoop) = Replace(Tags(intLoop), "\", "")
    Next

    TempString = Replace(TempString, "\", " ")

    Dim TempNum As Long
    Dim TempStr As String
    Dim MakeDate As Date
    Dim ayear As Integer

    For intLoop = 0 To Index
        TempNum = Len(Tags(intLoop))
        TempStr = String(TagLength(intLoop), " ")

        If InStr(Tags(intLoop), "#") > 0 Or InStr(Tags(intLoop), "/") > 0 Then
            If InStr(Tags(intLoop), "/") > 0 Then
                TempStr = TempStr & Format(CDate(Params(intLoop)), Tags(intLoop))
                TempStr = Right(TempStr, TagLength(intLoop))
                TempStr = Right(TempStr, TagLength(intLoop) - 1)
            Else
                TempStr = TempStr & Format(CDbl(Params(intLoop)), Tags(intLoop))
                TempStr = Right(TempStr, TagLength(intLoop))
                TempStr = Replace(TempStr, "#", "")
                If HideZero = True Then
                    If Val(Format(CDbl(Params(intLoop)), Tags(intLoop))) = 0 Then
                        TempStr = String(TagLength(intLoop), " ")
                    End If
                End If
            End If
        Else
            TempStr = CStr(Params(intLoop)) & TempStr
            TempStr = Left(TempStr, TagLength(intLoop))
        End If
        Tags(intLoop) = String(TagLength(intLoop), Chr(32))
        Mid(Tags(intLoop), 1, TagLength(intLoop)) = TempStr
    Next

    For intLoop = 0 To Index

        Mid(TempString, StartPos(intLoop), TagLength(intLoop)) = Left(Tags(intLoop), TagLength(intLoop))
    Next

    PrintItUsing = TempString

    On Error GoTo 0
    Exit Function

PrintItUsing_Error:

    PrintItUsing = "Unknown Error"

End Function

Public Function FixPath(PathName As String) As String
    Dim TempStr As String
    If Right(PathName, 1) <> "\" Then
        TempStr = PathName & "\"
    Else
        TempStr = PathName
    End If
    FixPath = TempStr
End Function

Public Sub Pause(ByVal nSecond As Single)
    Dim t0 As Single
   On Error GoTo Pause_Error

    t0 = Timer
    Do While Timer - t0 < nSecond
        Dim dummy As Integer
        dummy = DoEvents()
        ' if we cross midnight, back up one day
        If Timer < t0 Then
            t0 = t0 - CLng(24) * CLng(60) * CLng(60)
        End If
    Loop

   On Error GoTo 0
   Exit Sub

Pause_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Module modFalconWorks"
End Sub


Public Function StringCount(InputString As String, TestString As String) As Integer
Dim Character As String
Dim CharCount As Integer
Dim Addup As Integer
If InputString = "" Then
    StringCount = 0
    Exit Function
End If
Character = TestString
CharCount = Len(InputString)
For L = 1 To CharCount
    If Mid(InputString, L, Len(Character)) = Character Then
        Addup = Addup + 1
    End If
Next
StringCount = Addup
End Function

Public Function CenterIt(text As String, Width As Integer)
    Dim Half As Integer
    Dim Filler As String
    Dim LE As Integer
    Half = Width \ 2
    Filler = String(Width, " ")
    LE = Len(text)
    Mid(Filler, (Width - LE) \ 2, LE) = text
    CenterIt = Filler
End Function

Public Sub LogError(ErrMsg As String, Optional ErrorLog As String)
    Static Active As Boolean
   On Error GoTo LogError_Error

    If Active = True Then Exit Sub
    Active = True
    ErrMsg = Format(Now, "hh:mm:ss MM/DD/YYYY ") & ErrMsg
    Dim ErrFile As String
    Dim ErrHandle As Integer
    ErrHandle = FreeFile
    If ErrorLog = "" Then
        ErrFile = App.Path & "\Error.log"
    Else
        ErrFile = ErrorLog
    End If
    Open ErrFile For Append As ErrHandle
    Print #ErrHandle, ErrMsg
    Close ErrHandle
    Active = False

   On Error GoTo 0
   Exit Sub

LogError_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure LogError of Module modFalconWorks"
End Sub



Public Function Epoch2Date(EpochDate As Long) As Date
    Dim L2Date As Date
    Dim Long2Date As Date
    Long2Date = CDate("01/01/1970") + (EpochDate / 86400)
    L2Date = CDate(Now - 155)
    Epoch2Date = Long2Date
End Function

Public Function Date2Epoch(dtmDate As Date) As Long
    Date2Epoch = (dtmDate - CDate("1/1/1970")) * 86400
End Function

Public Function IsValidDay(DayInput As Variant) As Boolean
    On Error GoTo nope
    IsValidDay = IsDate(DayInput)
nope:
End Function

Public Function IsWeekend(D As Variant) As Boolean

    If Not IsDate(D) Then
        IsWeekend = False
        Exit Function
    End If

    If Weekday(D) = 1 Or Weekday(D) = 7 Then
        IsWeekend = True
    Else
        IsWeekend = False
    End If

End Function


Public Function BrowseFileName(MeForm As Form, Optional FilterString As String, Optional StartPath As String) As String
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    Dim TheStartPath As String

   On Error GoTo BrowseFileName_Error

    If StartPath > "" Then
        If DirExist(StartPath) = True Then
            TheStartPath = StartPath
        Else
            TheStartPath = "C:\"
        End If
    Else
        TheStartPath = "C:\"
    End If
   
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = MeForm.hWnd
    OpenFile.hInstance = App.hInstance
    If FilterString = "" Then
        sFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Else
        sFilter = FilterString
    End If
    
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = TheStartPath
    OpenFile.lpstrTitle = "Search Files"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)

    If lReturn = 0 Then
        BrowseFileName = ""
    Else
        BrowseFileName = Trim(OpenFile.lpstrFile)
    End If

   On Error GoTo 0
   Exit Function

BrowseFileName_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure BrowseFileName of Module modFalconWorks"
End Function

Public Function BrowseFolder(FormName As Form, Optional sTitle As String, Optional Rootfolder As String) As String
    Dim initBrowseInfo As BrowseInfo
    Dim pidl As Long
    Dim Path As String
    Dim pos As Long


    'Fill the BROWSEINFO structure with the
    'needed data. To accommodate comments, the
    'With/End With syntax has not been used, though
    'it should be your 'final' version.

    With initBrowseInfo

        'hwnd of the window that receives messages
        'from the call. Can be your application
        'or the handle from GetDesktopWindow()
        .hOwner = FormName.hWnd

        'pointer to the item identifier list specifying
        'the location of the "root" folder to browse from.
        'If NULL, the desktop folder is used.
            .pIDLRoot = 0&
        
        'message to be displayed in the Browse dialog
        If sTitle = "" Then sTitle = "Select your Windows\System\ directory"
        .lpszTitle = sTitle

        'the type of folder to return.
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'show the browse for folders dialog
    pidl = SHBrowseForFolder(initBrowseInfo)

    'the dialog has closed, so parse & display the
    'user's returned folder selection contained in pidl
    Path = Space$(MAX_PATH)

    If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
        pos = InStr(Path, Chr$(0))
        BrowseFolder = Left(Path, pos - 1)
    End If

    Call CoTaskMemFree(pidl)

End Function


Public Function SortArray(ByRef TheArray As Variant)
    Dim x As Long
   On Error GoTo SortArray_Error

    Sorted = False
    Do While Not Sorted
        Sorted = True
        For x = 0 To UBound(TheArray) - 1
            If TheArray(x) > TheArray(x + 1) Then
                Temp = TheArray(x + 1)
                TheArray(x + 1) = TheArray(x)
                TheArray(x) = Temp
                Sorted = False
            End If
        Next x
    Loop

   On Error GoTo 0
   Exit Function

SortArray_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure SortArray of Module modFalconWorks"
End Function


Public Sub DebugLog(ByVal inText As String, Optional LogFile As String, Optional Mirror As Boolean)
    
   On Error GoTo DebugLog_Error

Exit Sub
    
    If Mirror = True Then Debug.Print inText
    If LogFile > "" Then
        LogError inText, LogFile
    Else
        LogError inText, App.Path & "\DEBUGLOG.txt"
    End If

   On Error GoTo 0
   Exit Sub

DebugLog_Error:

'    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure DebugLog of Module modFalconWorks", App.Path & "\debuglog.txt"

End Sub

Public Function TimeStamp() As Currency
    TimeStamp = (Round(Now(), 0) * 24 * 60 * 60 + Timer()) * 1000
End Function

'   function Sleep let system execute other programs while the milliseconds are not elapsed.

Public Function Sleep(milliseconds As Currency)

    If milliseconds < 0 Then Exit Function

    Dim start As Currency
    start = TimeStamp()

    While (TimeStamp() < milliseconds + start)
        DoEvents
    Wend
End Function

' encrypt a string using a password
'
' you must reapply the same function (and same password) on
' the encrypted string to obtain the original, non-encrypted string
'
' you get better, more secure results if you use a long password
' (e.g. 16 chars or longer). This routine works well only with ANSI strings.

Function EncryptString(ByVal text As String, ByVal PassWord As String) As String
    Dim passLen As Long
    Dim i As Long
    Dim passChr As Integer
    Dim passNdx As Long
    
    passLen = Len(PassWord)
    ' null passwords are invalid
    If passLen = 0 Then Err.Raise 5
    
    ' move password chars into an array of Integers to speed up code
    ReDim passChars(0 To passLen - 1) As Integer
    CopyMemory passChars(0), ByVal StrPtr(PassWord), passLen * 2
    
    ' this simple algorithm XORs each character of the string
    ' with a character of the password, but also modifies the
    ' password while it goes, to hide obvious patterns in the
    ' result string
    For i = 1 To Len(text)
        ' get the next char in the password
        passChr = passChars(passNdx)
        ' encrypt one character in the string
        Mid$(text, i, 1) = Chr$(Asc(Mid$(text, i, 1)) Xor passChr)
        ' modify the character in the password (avoid overflow)
        passChars(passNdx) = (passChr + 17) And 255
        ' prepare to use next char in the password
        passNdx = (passNdx + 1) Mod passLen
    Next

    EncryptString = text
    
End Function

Public Function Encrypt(InputString As String, PassWord As String, Encode As Boolean, Optional OtherKey As String) As String
Dim Matrix(15, 15) As String
Dim Wheel(1) As String * 256
Dim L As Long
Dim TempString As String
Dim QQ As String
Dim OO As String
Dim A, q, r, C, acount, GG, RR, CC As Integer
Dim StringLength As Integer
Dim TempNum As Integer

StringLength = Len(PassWord)
For L = 1 To StringLength
    TempNum = TempNum + Asc(Mid(PassWord, L, 1))
    TempNum = TempNum And L
    TempNum = TempNum Xor StringLength
Next
TempNum = (TempNum Or 256) And StringLength

For L = 0 To 255
    Matrix(L \ 16, L Mod 16) = Chr(L)
    Mid(Wheel(0), L + 1, 1) = Chr(L)
    Mid(Wheel(1), 256 - L, 1) = Chr(L)
Next

shiftem Wheel(0), TempNum, 0
shiftem Wheel(1), 256 - TempNum, 1

For L = 0 To 255
    For r = 0 To 15
        For C = 15 To 0 Step -1
          Matrix(r, C) = Mid$(Wheel(0), L + 1, 1)
        Next
    Next
Next

For r = 0 To 15
    For C = 0 To 15
        GG = Asc(Matrix(r, C))
        RR = GG \ 16
        CC = GG Mod 16
        Swap Matrix(r, C), Matrix(RR, CC)
    Next
Next

' Encode / Decode

If Encode = True Then
  For q = 1 To StringLength
    A = InStr(Wheel(0), Mid$(InputString, q, 1))
    QQ = Mid$(Wheel(1), A, 1)
    OO = Matrix(Asc(QQ) \ 16, Asc(QQ) Mod 16)
    TempString = TempString + OO
  Next
Else
   ' Do the UnScrambling
  For q = 1 To StringLength
    A = InStr(Wheel(1), Mid$(InputString, q, 1))
    OO = Mid$(Wheel(0), A, 1)
    QQ = Matrix(Asc(OO) \ 16, Asc(OO) Mod 16)
    TempString = TempString + QQ
  Next
End If

Encrypt = TempString
End Function

Public Sub shiftem(strng$, Places As Integer, Dir As Integer)
wa% = Len(strng$)
If Dir = 1 Then
Temp$ = ""
Temp$ = Right$(strng$, Places) + Left$(strng$, wa% - Places)
strng$ = Temp$
Else
Temp$ = ""
Temp$ = Right$(strng$, wa% - Places) + Left$(strng$, Places)
strng$ = Temp$
End If
End Sub


Public Function DateLastModified(ByVal sPath As String) As Date
   Dim fso As Object 'New FileSystemObject
   Dim fil As Object 'File
   Dim fol As Object 'Folder
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.FolderExists(sPath) Then
      Set fol = fso.GetFolder(sPath)
      DateLastModified = fol.DateLastModified
   ElseIf fso.FileExists(sPath) Then
      Set fil = fso.GetFile(sPath)
      DateLastModified = fil.DateLastModified
   End If
End Function

Public Sub BubbleSort1(ByRef pvarArray As Variant)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray) - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If pvarArray(i) > pvarArray(i + 1) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + 1)
                pvarArray(i + 1) = varSwap
                blnSwapped = True
            End If
        Next
        iMax = iMax - 1
    Loop Until Not blnSwapped
End Sub

Public Static Sub StrSort(words() As String, Ascending As Boolean, AllLowerCase As Boolean)
 
'Pass in string array you want to sort by reference and
'read it back

'Set Ascending to True to sort ascending, '
'false to sort descending

'If AllLowerCase is True, strings will be sorted
'without regard to case.  Otherwise, upper
'case characters take precedence over lower
'case characters

Dim i As Integer
Dim j As Integer
Dim NumInArray, LowerBound As Long
NumInArray = UBound(words)
LowerBound = LBound(words)
For i = LowerBound To NumInArray
    j = 0
    For j = LowerBound To NumInArray
        If AllLowerCase = True Then
            If Ascending = True Then
                If StrComp(LCase(words(i)), LCase(words(j))) = -1 Then
                    Call Swap(words(i), words(j))
                End If
            Else
                If StrComp(LCase(words(i)), LCase(words(j))) = 1 Then
                    Call Swap(words(i), words(j))
                End If
            End If
        Else
            If Ascending = True Then
                If StrComp(words(i), words(j)) = -1 Then
                    Call Swap(words(i), words(j))
                End If
            Else
                If StrComp(words(i), words(j)) = 1 Then
                    Call Swap(words(i), words(j))
                End If
            End If
        End If
    Next j
Next i
End Sub

Public Static Sub NumSort(nums() As Variant, Ascending As Boolean)

'Pass in numeric array you want to sort by reference and
'read it back.  The array should be declared as an array
'of variants

'Set Ascending to True to sort ascending,
'false to sort descending

    Dim i As Integer
    Dim j As Integer
    Dim NumInArray, LowerBound As Integer
    NumInArray = UBound(nums)
    LowerBound = LBound(nums)
    For i = LowerBound To NumInArray
        j = 0
        For j = LowerBound To NumInArray
            If Ascending = True Then
                If nums(i) < nums(j) Then
                    NumSwap nums(i), nums(j)
                End If
            Else
                If nums(i) > nums(j) Then
                    NumSwap nums(i), nums(j)
                End If
            End If
        Next j
    Next i
End Sub

Private Sub NumSwap(var1 As Variant, var2 As Variant)
    Dim x As Variant
    x = var1
    var1 = var2
    var2 = x
End Sub

Public Function ConvertDecToBaseN(ByVal dValue As Double) As String
Dim byBase As Integer
Const BASENUMBERS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
 
Dim sResult As String
Dim dRemainder As Double
 
On Error GoTo errorhandler
byBase = 36
sResult = ""
If (byBase > 2) And (byBase < 37) Then
  dValue = Abs(dValue)
  Do
    dRemainder = dValue - (byBase * Int((dValue / byBase)))
    sResult = Mid$(BASENUMBERS, dRemainder + 1, 1) & sResult
    dValue = Int(dValue / byBase)
  Loop While (dValue > 0)
End If
ConvertDecToBaseN = sResult
Exit Function
 
errorhandler:
 
  Err.Raise Err.Number, "ConvertDecTobaseN", Err.Description
    
End Function
Public Function Dec2Bin(DeciNumber As Double, Optional Padnumber As Integer) As String
    Dim Number As Double
    Dim Binary_String As String
    Dim PadString As String
    Number = DeciNumber
    Binary_String = ""
    
    Do While Number > 0
        Binary_String = Number Mod 2 & Binary_String
        Number = Number \ 2
    Loop
    If Padnumber > 0 Then
         PadString = String(Padnumber, "0")
         PadString = PadString & Binary_String
         Binary_String = Right(PadString, Padnumber)
    End If
    
    Dec2Bin = Binary_String
    
End Function

Public Function NetWorkByteOrder(ByVal Number As Double) As String
Dim TMPStr As String
Dim OutStr As String

TMPStr = Dec2Bin(Number, 32)
OutStr = Chr(Bin2Dec(Mid(TMPStr, 1, 8)))
OutStr = OutStr & Chr(Bin2Dec(Mid(TMPStr, 9, 8)))
OutStr = OutStr & Chr(Bin2Dec(Mid(TMPStr, 17, 8)))
OutStr = OutStr & Chr(Bin2Dec(Mid(TMPStr, 25, 8)))
NetWorkByteOrder = OutStr

End Function


Function Bin2Dec(num As String) As Double
  Dim n As Integer
  Dim A As Integer
     n = Len(num) - 1
     A = n
     Do While n > -1
        x = Mid(num, ((A + 1) - n), 1)
        Bin2Dec = IIf((x = "1"), Bin2Dec + (2 ^ (n)), Bin2Dec)
        n = n - 1
     Loop
End Function

Public Function HexToDecimal(HexVal As String) As Double
    Dim binVal As String
    Dim decVal As Double
    
    binVal = HextoBinary(HexVal)
    
    decVal = Bin2Dec(binVal)
    HexToDecimal = decVal
End Function

Public Function HextoBinary(HexVal As String) As String
    Dim binVal As String
    If Len(HexVal) = 1 Then
        binVal = Hex_MAP_BIN(HexVal)
        If binVal <> "" Then
            HextoBinary = binVal
            'Debug.Print HexVal & "-" & binVal
        Else
            Exit Function
        End If
    Else
    
         'Debug.Print Mid(HexVal, Len(HexVal), 1)
         'Debug.Print Mid(HexVal, 1, Len(HexVal) - 1)
        binVal = HextoBinary(Mid(HexVal, 1, Len(HexVal) - 1)) & HextoBinary(Mid(HexVal, Len(HexVal), 1))
        HextoBinary = binVal
       
    End If
    
    
    
End Function


Public Function Hex_MAP_BIN(HexVal As String) As String
    Select Case UCase(HexVal)
        Case "0"
            Hex_MAP_BIN = "0000"
        Case "1"
            Hex_MAP_BIN = "0001"
        Case "2"
            Hex_MAP_BIN = "0010"
        Case "3"
            Hex_MAP_BIN = "0011"
        Case "4"
            Hex_MAP_BIN = "0100"
        Case "5"
            Hex_MAP_BIN = "0101"
        Case "6"
            Hex_MAP_BIN = "0110"
        Case "7"
            Hex_MAP_BIN = "0111"
        Case "8"
            Hex_MAP_BIN = "1000"
        Case "9"
            Hex_MAP_BIN = "1001"
        Case "A"
            Hex_MAP_BIN = "1010"
        Case "B"
            Hex_MAP_BIN = "1011"
        Case "C"
            Hex_MAP_BIN = "1100"
        Case "D"
            Hex_MAP_BIN = "1101"
        Case "E"
            Hex_MAP_BIN = "1110"
        Case "F"
            Hex_MAP_BIN = "1111"
        Case Else
        
    End Select
    
End Function



Public Function FromNumBase(ByVal num As String, ByVal base As Integer) As String
    Dim aMay As Integer
    Dim aMin As Integer
    aMin = AscW("a") - 10
    aMay = AscW("A") - 10
    Dim i As Integer
    Dim n As Double
    Dim strC As String
    Dim j As Integer
    Dim k As Integer
    
    i = 0
    n = 0
    
    Do While True
        If Left(num, 1) = "0" Then
            num = Mid(num, 2)
        Else
            Exit Do
        End If
    Loop

    For j = Len(num) To 1 Step -1
        strC = Mid(num, j, 1)
        Select Case strC
            Case "0"
                i = i + 1
            Case " "
                ' nada
            Case "1" To "9"
                k = CInt(strC)
                If k - base >= 0 Then
                    'Continue For
                Else
                    n = n + CDbl(k * (base ^ i))
                    i = i + 1
                End If
            Case "A" To "Z"
                k = AscW(strC) - aMay
                If k - base >= 0 Then
                    'Continue For
                Else
                    n = n + CDbl(k * (base ^ i))
                    i = i + 1
                End If
            Case "a" To "z"
                k = AscW(strC) - aMin
                If k - base >= 0 Then
                    'Continue For
                Else
                    n = n + CDbl(k * (base ^ i))
                    i = i + 1
                End If
        End Select
    Next
    FromNumBase = CStr(n)
End Function

Public Function ShiftIt(ByVal InPutStr As String, Places As Integer, Dir As Integer) As String
Dim L As Long
Dim aAsc As Integer
Dim bAsc As Integer
Dim tTemp As String
Dim OutString As String
Dim Z As Integer


   On Error GoTo ShiftIt_Error

Z = Len(InPutStr)
If Z = 0 Then
    ShiftIt = ""
    Exit Function
End If
OutString = String(Z, " ")
If Dir = 1 Then
    For L = 1 To Z
        aAsc = Asc(Mid(InPutStr, L, 1))
        bAsc = aAsc + Places
        If bAsc > 255 Then bAsc = Abs(aAsc - 255)
        Mid(OutString, L, 1) = Chr(bAsc)
    Next
Else
    For L = 1 To Z
        aAsc = Asc(Mid(InPutStr, L, 1))
        bAsc = aAsc - Places
        If bAsc < 0 Then bAsc = Abs(255 - aAsc)
        Mid(OutString, L, 1) = Chr(bAsc)
    Next
End If
ShiftIt = OutString

   On Error GoTo 0
   Exit Function

ShiftIt_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure ShiftIt of Module modFalconWorks"
End Function

Public Function QtStr(InPutStr As String) As String
QtStr = Chr(34) & InPutStr & Chr(34)
End Function


Public Function BinarySearch(ByRef Arr() As String, ByRef Search As String) As Long
Dim First As Long
Dim Last As Long
Dim Middle As Long
First = LBound(Arr)
Last = UBound(Arr)
Debug.Print "Binary Searching"
Do
    Middle = (First + Last) \ 2
    Select Case StrComp(Arr(Middle), Search, vbBinaryCompare)
        Case -1: First = Middle + 1
        Case 1: Last = Middle - 1
        Case 0
            BinarySearch = Middle
            Exit Function
        End Select
Loop Until First > Last
End Function

Public Function Bool2Word(Value As Boolean) As String
If Value = True Then
    Bool2Word = "TRUE"
Else
    Bool2Word = "FALSE"
End If
End Function

Public Function Word2Bool(Value As String) As Boolean
Select Case UCase(Value)
    Case "YES", "TRUE", "1"
        Word2Bool = True
    Case Else
        Word2Bool = False
End Select
End Function

Public Function ShowColorDialog(Optional ByVal hParent As Long, _
    Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) _
    As Long
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With CC
        .lStructSize = Len(CC)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(CC) Then
        'if not canceled, return the color
        ShowColorDialog = CC.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If
End Function

Public Sub WriteINI(Header As String, keyname As String, Value As String, Optional INIFName As String)
    Dim U As Long
    Dim OpenName As String
    If INIFName = "" Then
        OpenName = INIFileName
    Else
        OpenName = INIFName
    End If
    U = WritePrivateProfileString(Header, keyname, Value, OpenName)
    If U = 0 Then
    End If
End Sub
'
'Public Function ReadINI(Header As String, keyname As String, sINIFileName As String) As String
'    Dim sRet As String
'    sRet = String(255, Chr(0))
'    ReadINI = Left(sRet, GetPrivateProfileString(Header, ByVal keyname, "", sRet, Len(sRet), sINIFileName))
'End Function

Public Sub InfoLog(ByVal inText As String, LogFile As String)
Dim Handle As Integer
Dim Header As String
Dim OutText As String

Header = Format(Now, "<YYYY-DD-mmm HH:MM:SS>")
OutText = inText
Handle = FreeFile

Open LogFile For Append Shared As Handle
Print #Handle, OutText
Close Handle

End Sub

Public Function FileExten(ByVal FileName As String) As String
Dim A As Integer
Dim B As Integer

A = Len(FileName)
B = InStrRev(FileName, ".")
FileExten = Right(FileName, (A + 1) - B)
End Function


Public Function SizeFileStr(MySize As Variant) As String
    Dim fSize As Double
    Dim TheSize As Double
    Dim DoubleBytes As Double
    
   On Error GoTo SizeFileStr_Error

    On Error Resume Next
    fSize = MySize
    

    If fSize = 0 Then
        SizeFileStr = "00KB"
        Exit Function
    End If
    TheSize = fSize

Select Case TheSize
            Case Is >= 1099511627776#
                DoubleBytes = CDbl(TheSize / 1099511627776#) 'TB
                SizeFileStr = Format(DoubleBytes, "Standard") & "TB"
            Case 1073741824 To 1099511627775#
                DoubleBytes = CDbl(TheSize / 1073741824) 'GB
                SizeFileStr = Format(DoubleBytes, "Standard") & "GB"
            Case 1048576 To 1073741823
                DoubleBytes = CDbl(TheSize / 1048576) 'MB
                SizeFileStr = Format(DoubleBytes, "Standard") & "MB"
            Case 1024 To 1048575
                DoubleBytes = CDbl(TheSize / 1024) 'KB
                SizeFileStr = Format(DoubleBytes, "Standard") & "KB"
            Case 0 To 1023
                DoubleBytes = TheSize ' bytes
                SizeFileStr = Format(DoubleBytes, "Standard") & " bytes"
            Case Else
                SizeFileStr = "00KB"
        End Select

   On Error GoTo 0
   Exit Function

SizeFileStr_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure SizeFileStr of Module modFalconWorks"
End Function

Public Sub LogErrorEx(ByVal Message As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\DCC_DebugLog.txt" For Append As #f
        Print #f, Now & " - " & Message
    Close #f
End Sub


