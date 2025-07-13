VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Tray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   DrawStyle       =   1  'Dash
   DrawWidth       =   2
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   159
   ToolboxBitmap   =   "Tray.ctx":0000
   Begin VB.Timer Timer1 
      Left            =   1755
      Top             =   150
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   765
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=============
' From  www.solution-shelf.com
' Free Tray Icon  .CTL  component
' adapted from KB Article ID: Q162613
' add to your Templates\UserCtls  directory
' modified from earlier releases for clarity -- TL Price

'**********************************
' USE ENTIRELY AT YOUR OWN RISK.
'**********************************

' this control asserts Separate -or- Combined MouseClick events
Enum EventModeType
    Separate_EVENTS = 0
    Combined_EVENT
End Enum

' Custom EVENTS from this Control
'-----------------
' S E P A R A T E
Event LeftClick()
Event RightClick()
' C O M B I N E D
Event MouseClick(Button As Long)
Event MouseHover()


Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' ShellNotify UDT
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' ShellNotify Message Constants
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

' Mouse MSG Constants
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_LBUTTONDOWN = &H201

' Mouse Button ID constants
Private Const LEFTBUTTON = 1
Private Const RIGHTBUTTON = 2

' flag to help  -filter-  messages after event detection
Dim msgHoldOff As Boolean

' flag to determine UDT create-overwrite-delete actions
Private TrayIconIsLoaded As Boolean

' UDT var for interaction with shell
Private TrayIconNID As NOTIFYICONDATA

'Default Property Values:
Const m_def_EventMode = Separate_EVENTS
Const m_def_TrayToolTip = " FalconWorks SoftWare "

'Property Variables:
Dim m_EventMode As Long
Dim m_TrayToolTip As String

Public Sub Show(UpDateData As Boolean)
Attribute Show.VB_Description = "Method -- Sets tray-image as visible or not via boolean arg."
' ====================================
' Loads, Modifies, or UnLoads TrayIcon UDT var

    If UpDateData Then
        TrayIconNID.cbSize = Len(TrayIconNID)
        TrayIconNID.hWnd = UserControl.hWnd
        TrayIconNID.uId = vbNull
        TrayIconNID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ' assign notify callback message to use
        TrayIconNID.uCallBackMessage = WM_MOUSEMOVE
        '-MUST-  hold a valid icon (else you'll see nothing)
        TrayIconNID.hIcon = UserControl.Picture
        '
        If TrayToolTip <> "" Then
            ' I make the string have leading and trailing SPACE chars -- for readability
            TrayIconNID.szTip = " " & TrayToolTip & " " & Chr$(0)
        Else
            ' if no text then no spaces
            TrayIconNID.szTip = Chr$(0)
        End If
        '
        If TrayIconIsLoaded Then
            ' already loaded, just modify
            Shell_NotifyIcon NIM_MODIFY, TrayIconNID
        Else
            ' not loaded yet -- add it
            Shell_NotifyIcon NIM_ADD, TrayIconNID
            TrayIconIsLoaded = True
        End If
        '
    Else
        ' ==  time to clean up and get outta here
        If TrayIconIsLoaded Then
            Shell_NotifyIcon NIM_DELETE, TrayIconNID
            TrayIconIsLoaded = False
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
' =========================================================================================
    Dim Msg As Long

    If msgHoldOff = True Then
        Exit Sub    ' STABILIZER, see below
    End If

    ' MUST have Scalemode = 3 = VbPixel, or divide X by  screen.TwipsPerPixelX
    Msg = CLng(X)
    '
    Select Case Msg
    Case WM_LBUTTONDOWN
        ' I use the DOWN messages, you can use the UP messages if you prefer

        ' set flag to ignore any additional messages
        msgHoldOff = True
        If m_EventMode = Separate_EVENTS Then
            RaiseEvent LeftClick
        Else
            ' use combined event  @LEFT
            RaiseEvent MouseClick(LEFTBUTTON)
        End If

        ' wait mSec before processing any new messages -- STABILIZE
        Timer1.Interval = 15

    Case WM_RBUTTONDOWN:
        ' -----------------------------
        ' set flag to ignore any additional messages
        msgHoldOff = True
        If m_EventMode = Separate_EVENTS Then
            RaiseEvent RightClick
        Else
            ' use combined event  @RIGHT
            RaiseEvent MouseClick(RIGHTBUTTON)
        End If
        ' wait mSec before processing any new messages -- STABILIZE
        Timer1.Interval = 15
        ' Case WM_MOUSEMOVE
        '  msgHoldOff = True
        '  RaiseEvent MouseHover
        ' Timer1.Interval = 600
    End Select
End Sub

Private Sub Timer1_Timer()
' ========================   we have waited -N- mSec
    Timer1.Interval = 0
    msgHoldOff = False
End Sub

Private Sub UserControl_Resize()
' =============================
    UserControl.Size 36 * Screen.TwipsPerPixelX, 36 * Screen.TwipsPerPixelY
    UserControl.Refresh
End Sub

Public Property Get TrayImage() As Picture
Attribute TrayImage.VB_Description = "Returns/sets a graphic to be displayed  on the system Tray. (Bitmaps and Cursors are converted to type Icon)"
Attribute TrayImage.VB_MemberFlags = "200"
' ======================================
    Set TrayImage = UserControl.Picture
End Property

Public Property Set TrayImage(ByVal New_TrayImage As Picture)
' ===========================================================
    If Not New_TrayImage Is Nothing Then
        ' got Icon ?
        If New_TrayImage.Type = vbPicTypeIcon Then
            Set UserControl.Picture = New_TrayImage
        Else
            ' no -- then we'll attempt to create one      (can't handle meta-files)
            UserControl.ImgList.ImageWidth = 32
            UserControl.ImgList.ImageHeight = 32
            UserControl.ImgList.ListImages.Add 1, "dummy", New_TrayImage
            Set UserControl.Picture = UserControl.ImgList.ListImages(1).ExtractIcon
            UserControl.ImgList.ListImages.Clear
        End If
    Else
        Set UserControl.Picture = Nothing
    End If
    PropertyChanged "TrayImage"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' ===========================================================
    Set Picture = PropBag.ReadProperty("TrayImage", Nothing)
    m_TrayToolTip = PropBag.ReadProperty("TrayToolTip", m_def_TrayToolTip)
    m_EventMode = PropBag.ReadProperty("EventMode", m_def_EventMode)
    ' force this scalemode to simplify code
    UserControl.ScaleMode = 3
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' =============================================================
    Call PropBag.WriteProperty("TrayImage", Picture, Nothing)
    Call PropBag.WriteProperty("TrayToolTip", m_TrayToolTip, m_def_TrayToolTip)
    Call PropBag.WriteProperty("EventMode", m_EventMode, m_def_EventMode)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
' =====================================
    m_TrayToolTip = m_def_TrayToolTip
    m_EventMode = m_def_EventMode
End Sub


Public Property Get TrayToolTip() As String
Attribute TrayToolTip.VB_Description = "Sets/Returns text displayed when mouse over tray icon."
' ========================================
    TrayToolTip = m_TrayToolTip
End Property
Public Property Let TrayToolTip(ByVal New_TrayToolTip As String)
' ==============================================================
    m_TrayToolTip = New_TrayToolTip
    PropertyChanged "TrayToolTip"
End Property


Public Property Get EventMode() As EventModeType
Attribute EventMode.VB_Description = "Sets whether Mouse clicks are combined in a single MouseClick Event or are separated into RightClick / LeftClick Events."
Attribute EventMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute EventMode.VB_UserMemId = 0
' ===============================================
'If Ambient.UserMode Then Err.Raise 393
    EventMode = m_EventMode
End Property
Public Property Let EventMode(ByVal New_EventMode As EventModeType)
' ===================================================================
    m_EventMode = New_EventMode
    If EventMode > Combined_EVENT Then EventMode = Combined_EVENT
    If EventMode < 0 Then EventMode = 0
    PropertyChanged "EventMode"
End Property


Public Property Get IsVisible() As Boolean
Attribute IsVisible.VB_Description = "Flags whether icon is loaded and visible in Tray -- or NOT."
Attribute IsVisible.VB_MemberFlags = "400"
' =====================================
    IsVisible = TrayIconIsLoaded
End Property


