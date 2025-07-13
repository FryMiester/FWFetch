VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKillMe 
      Interval        =   1000
      Left            =   5520
      Top             =   240
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDisplayText 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Shape Shape 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DieTime As Integer

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtDisplayText.Visible = True
txtDisplayText.Enabled = True
If txtDisplayText.Visible = True And txtDisplayText.Enabled = True Then
    txtDisplayText.SetFocus
End If

End Sub

Private Sub Form_Load()
PutOnTop Me.hWnd
DieTime = 5

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
TakeOffTop Me.hWnd
End Sub

Private Sub tmrKillMe_Timer()
Static loops As Integer
loops = loops + 1
If loops = DieTime Then
    loops = 0
    Unload Me
End If
End Sub

Private Sub txtDisplayText_Change()
Me.Caption = App.ProductName & " Message Box"
txtDisplayText.Visible = True
txtDisplayText.Enabled = True
If txtDisplayText.Visible = True And txtDisplayText.Enabled = True Then
    txtDisplayText.SetFocus
End If
'txtDisplayText.SelStart = 0
'txtDisplayText.SelLength = Len(txtDisplayText.text)

End Sub

Private Sub txtTimeOut_Change()
If Val(txtTimeOut.text) <> 0 Then DieTime = Val(txtTimeOut.text)
If txtTimeOut.text = "0" Then tmrKillMe.Enabled = False
End Sub
