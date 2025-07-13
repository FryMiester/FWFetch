VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "FWfetch - About"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Image Image 
      Height          =   4710
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   2460
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim showText As String
Label1.Caption = App.ProductName & " version " & App.Major & "." & App.Minor & "." & App.Revision
showText = "FWFetch is a utility program designed to make downloading files from IRC servers much easier than is currently possible. "
showText = showText & "FWFetch watches the channel for file list announcements and then grabs those lists. It has a built-in search engine "
showText = showText & " that operates almost exactly the same way as the search engines found in #ebooks and many others. This search engine allows you to filter "
showText = showText & "the results with a not operator... 'Clive Cussler -Dirk'. You can do your searching on your own time, then mark files for downloading, now, or later. "
showText = showText & vbCrLf & vbCrLf
showText = showText & "This is NEW software, for issues, comments, suggestions or gripes, contact FryMiester in #ebooks on IRCHighway with your regular IRC client."


Label2.Caption = showText


End Sub
