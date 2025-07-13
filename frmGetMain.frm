VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetMain 
   Caption         =   "FWFetch"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   16095
   Icon            =   "frmGetMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   16095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowRequest 
      Caption         =   "Show Requests"
      Height          =   240
      Left            =   13515
      TabIndex        =   99
      Top             =   420
      Width           =   1710
   End
   Begin VB.CheckBox chkShowSearch 
      Caption         =   "Show Searches"
      Height          =   225
      Left            =   13530
      TabIndex        =   98
      Top             =   120
      Width           =   1500
   End
   Begin VB.ListBox lstRevert 
      Height          =   255
      Left            =   15075
      TabIndex        =   95
      Top             =   2535
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.ListBox lstCollector 
      Height          =   255
      Left            =   15060
      TabIndex        =   94
      Top             =   2205
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstGetList 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3555
      MultiSelect     =   2  'Extended
      TabIndex        =   92
      ToolTipText     =   "Double Click To Remove, Right Click to Paste"
      Top             =   4380
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdShowGetList 
      Caption         =   "Show Get List"
      Height          =   375
      Left            =   2610
      TabIndex        =   91
      Top             =   7410
      Width           =   1695
   End
   Begin VB.Timer tmrOffLine 
      Interval        =   5000
      Left            =   3795
      Top             =   2895
   End
   Begin VB.ListBox lstOffLine 
      Height          =   255
      Left            =   135
      TabIndex        =   89
      Top             =   6990
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "USE TRANSFERS SCHEDULE"
      Height          =   780
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   915
      Width           =   1200
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Private Chat"
      Height          =   495
      Index           =   5
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Configuration"
      Height          =   495
      Index           =   4
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Browse File Lists"
      Height          =   495
      Index           =   3
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Server Maint."
      Height          =   495
      Index           =   2
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Search"
      Height          =   495
      Index           =   1
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Chat"
      Height          =   495
      Index           =   0
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   105
      Width           =   1455
   End
   Begin VB.Frame FrameMain 
      Height          =   975
      Index           =   5
      Left            =   15435
      TabIndex        =   70
      Top             =   975
      Visible         =   0   'False
      Width           =   855
      Begin VB.Frame FramePrivateChat 
         Height          =   7335
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   10455
         Begin VB.TextBox txtPrivateChat 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   6360
            Width           =   10215
         End
         Begin VB.CommandButton cmdCloseChat 
            Caption         =   "Close"
            Height          =   375
            Left            =   9240
            TabIndex        =   72
            Top             =   6840
            Width           =   975
         End
         Begin RichTextLib.RichTextBox RTBPrivate 
            Height          =   6015
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   10610
            _Version        =   393217
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmGetMain.frx":27852
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   615
      Index           =   3
      Left            =   13890
      TabIndex        =   67
      Top             =   5370
      Visible         =   0   'False
      Width           =   720
      Begin VB.Frame FrameBrowseList 
         Height          =   7455
         Left            =   600
         TabIndex        =   68
         Top             =   480
         Width           =   11175
         Begin MSComctlLib.ListView LVBrowse 
            Height          =   7095
            Left            =   120
            TabIndex        =   69
            Top             =   225
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   12515
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   870
      Index           =   2
      Left            =   13860
      TabIndex        =   62
      Top             =   6330
      Visible         =   0   'False
      Width           =   930
      Begin VB.Frame frameServers 
         Caption         =   "Servers"
         Height          =   7395
         Left            =   795
         TabIndex        =   63
         Top             =   810
         Width           =   10245
         Begin VB.CommandButton cmdServerPause 
            Caption         =   "Pause Server Maintenance"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   165
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   5145
            Width           =   1845
         End
         Begin MSComctlLib.ListView LVServers 
            Height          =   4815
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "Double Click to Get List,  Right Click For Options"
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   8493
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin RichTextLib.RichTextBox sysInfo 
            Height          =   1215
            Left            =   120
            TabIndex        =   65
            Top             =   6000
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2143
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmGetMain.frx":278CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblListCount 
            Height          =   495
            Index           =   1
            Left            =   3360
            TabIndex        =   66
            Top             =   5160
            Width           =   2775
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   1980
      Index           =   1
      Left            =   13590
      TabIndex        =   54
      Top             =   3000
      Visible         =   0   'False
      Width           =   2145
      Begin VB.Frame FrameSearch 
         Height          =   5070
         Left            =   -4650
         TabIndex        =   55
         Top             =   135
         Width           =   10470
         Begin VB.ComboBox cboOldSearches 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IntegralHeight  =   0   'False
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   97
            Text            =   "cboOldSearches"
            Top             =   270
            Width           =   2970
         End
         Begin VB.CommandButton cmdReScan 
            Caption         =   "WITHIN RESULTS"
            Height          =   420
            Left            =   8265
            TabIndex        =   93
            Top             =   255
            Width           =   1095
         End
         Begin VB.CommandButton cmdSearchAbort 
            Caption         =   "ABORT"
            Height          =   420
            Left            =   9570
            TabIndex        =   90
            Top             =   255
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddSelected 
            Caption         =   "Add Selected To Get List"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   6720
            Width           =   3255
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "SEARCH"
            Height          =   420
            Left            =   7095
            TabIndex        =   57
            Top             =   255
            Width           =   1095
         End
         Begin VB.TextBox txtSearch 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3555
            TabIndex        =   56
            Top             =   300
            Width           =   3510
         End
         Begin MSComctlLib.ListView lvTitles 
            Height          =   5775
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   10186
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.TreeView TV1 
            Height          =   1455
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2566
            _Version        =   393217
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblListCount 
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
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   61
            Top             =   6480
            Width           =   3135
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   630
      Index           =   0
      Left            =   14130
      TabIndex        =   47
      Top             =   1695
      Width           =   1020
      Begin VB.Frame FrameIRC 
         Caption         =   "ChatMode"
         Height          =   7335
         Left            =   -240
         TabIndex        =   48
         Top             =   -75
         Width           =   10815
         Begin VB.CommandButton cmdIRCSearch 
            BackColor       =   &H00C0C0FF&
            Caption         =   "@Search"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Right Click to cycle through available SearchBots"
            Top             =   6840
            Width           =   1575
         End
         Begin VB.TextBox txtChat 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   51
            Top             =   6840
            Width           =   7335
         End
         Begin VB.ListBox lstUsers 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5910
            Left            =   9120
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh List"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   49
            Top             =   6720
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox RTBChat 
            Height          =   6495
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   11456
            _Version        =   393217
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmGetMain.frx":2794A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblUserCnt 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9555
            TabIndex        =   82
            Top             =   6240
            Width           =   870
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   7545
      Index           =   4
      Left            =   4605
      TabIndex        =   24
      Top             =   735
      Visible         =   0   'False
      Width           =   10560
      Begin VB.Frame FrameConfiguration 
         Height          =   6990
         Left            =   705
         TabIndex        =   25
         Top             =   240
         Width           =   8850
         Begin VB.CheckBox chkToTray 
            Caption         =   "Minimize To System Tray"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   345
            TabIndex        =   104
            Top             =   6300
            Width           =   2865
         End
         Begin VB.TextBox txtListExpire 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3225
            TabIndex        =   102
            Text            =   "3"
            Top             =   5760
            Width           =   660
         End
         Begin VB.CommandButton cmdAddIgnore 
            Caption         =   "Add To Ignore List"
            Height          =   240
            Left            =   5550
            TabIndex        =   101
            Top             =   5550
            Width           =   2085
         End
         Begin VB.TextBox txtAddIgnore 
            Height          =   285
            Left            =   5535
            TabIndex        =   100
            Top             =   5235
            Width           =   2130
         End
         Begin VB.TextBox txtSchedStop 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3210
            TabIndex        =   88
            Text            =   "00:00"
            Top             =   5385
            Width           =   705
         End
         Begin VB.CheckBox chkSchedule 
            Caption         =   "Use Schedule For Transfers"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   85
            Top             =   4665
            Width           =   4710
         End
         Begin VB.TextBox txtSchedTime 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   84
            Text            =   "00:00"
            Top             =   5010
            Width           =   675
         End
         Begin VB.TextBox txtChanName 
            Height          =   285
            Left            =   360
            TabIndex        =   38
            Text            =   "#ebooks"
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtNetWorkName 
            Height          =   285
            Left            =   360
            TabIndex        =   37
            Text            =   "IRCHighway"
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtNetworkAddr 
            Height          =   285
            Left            =   360
            TabIndex        =   36
            Text            =   "IRC.IRCHighway.net"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtBotNick 
            Height          =   285
            Left            =   360
            TabIndex        =   35
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtNickServPass 
            Height          =   285
            Left            =   360
            TabIndex        =   34
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtDownloadFolder 
            Height          =   285
            Left            =   360
            TabIndex        =   32
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Settings"
            Height          =   375
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   6000
            Width           =   1575
         End
         Begin VB.TextBox txtNeworkPort 
            Height          =   285
            Left            =   1680
            TabIndex        =   30
            Text            =   "6667"
            Top             =   1560
            Width           =   975
         End
         Begin VB.Timer tmrDirty 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   8040
            Top             =   240
         End
         Begin VB.CheckBox chkFilterRequests 
            Caption         =   "Filter Requests In Chat Window"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   3840
            Width           =   3855
         End
         Begin VB.CheckBox chkFilterSearches 
            Caption         =   "Filter Searches In Chat Window"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   4200
            Width           =   3975
         End
         Begin VB.ListBox lstIgnores 
            Height          =   4545
            Left            =   5520
            TabIndex        =   27
            ToolTipText     =   "Double Click To Remove"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox chkAutoConnect 
            Caption         =   "Auto Connect"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   26
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label lblStuff 
            Caption         =   "File List Expiry Days"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   330
            TabIndex        =   103
            Top             =   5850
            Width           =   2760
         End
         Begin VB.Label lblInfo 
            Caption         =   "Stop Scheduled Transfers at "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   345
            TabIndex        =   87
            Top             =   5430
            Width           =   2715
         End
         Begin VB.Label lblInfo 
            Caption         =   "Start Scheduled Transfers at "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   345
            TabIndex        =   86
            Top             =   5025
            Width           =   2715
         End
         Begin VB.Label lblInfo 
            Caption         =   "Channel Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   46
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblInfo 
            Caption         =   "Network Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   45
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblInfo 
            Caption         =   "Network Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   44
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label lblInfo 
            Caption         =   "Bot Nickname"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   43
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label lblInfo 
            Caption         =   "Bot NickServ Password"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   42
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label lblInfo 
            Caption         =   "Download Folder"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   41
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label lblInfo 
            Caption         =   "Network Port"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   2760
            TabIndex        =   40
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lblJunk 
            Alignment       =   2  'Center
            Caption         =   "Ignored"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   690
      Picture         =   "frmGetMain.frx":279C6
      ScaleHeight     =   2775
      ScaleWidth      =   2775
      TabIndex        =   23
      Top             =   1110
      Width           =   2775
   End
   Begin VB.FileListBox File2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdTreeViewToggle 
      Caption         =   "List View Active"
      Height          =   375
      Left            =   2610
      TabIndex        =   20
      Top             =   7875
      Width           =   1695
   End
   Begin VB.Timer tmrUpGrade 
      Interval        =   1000
      Left            =   3000
      Top             =   3360
   End
   Begin VB.ListBox lstResults 
      Height          =   255
      Left            =   10680
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin FWFetch.ctlSVRDownLoad FLGetter 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6585
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin FWFetch.ctlDownLoad fGetter 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin FWFetch.ctlDownLoad fGetter 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin FWFetch.ctlDownLoad fGetter 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin FWFetch.ctlDownLoad fGetter 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin FWFetch.ctlDownLoad fGetter 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   3735
      _ExtentX        =   7858
      _ExtentY        =   661
   End
   Begin MSWinsockLib.Winsock L0sckDCC 
      Left            =   2640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogFile 
      Caption         =   "Open Transfer Log"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "PAUSE TRANSFERS"
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   500
      Left            =   2640
      Top             =   3360
   End
   Begin VB.Timer tmrGetNewBook 
      Interval        =   1000
      Left            =   3360
      Top             =   3360
   End
   Begin VB.Timer tmtKeepAlive 
      Interval        =   1000
      Left            =   3000
      Top             =   2880
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2250
      TabIndex        =   4
      Top             =   645
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrListMaint 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2640
      Top             =   2880
   End
   Begin VB.ListBox lstGetNewList 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrServerMaint 
      Interval        =   20000
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Timer tmrDoBuff 
      Interval        =   500
      Left            =   3360
      Top             =   2880
   End
   Begin VB.ListBox lstDialogBuff 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin FWFetch.Tray Tray 
      Left            =   2040
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   953
      TrayImage       =   "frmGetMain.frx":3D408
   End
   Begin RichTextLib.RichTextBox RTBLog 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmGetMain.frx":64C6A
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "CONNECT"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblGetList 
      Caption         =   "Get List Count = 0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   96
      Top             =   7125
      Width           =   1695
   End
   Begin VB.Label lblBrowseInfo 
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
      Height          =   855
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblSpy 
      Height          =   255
      Left            =   4095
      TabIndex        =   19
      Top             =   3195
      Width           =   375
   End
   Begin VB.Label lblBadcount 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4095
      TabIndex        =   10
      Top             =   2835
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Failed Transfers"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3135
      TabIndex        =   9
      Top             =   2595
      Width           =   975
   End
   Begin VB.Label lblGoodCount 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4095
      TabIndex        =   8
      Top             =   2235
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Sucessful Transfers"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3135
      TabIndex        =   7
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label lblMessages 
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
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   4800
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuServersPopup 
      Caption         =   "Servers"
      Visible         =   0   'False
      Begin VB.Menu mnu_Disable 
         Caption         =   "Toggle Disabled"
      End
      Begin VB.Menu mnu_DeleteServer 
         Caption         =   "Delete This Server"
      End
   End
   Begin VB.Menu mnuGetList 
      Caption         =   "GetList"
      Visible         =   0   'False
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste From Clipboard"
      End
   End
   Begin VB.Menu mnu_About1 
      Caption         =   "About"
      Begin VB.Menu mnu_about2 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmGetMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tmpUser As UserPackage
Dim Rnd As Integer
Dim AmConfigured As Boolean
Dim InChatWith As String
Dim OpenFile As String
Dim FileLines() As String
Dim CurrLine As Long
Dim PassThrough As Boolean
Dim ScrollBack As New Collection
Dim ScrollSearch As New Collection
Dim ScrollNum As Integer
Dim ScrollSearchNum As Integer
Dim UpgradeLoops As Integer
Dim ThisServerInfo As ListServerInfo
Dim SearchMODE As Boolean
Dim SVRPause As Boolean
Dim IsActive As Boolean
Dim AmConnected As Boolean
Dim TransferActive(4) As Boolean
Dim ThisServer As clsListServer
Dim RemotePinger As String
Dim tTop As Long
Dim tLeft As Long
Dim tHeight As Long
Dim tWidth As Long
Dim CurrSVR As Integer

Private Sub cboOldSearches_Click()
' Bring Up old search list
Dim IDX As Integer
Dim fName As String
Dim Handle As Integer
Dim cCount As Integer
Dim InString As String

If cboOldSearches.ListIndex = 0 Then
   Exit Sub
End If
If cboOldSearches.ListIndex = 1 Then
   ClearSearches
   
   Exit Sub
End If

frmGetMain.lvTitles.ListItems.Clear
frmGetMain.TV1.Nodes.Clear
frmGetMain.TV1.Nodes.Add , , "root", "Search"

Me.File1.Path = MySearchResultFolder
Me.File1.Pattern = "*.txt"
cCount = Me.File1.ListCount
IDX = Me.cboOldSearches.ListIndex - 1
fName = MySearchResultFolder & Me.File1.List(IDX)


Handle = FreeFile
Open fName For Input As Handle

Do Until EOF(Handle)
    Line Input #Handle, InString
    If Left(InString, 1) = "!" Then
        AddTitleLV frmGetMain.lvTitles, InString
        DoTreeview frmGetMain.TV1, InString
    End If
Loop
Close Handle

End Sub
Private Sub ClearSearches()
Dim Ask As Long
Dim cCount As Integer
Dim L As Long
Dim tText As String

cCount = cboOldSearches.ListCount
If cCount < 3 Then
   cboOldSearches.ListIndex = 0
   Exit Sub
End If
Ask = VBA.MsgBox("This will DELETE all saved search results, Are You Sure?", vbYesNoCancel)
If Ask <> vbYes Then Exit Sub
Me.File1.Path = MySearchResultFolder
Me.File1.Pattern = "*.txt"
cCount = Me.File1.ListCount
For L = 0 To cCount - 1
   tText = Me.File1.List(L)
   tText = Me.File1.Path & "\" & tText
   If FileExists(tText) = True Then
      Kill tText
   End If
Next
Me.cboOldSearches.Clear
LoadCBO
End Sub


Private Sub cboOldSearches_KeyUp(KeyCode As Integer, Shift As Integer)
Static tText As String
Dim Ask As Long

If KeyCode = 46 Then
   'KeyCode = vbNull
   Debug.Print "BOO!"
   Ask = VBA.MsgBox("Delete The Search " & tText, vbYesNoCancel)
   
End If

If Ask = vbYes Then

End If

tText = cboOldSearches.SelText

Debug.Print tText
Debug.Print KeyCode


End Sub

Private Sub chkShowRequest_Click()
If chkShowRequest.Value = 1 Then
   chkFilterRequests.Value = 0
Else
   chkFilterRequests.Value = 1
End If
End Sub

Private Sub chkShowSearch_Click()
If chkShowSearch.Value = vbChecked Then
   chkFilterSearches.Value = 0
Else
   chkFilterSearches.Value = 1
End If
End Sub

Private Sub cmdAddIgnore_Click()
' Add to ignore List
Me.lstIgnores.AddItem Me.txtAddIgnore.text
Me.txtAddIgnore.text = ""
WriteIgnores
ReadIgnores
cmdFrame_Click 0
End Sub

Private Sub cmdCloseChat_Click()
RTBPrivate.text = ""
FramePrivateChat.Caption = ""
End Sub

Private Sub cmdFrame_Click(Index As Integer)
If lstGetList.Visible = True Then cmdShowGetList_Click
Dim L As Long
For L = 0 To 5
    FrameMain(L).Visible = False
    cmdFrame(L).BackColor = &H8000000F
Next
FrameMain(Index).Visible = True
cmdFrame(Index).BackColor = &H80000016
If Index = 3 Then
   File2.Path = MyListFolder
   File2.Refresh
   Picture1.Visible = False
   File2.Refresh
Else
   Picture1.Visible = True
   lblBrowseInfo.Visible = False
End If
End Sub

Private Sub cmdIRCSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim SearchBot As String
If AmConnected = False Then Exit Sub
If Button = vbRightButton Then
    
      Rnd = CLng(Val(cmdIRCSearch.Tag))
      Rnd = Rnd + 1
      If Rnd > SearchBots.Count Then Rnd = 1
      SearchBot = NulTrim(SearchBots.Item(Rnd))
      cmdIRCSearch.Caption = SearchBot
      cmdIRCSearch.Tag = CStr(Rnd)

End If
End Sub

Private Sub cmdRefresh_Click()
lstUsers.Clear
sendToIRC "NAMES " & MyChannel
ServerPoll
lblUserCnt.Caption = lstUsers.ListCount
End Sub

Private Sub cmdReScan_Click()
Dim Cnt As Integer
Dim L As Long
Dim G As Integer
Dim Sample As String
Dim Compare As String
Dim SearchStr As String
Dim Terms() As String

   On Error GoTo cmdReScan_Click_Error

lstCollector.Clear

SearchStr = txtSearch.text
Terms = Split(SearchStr, " ")

Cnt = lvTitles.ListItems.Count
For L = 1 To Cnt
      Sample = lvTitles.ListItems(L).SubItems(1)
      For G = 0 To UBound(Terms)
            If InStr(UCase(Sample), UCase(Terms(G))) > 0 Then
                  lstCollector.AddItem Sample
            End If
      Next G
Next L
lvTitles.ListItems.Clear
For L = 0 To lstCollector.ListCount - 1
      AddTitleLV Me.lvTitles, lstCollector.List(L)
Next
cmdSearchAbort.Caption = "UN-Filter"

   On Error GoTo 0
   Exit Sub

cmdReScan_Click_Error:

      DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdReScan_Click of Form frmGetMain", "Debuglog.txt"

End Sub

Private Sub cmdSchedule_Click()
If UseSchedule = False Then
      UseSchedule = True
      cmdSchedule.Caption = UCase("Using Scheduled Transfers")
      cmdSchedule.BackColor = vbGreen
Else
      UseSchedule = False
      cmdSchedule.Caption = UCase("USE TRANSFERS SCHEDULE")
      cmdSchedule.BackColor = &H8000000F
End If
End Sub

Private Sub cmdSearchAbort_Click()

If lvTitles.ListItems.Count > 0 Then
   If cmdSearchAbort.Caption = "Clear" Then
      lvTitles.ListItems.Clear
   End If
   txtSearch.text = ""
   cmdSearchAbort.Caption = "ABORT"
Else
      SearchAbort = True
End If

End Sub

Private Sub cmdServerPause_Click()
If ServersPaused = True Then
      ServersPaused = False
      cmdServerPause.BackColor = &H8000000F
      cmdServerPause.Caption = "Pause Server Maintenance"
Else
      ServersPaused = True
      cmdServerPause.BackColor = &HFFFF&
      cmdServerPause.Caption = "Resume Server Maintenance"
End If
End Sub

Private Sub lstIgnores_DblClick()
Dim Cnt As Integer
Dim L As Long
Cnt = lstIgnores.ListCount
For L = 0 To Cnt - 1
    If lstIgnores.Selected(L) = True Then
        lstIgnores.RemoveItem L
        Exit For
    End If
Next
WriteIgnores

End Sub

Private Sub lstUsers_DblClick()
Dim L As Long
Dim Tmp As String

For L = 0 To lstUsers.ListCount - 1
    If lstUsers.Selected(L) = True Then
        Tmp = lstUsers.List(L)
        Exit For
    End If
Next

If Tmp > "" Then
    Tmp = Replace(Tmp, "~", "")
    Tmp = Replace(Tmp, "&", "")
    Tmp = Replace(Tmp, "@", "")
    Tmp = Replace(Tmp, "%", "")
    Tmp = Replace(Tmp, "+", "")
    InChatWith = Tmp
    Me.FramePrivateChat.Caption = InChatWith
    cmdFrame_Click 5
End If
    

End Sub

Private Sub mnuPaste_Click()
Dim tText As String
Dim inLines() As String
Dim Cnt As Integer
Dim L As Long
tText = Clipboard.Gettext
If tText = "" Then Exit Sub
inLines = Split(tText, vbCrLf)
Cnt = UBound(inLines)
For L = 0 To Cnt
    If Left(Trim(inLines(L)), 1) = "!" Then  ' Add this check
        lstGetList.AddItem inLines(L)
    End If
Next
Clipboard.Clear
End Sub

Private Sub Picture1_DblClick()
sendToIRC "WHO " & MyChannel
End Sub

Private Sub tmrDirty_Timer()
If Dirty = True Then
    cmdSave.BackColor = vbYellow
Else
    cmdSave.BackColor = &H8000000F
End If
End Sub


Private Sub tmrListMaint_Timer()
Static lLoops As Long
Dim SVRName As String
Dim Index As Integer
Dim LstCnt As Integer
Dim SvrNum As Integer

lLoops = lLoops + 1
If lLoops = 15 Then
   lLoops = 0
   'If Me.FLGetter.AmActive = True Then Exit Sub
   LstCnt = Me.lstGetNewList.ListCount
   If LstCnt = 0 Then Exit Sub
   Index = rand(0, LstCnt - 1)
   SVRName = Me.lstGetNewList.List(Index)
   SvrNum = CInt(Me.lstGetNewList.ItemData(Index))
   Me.FLGetter.NewPoll SvrNum
   Me.lstGetNewList.RemoveItem Index
   ' BailOut Time in case server never answers
   
   Debug.Print "Polling " & SVRName
' BOOKMARK
End If
End Sub

Private Sub tmrOffLine_Timer()
Static Countup As Integer
Dim Z As Integer
Dim num As Integer
Dim SVR As String
Dim TMPStr As String

Countup = Countup + 1

If Countup = 120 Then
      Countup = 0
      num = lstOffLine.ListCount
      If num = 0 Then Exit Sub
      For Z = 0 To num - 1
            TMPStr = lstOffLine.List(Z)
            SVR = nWord(TMPStr, " ")
            SVR = Right(SVR, Len(SVR) - 1)
            If ISOnline(lstUsers, SVR) = True Then
                lstGetList.AddItem lstOffLine.List(Z)
                lstOffLine.RemoveItem Z
            End If
      Next
End If
End Sub

Private Sub txtNeworkPort_Change()
Dirty = True

End Sub


Private Sub cmdSave_Click()
    SaveConfigs
    cmdFrame_Click 0
End Sub



Private Sub cmdBrowse_Click()
Dim TMPStr As String
TMPStr = BrowseFolder(Me, "Download Folder")
If TMPStr = "" Then Exit Sub
txtDownloadFolder.text = TMPStr
Dirty = True
End Sub


Private Sub txtBotNick_Change()
Dirty = True

End Sub

Private Sub txtNickServPass_Change()
Dirty = True

End Sub


Private Sub txtChanName_Change()
Dirty = True
End Sub

Private Sub txtNetworkAddr_Change()
Dirty = True

End Sub

Private Sub txtNetWorkName_Change()
Dirty = True
End Sub



Private Sub File2_DblClick()
Dim Cnt As Long
Dim L As Long
Dim Tmp As String
Dim Handle As Integer
Dim OpenFile As String

Me.LVBrowse.ListItems.Clear
BuildBookListview Me.LVBrowse
File2.Path = MyListFolder
File2.Pattern = "*.txt"


For L = 0 To File2.ListCount - 1
    If File2.Selected(L) = True Then
        OpenFile = File2.Path & "\" & File2.List(L)
        Exit For
    End If
Next
lblBrowseInfo.Caption = ""
lblBrowseInfo.Caption = "WORKING ..."
Tmp = QuickRead(OpenFile)
FileLines = Split(Tmp, vbCrLf)
Tmp = ""

For L = 1 To 22
    AddTitleLV Me.LVBrowse, FileLines(L)
Next

CurrLine = 22
LVBrowse.SetFocus
lblBrowseInfo.Caption = NoPath(OpenFile) & vbCrLf & "DONE " & UBound(FileLines) & " Files"


End Sub

Private Sub LVBrowse_DblClick()
Dim Cnt As Long
Dim Trigger As String
Cnt = LVBrowse.ListItems.Count
For L = 1 To Cnt
    If LVBrowse.ListItems(L).Selected = True Then
        Trigger = LVBrowse.ListItems(L).SubItems(3)
        LVBrowse.ListItems(L).Selected = False
        Exit For
    End If
Next
frmGetMain.lstGetList.AddItem Trigger

End Sub

Private Sub LVBrowse_KeyDown(KeyCode As Integer, Shift As Integer)
Dim L As Long
Dim Points As Long
If CurrLine < 0 Then CurrLine = 0
Select Case KeyCode
    Case vbKeyPageDown
        KeyCode = vbKeyShift
        LVBrowse.ListItems.Clear
        Points = CurrLine + 22
        If Points > UBound(FileLines) Then Points = UBound(FileLines)
        
        For L = CurrLine To Points
            AddTitleLV Me.LVBrowse, FileLines(L)
        Next
        CurrLine = Points
    Case vbKeyPageUp
        KeyCode = vbKeyShift
        LVBrowse.ListItems.Clear
        CurrLine = CurrLine - 44
        If CurrLine < 0 Then CurrLine = 0
        Points = CurrLine + 22
        
        For L = CurrLine To Points
            AddTitleLV Me.LVBrowse, FileLines(L)
        Next
        CurrLine = Points
End Select
End Sub



Private Sub cmdAddSelected_Click()
Dim Cnt As Integer
Dim L As Long
Dim tText As String

Cnt = lvTitles.ListItems.Count
For L = 1 To Cnt
    If lvTitles.ListItems(L).Selected = True Then
        tText = lvTitles.ListItems(L).SubItems(1)
        If Left(Trim(tText), 1) = "!" Then  ' Add this check
            StoreRequest tText
        End If
        lvTitles.ListItems(L).Selected = False
    End If
Next
End Sub



Private Sub cmdExit_Click()
FrameSearchResults.Visible = False
End Sub



Private Sub cmdIRCSearch_Click()
If SearchMODE = False Then
    SearchMODE = True
    cmdIRCSearch.BackColor = vbGreen
Else
    SearchMODE = False
    cmdIRCSearch.BackColor = &HC0C0FF
End If


End Sub

Private Sub cmdTreeViewToggle_Click()
If Tree_View = True Then
    Tree_View = False
    TV1.Visible = False
    lvTitles.Visible = True
    cmdTreeViewToggle.Caption = "List View Active"
Else
    Tree_View = True
    TV1.Move lvTitles.Left, lvTitles.Top, lvTitles.Width, lvTitles.Height
    lvTitles.Visible = False
    TV1.Visible = True
    cmdTreeViewToggle.Caption = "Tree View Active"
End If
End Sub



Private Sub frameServers_DblClick()
If SVRPause = False Then
    SVRPause = True
Else
    SVRPause = False
End If
If SVRPause = False Then
    frameServers.ForeColor = vbBlack
Else
    frameServers.ForeColor = vbYellow
End If

End Sub


Public Sub ReceiveFLFile(BufferLine As String)

Dim NickName As String
Dim fName As String
Dim fSize As Long
Dim IPAddr As String
Dim StoredIP As String
Dim Tmp As String
Dim Port As String
Dim ProcessLine As String
Dim Record As Long
Dim clsServer As clsListServer
Set clsServer = New clsListServer

Tmp = BufferLine
Garbage = nWord(Tmp, "@")
StoredIP = nWord(Tmp, " ")

IsActive = True
ArchPath = MyInComingFolder
ProcessDCCSend BufferLine, NickName, fName, IPAddr, Port, fSize
Debug.Print "Received List Download Code from " & NickName
'BOOKMARK
Record = GetServerByHash(StoredIP)
GetServerByClass clsServer, Record
Debug.Print clsServer.LSNickName

FLGetter.Activate clsServer, BufferLine

Set clsServer = Nothing
End Sub

Private Sub cmdConnect_Click()
Dim Ask As Integer

' Connect or Disconnect or Wait
Select Case CInt(Val(cmdConnect.Tag))
    Case 0
        ' you are disconnected, connect now
        ' Do_Connect
        If MyNetworkAddr = "" Then Exit Sub
        cmdConnect.Caption = "CONNECTING"
        cmdConnect.BackColor = vbYellow
        sckServer.Close
        sckServer.Connect MyNetworkAddr, MyNetworkPort
        cmdConnect.Tag = "1"
        SysLog "Connecting ..."
    Case 1
        ' you are connected, are your sure you want to disconnect?
        Ask = VBA.MsgBox("Disconnect from " & MyChannel, vbYesNoCancel)
        If Ask <> vbYes Then Exit Sub
        sendToIRC "Quit :" & "Quit"
        Pause 1
        sckServer.Close
        cmdConnect.Tag = "0"
        cmdConnect.Caption = "NOT CONNECTED"
        cmdConnect.BackColor = Grayed
        RTBLog.text = ""
        ' Send Disconnect string
    Case 2
        
        ' you have been disconnected, waiting
    Case 3
        If MyNetworkAddr = "" Then Exit Sub
        cmdConnect.Caption = "CONNECTING"
        cmdConnect.BackColor = vbYellow
        sckServer.Close
        sckServer.Connect MyNetworkAddr, MyNetworkPort
        cmdConnect.Tag = "1"
        SysLog "Connecting ..."
    Case Else
End Select

End Sub

Private Sub LoadChanConfigs()
Dim PathHeader As String
Dim L As Long
Dim Z As Long
Dim Message As String
Dim ThisServer As clsListServer
Dim ServerHandle As Integer
Dim ServerFile As String
Dim ServerCnt As Long
Dim zz As Long

   On Error GoTo LoadChanConfigs_Error

LoadConFigs

If NulTrim(IRCConfig.ChannelInfo.ChannelName) = "" Then GoTo Bottom
cmdConnect.Enabled = True
MyNick = NulTrim(IRCConfig.ChannelInfo.BotNickName)
MyNetwork = NulTrim(IRCConfig.ChannelInfo.NetworkName)
MyNetworkAddr = NulTrim(IRCConfig.ChannelInfo.NetworkAddress)
MyNetworkPort = NulTrim(IRCConfig.ChannelInfo.NetworkPort)
MyNickServPass = ShiftIt(NulTrim(IRCConfig.ChannelInfo.NickServPass), 61, 1)
MyChannel = NulTrim(IRCConfig.ChannelInfo.ChannelName)
PathHeader = App.Path & "\" & MyChannel & "@" & MyNetwork
MyPathHeader = PathHeader
MyListFolder = PathHeader & "\Lists\"
MyAdditionalFolder = PathHeader & "\Suplimental\"
MyLogFolder = PathHeader & "\Logs\"
MyServersFolder = PathHeader & "\Servers\"
MyProcessFolder = PathHeader & "\Process\"
MyInComingFolder = PathHeader & "\Incoming\"
MySearchResultFolder = PathHeader & "\Saved_SearchBot_Results\"
MyDownloadFolder = NulTrim(IRCConfig.ChannelInfo.DLFolder) & "\"
GetListSaveFile = App.Path & "\FWGetlist.txt"
TransferLogFile = App.Path & "\Transfers.log"
' Load Servers
If DirExist(MyServersFolder) = False Then
    MkDir PathHeader
    MkDir MyServersFolder
    MkDir MyListFolder
    MkDir MyInComingFolder
    MkDir MyAdditionalFolder
    MkDir PathHeader & "\Saved_SearchBot_Results"
End If
If DirExist(MyDownloadFolder) = False Then
      MkDir MyDownloadFolder
End If
Me.Show
DoEvents

For L = 0 To 4
    fGetter(L).Initialize
Next
TV1.Nodes.Add , , "root", "Search"
FLGetter.Initialize
File2.Path = MyListFolder
File2.Pattern = "*.txt"
ReadIgnores
ServerPoll
Me.Refresh
DoEvents
Exit Sub

Bottom:
AmConfigured = False
cmdConnect.Enabled = False
Me.Show
InitializeConfig
cmdFrame_Click 4
Me.Refresh

Do While Not AmConfigured
'    SSTab1.Tab = 4
    DoEvents
Loop

Form_Load

   On Error GoTo 0
   Exit Sub

LoadChanConfigs_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadChanConfigs of Form frmGetMain", "Debuglog.txt"

End Sub

Private Sub cmdLogFile_Click()

Dim inFile As String
Dim RtnVal As Long
inFile = App.Path & "\Transfers.log"
RtnVal = ShellExecute(Me.hWnd, "Open", inFile, "", App.Path, 3)

End Sub

Private Sub cmdPause_Click()
If AmPaused = False Then
    AmPaused = True
    cmdPause.Caption = "TRANSFERS PAUSED"
    cmdPause.BackColor = vbYellow
Else
    AmPaused = False
    cmdPause.Caption = "PAUSE TRANSFERS"
    cmdPause.BackColor = &H8000000F
End If
End Sub

Private Sub cmdSearch_Click()
Dim SearchTerms As String
Dim SearchInfo As SearchHeader
Dim Sample As String
Dim L As Long
Dim Cnt As Integer

SearchAbort = False
cmdSearchAbort.Caption = "ABORT"

lstRevert.Clear
lvTitles.ListItems.Clear

If txtSearch.text = "" Then Exit Sub
'lblListCount(0).Caption = "SEARCHING...."
lstGetList.Visible = False
lvTitles.Visible = True
TV1.Nodes.Clear
TV1.Nodes.Add , , "root", "Search"
'frmSpinner.Show
On Error Resume Next
SearchTerms = txtSearch.text
With SearchInfo
    .Completed = False
    .MaxResults = 2500
    .SearchTerms = SearchTerms
    .ListFolder = MyListFolder
    .OutputName = App.Path & "\SResults.txt"
    .ResultFolder = App.Path
End With
Me.MousePointer = vbHourglass
DoSearch SearchInfo

Me.MousePointer = vbNormal

SearchAbort = False
cmdSearchAbort.Caption = "Clear"
End Sub

Private Sub cmdShowGetList_Click()
If lstGetList.Visible = False Then
    ' Fill List
    FillRequestList
    
    lstGetList.Move FrameMain(0).Left, FrameMain(0).Top, FrameMain(0).Width, FrameMain(0).Height
    lstGetList.Visible = True
    lvTitles.Visible = False
    cmdShowGetList.Caption = "Hide Get List"
    'lstGetList.ZOrder = 0
Else
    lstGetList.Visible = False
    lvTitles.Visible = True
    cmdShowGetList.Caption = "Show Get List"
End If

End Sub


Private Sub Form_Load()

'FrameMain(0).Move 4545, 7500, 11460, 7650
cmdFrame(0).BackColor = &H80000016
Picture1.Move File2.Left, File2.Top
For L = 1 To 5
'    FrameMain(L).Move FrameMain(0).Left, FrameMain(0).Top, FrameMain(0).Width, FrameMain(0).Height
    FrameMain(L).Visible = False
Next

Initialize

InitResizeArray Me
SearchMODE = True
cmdIRCSearch.BackColor = vbGreen
If chkFilterRequests.Value = 1 Then
   chkShowRequest.Value = 0
Else
   chkShowRequest.Value = 1
End If
If chkFilterSearches.Value = 1 Then
   chkShowSearch.Value = 0
Else
   chkShowSearch.Value = 1
End If
RTBChat.text = "You can go online to search or you can click the search button above and search through server lists"
RTBChat.text = RTBChat.text & " you have collected."
End Sub

Private Sub Initialize()
Dim Handle As Integer
Dim inFile As String
Dim L As Integer
Dim InString As String
Dim Ctr As Control

'InitializeConfig
LoadChanConfigs
inFile = App.Path & "\FWGetlist.txt"

MyNick = NulTrim(IRCConfig.ChannelInfo.BotNickName)

tTop = 750
tLeft = 4545
tWidth = 11460
tHeight = 7650

For Each Ctr In Me.Controls
   If TypeOf Ctr Is Frame Then
   Ctr.Move tLeft, tTop, tWidth, tHeight
   End If
Next

If MyNick = "" Then
    cmdFrame_Click 3
    LoadChanConfigs
Else
End If

LVServers.ColumnHeaders.Clear
lvTitles.ColumnHeaders.Clear
lstGetList.Clear

BuildMyListview LVServers
BuildBookListview lvTitles

LoadCBO
cboOldSearches.ListIndex = 0
FillRequestList

Me.FrameSearch.Move 10, 10, tWidth, tHeight
Me.frameServers.Move 10, 10, tWidth, tHeight
Me.FrameConfiguration.Move 10, 10, tWidth, tHeight
Me.FramePrivateChat.Move 10, 10, tWidth, tHeight
Me.FrameBrowseList.Move 10, 10, tWidth, tHeight
Me.FrameIRC.Move 10, 10, tWidth, tHeight


cmdFrame(0).BackColor = &H80000016

For L = 1 To 5
'    FrameMain(L).Move FrameMain(4).Left, FrameMain(4).Top, FrameMain(4).Width, FrameMain(4).Height
    FrameMain(L).Visible = False
Next
cmdFrame_Click 0

chkShowRequest.Value = chkFilterRequests.Value
chkShowSearch.Value = chkFilterSearches.Value
   


MyVersion = "FWFetch ver. " & App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = MyVersion & " by FryMiester © 2023" & "   " & MyChannel & "@" & MyNetwork & " " & MyNick
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Handle As Integer
Dim Outfile As String
Dim LstCnt As Integer
Dim L As Integer
Dim ThisServer As clsListServer
Dim IPAddr As String
Dim Ask As Long

   On Error GoTo Form_QueryUnload_Error

If sckServer.State = sckConnected Then
    Ask = VBA.MsgBox("You are connected to " & MyNetwork & ", Exit Anyway", vbYesNoCancel)
    If Ask <> vbYes Then
        Cancel = vbCancel
        Exit Sub
    End If
End If

Me.sckServer.Close

Close
UnloadAllForms
End


   On Error GoTo 0
   Exit Sub

Form_QueryUnload_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_QueryUnload of Form frmGetMain", "Debuglog.txt"
    End
End Sub

Private Sub Form_Resize()
    On Error GoTo Form_Resize_Error
    
    If Me.WindowState = vbMinimized Then
         If chkToTray.Value = vbChecked Then
            Tray.TrayToolTip = "FWFetch "
            Tray.Show True
            Me.Hide
         Else
            Tray.Show False
            Me.WindowState = vbMinimized
         End If
    ElseIf Me.WindowState = vbNormal Then
        ResizeControls Me
        Tray.Show False
        Me.Show
    ElseIf Me.WindowState = vbMaximized Then
        ResizeControls Me
        Tray.Show False
        Me.Show
    Else
        ResizeControls Me
        Me.Show
    End If
    
    
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form Form1"
End Sub




Private Sub lblInfo_Click(Index As Integer)
 '   IRCNotice "FryMiester", "FWFETCH Version " & App.Major & App.Minor & App.Revision
End Sub

Private Sub lblSpy_DblClick()
' Raw WhoIs list
sendToIRC "WHOIS " & MyNick
End Sub

Private Sub lstGetList_DblClick()
Dim Cnt As Integer
Dim L As Integer
Dim tTempstr As String

Cnt = lstGetList.ListCount
For L = 0 To Cnt - 1
If lstGetList.Selected(L) = True Then
    tTempstr = lstGetList.List(L)
    RemoveGETListItem tTempstr, False
    Exit For
End If
Next
End Sub

Private Sub lstGetList_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Cnt As Integer
Dim L As Integer
Cnt = lstGetList.ListCount
If Button = vbRightButton Then
    Me.PopupMenu mnuGetList, , lstGetList.Left + 1500, lstGetList.Top + 1500
End If
 
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Ask As Long
Dim Cnt As Integer
Dim User As String
Dim L As Long

Cnt = lstUsers.ListCount
For L = 0 To Cnt - 1
    If lstUsers.Selected(L) = True Then
        User = lstUsers.List(L)
        Exit For
    End If
Next
User = Replace(User, "~", "")
User = Replace(User, "&", "")
User = Replace(User, "@", "")
User = Replace(User, "%", "")
User = Replace(User, "+", "")

If Button = vbRightButton Then
    Ask = VBA.MsgBox("Add " & User & " to Ignore list", vbYesNoCancel)
    If Ask <> vbYes Then Exit Sub
    lstIgnores.AddItem User
    WriteIgnores
End If

End Sub

Private Sub LVServers_DblClick()
' Request new List file from selected server
Dim Rcd As Long
Dim Cnt As Integer
Dim L As Long
Dim txtServer As String
Dim RecordNum As Integer
Dim Success As Long
Dim SVRPkg As ListServerInfo
Dim MyServer As clsListServer
Set MyServer = New clsListServer

Cnt = LVServers.ListItems.Count
For L = 1 To Cnt
    If LVServers.ListItems(L).Selected = True Then
        txtServer = LVServers.ListItems(L).text
        RecordNum = CLng(LVServers.ListItems(L).SubItems(3))
        LVServers.ListItems(L).Selected = False
        Exit For
    End If
Next

FLGetter.NewPoll RecordNum, True

End Sub



Private Sub lvTitles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvTitles
            .Sorted = False
            .SortKey = ColumnHeader.Index - 1
            If ColumnHeader.text = "Size" Then .SortKey = ColumnHeader.Index + 1
            
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
            
            .Sorted = True
        End With
End Sub

Private Sub lvtitles_DblClick()
Dim Cnt As Integer
Dim L As Long
Dim tText As String

Cnt = lvTitles.ListItems.Count
For L = 1 To Cnt
    If lvTitles.ListItems(L).Selected = True Then
    
        'lstGetList.AddItem lvTitles.ListItems(L).SubItems(1)
        tText = lvTitles.ListItems(L).SubItems(1)
        If Left(Trim(tText), 1) <> "!" Then
            lvTitles.ListItems(L).Selected = False
            Exit Sub
        End If
        
        StoreRequest tText
        lvTitles.ListItems(L).Selected = False
    End If
Next
End Sub

Private Sub LVServers_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Rcd As Integer
Dim Nick As String
Dim Hash As String

Dim MD5 As CMD5
Set MD5 = New CMD5

If Button = vbRightButton Then
    If Not (LVServers.SelectedItem Is Nothing) Then
        If Not (LVServers.HitTest(X, y) Is Nothing) Then
            Index = LVServers.SelectedItem.Index
            Nick = LVServers.ListItems.Item(Index).text
            Hash = MD5.HaxMD5(Nick)
            Rcd = GetServerByHash(Hash)
            CurrSVR = Rcd
            Me.PopupMenu mnuServersPopup
        Else
            'Debug.Print "not over item"
        End If
    End If
End If
Set MD5 = Nothing

Set ThisHereServer = Nothing
End Sub

Private Sub lvTitles_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lstItem As ListItem
Set lstItem = lvTitles.HitTest(X, y)
If Button = 0 Then
    If lstItem Is Nothing Then
         lvTitles.ToolTipText = ""
    Else
         lvTitles.ToolTipText = lstItem.SubItems(1) 'lstItem.Text ' & " - " & lstItem.SubItems(1)
    End If
End If
End Sub

Private Sub mnu_about2_Click()
frmAbout.Show
End Sub

Private Sub mnu_DeleteServer_Click()
Dim DisAbled As Boolean
Dim Ask As Long
Dim Rcd As Long
'BOOKMARK

    Ask = VBA.MsgBox(NulTrim(ThisServer.LSNickName) & " Will be marked as deleted, The system MAY re-add them later. Continue", vbYesNoCancel)
    If Ask <> vbYes Then Exit Sub
    With ThisServer
        ThisServer.Deleted = True
        .FilesCount = 0
        .FilesSize = 0
        .HasSlots = False
        .IPAddress = ""
        .isDisabled = False
        .ISOnline = False
        .ListDate = 0
        .ListNeeded = False
        .LSNickName = ""
        .NextUpDate = 0
    End With
    Rcd = GetServerByHash(NulTrim(ThisServer.IPAddress))
    ThisServer.IPAddress = ""
    ThisServer.LSNickName = ""
    PutServerByClass ThisServer, ThisServer.RecordNum
    
    ServerPoll
End Sub

Private Sub mnu_Disable_Click()
Dim DisAbled As Boolean
Dim Ask As Long
Dim Rcd As Long


DisAbled = LServers(CurrSVR).isDisabled
Select Case DisAbled
Case True
    Ask = VBA.MsgBox(NulTrim(LServers(CurrSVR).ServerName) & " is currently DIS-Abled, Re-enable this server", vbYesNoCancel)
    If Ask <> vbYes Then Exit Sub
    
    LServers(CurrSVR).isDisabled = False
    SaveServerINI CurrSVR
Case False
    Ask = VBA.MsgBox(NulTrim(LServers(CurrSVR).ServerName) & " is currently NOT Disabled, DISable this server", vbYesNoCancel)
    If Ask <> vbYes Then Exit Sub
    LServers(CurrSVR).isDisabled = True
    SaveServerINI CurrSVR
End Select
ServerPoll


End Sub

Private Sub mnuExit_Click()
Dim Outfile As String
Dim Handle As Integer
Dim IPAddr As String
ININame = MyServersFolder & "Servers.ini"
Outfile = MyServersFolder & "Servers.dat"
UnloadAllForms
End
End Sub


Private Sub RTBChat_DblClick()

Dim StartPoint As Long
Dim NewStartPoint As Long
Dim EndPoint As Long
Dim LastCR As Long
Dim LineNum As Long
Dim tUser As String
Dim L As Long
Dim Compare As String
Dim TempStr As String

On Error Resume Next

StartPoint = RTBChat.SelStart
EndPoint = RTBChat.Find(vbNewLine, StartPoint + 1)
NewStartPoint = RTBChat.Find(">", EndPoint)

RTBChat.SelStart = StartPoint + 1
RTBChat.SelLength = (EndPoint - StartPoint) - 1
TempStr = RTBChat.SelText
Compare = NulTrim(UCase(TempStr))
For L = 0 To lstUsers.ListCount - 1
    tUser = lstUsers.List(L)
    tUser = Replace(tUser, "~", "")
    tUser = Replace(tUser, "&", "")
    tUser = Replace(tUser, "@", "")
    tUser = Replace(tUser, "%", "")
    tUser = Replace(tUser, "+", "")
    
    If UCase(Compare) = UCase(tUser) Then
        lstUsers.Selected(L) = True
        Exit For
    End If
Next

End Sub

Private Sub sckServer_Connect()
If sckServer.State = sckConnected Then
    doConnectString
End If
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
Dim sockBuff As String
Dim ThisNick As String
Dim nCount As Integer
Dim Sample As String
Dim TMPStr As String
Dim Ask As Long
Dim y() As String


sckServer.GetData sockBuff, vbString
      TMPStr = sockBuff
      TMPStr = Right(TMPStr, Len(TMPStr) - 1)
      ThisNick = nWord(TMPStr, "!")
      Garbage = nWord(TMPStr, ":")
      Sample = TMPStr
      
    DebugLog sockBuff, App.Path & "\Rawlog.txt"
    Debug.Print sockBuff
    
    
    
Select Case True
      Case InStr(sockBuff, ":DCC SEND") <> 0
            ReceiveFile sockBuff
            sockBuff = ""
            Exit Sub
    Case InStr(sockBuff, ":TRIGGER") <> 0
            GetSearchBots sockBuff
            sockBuff = ""
            Exit Sub
      Case InStr(sockBuff, ":ChanServ!") <> 0
            If InStr(sockBuff, MyChannel) <> 0 Then
                  If InStr(sockBuff, MyNick) <> 0 Then
                        ' Channel Joined
                        AmConnected = True
                        cmdConnect.BackColor = &HC0FFC0
                        cmdConnect.Caption = "CONNECTED"
                        RTBChat.text = ""
                        PaintScreen "-- CONNECTED TO " & MyChannel & " --", vbGreen
                        Me.lstUsers.Clear
                        sendToIRC "NAMES " & MyChannel
                        Pause 5
                        IRCSay MyChannel, "@Searchbot-Trigger"
                        tmrListMaint.Enabled = True
                  End If
            End If
      Case InStr(sockBuff, "NOTICE") <> 0 And InStr(sockBuff, MyNick) <> 0
            '':FryMiester!FryMiester@ihw-c8f96o.lmaj.jro2.1700.2600.IP NOTICE Guest-F4FDE73E :Test Message
            sockBuff = Right(sockBuff, Len(sockBuff) - 1)
            ThisNick = nWord(sockBuff, "!")
            Garbage = nWord(sockBuff, ":")
            Sample = sockBuff
            Sample = StripColor(Sample)
            lblMessages.Caption = "< " & ThisNick & " >  " & Sample
            PaintScreen "< " & ThisNick & " NOTICE: >  " & Sample, vbRed
            Sample = ""
            sockBuff = ""
            Exit Sub
            
      Case InStr(sockBuff, "KICK") <> 0 And InStr(sockBuff, MyNick) <> 0
            sockBuff = Right(sockBuff, Len(sockBuff) - 1)
            ThisNick = nWord(sockBuff, "!")
            Garbage = nWord(sockBuff, ":")
            Sample = sockBuff
            PaintScreen "< " & ThisNick & " KICK: >  " & Sample, vbRed
            Sample = ""
            sockBuff = ""
            sendToIRC "JOIN " & NulTrim(MyChannel)
            Exit Sub
      Case InStr(UCase(sockBuff), " JOIN :") <> 0
      ':bleb!oinkasdf@ihw-b1c6c9.ipv6.telus.net JOIN :#ebooks
            sockBuff = Right(sockBuff, Len(sockBuff) - 1)
            ThisNick = nWord(sockBuff, "!")
            DoUserList ThisNick
            sockBuff = ""
            Exit Sub
      Case InStr(UCase(sockBuff), " PART :") <> 0
            sockBuff = Right(sockBuff, Len(sockBuff) - 1)
            ThisNick = nWord(sockBuff, "!")
            DoUserQuits ThisNick
            sockBuff = ""
            Exit Sub
      Case InStr(UCase(sockBuff), " QUIT :") <> 0
            sockBuff = Right(sockBuff, Len(sockBuff) - 1)
            ThisNick = nWord(sockBuff, "!")
            DoUserQuits ThisNick
            sockBuff = ""
            Exit Sub
       




Case Else
End Select


If Len(sockBuff) > 6 Then
    If Left(sockBuff, 6) = "PING :" Then
        'Debug.Print "PONG : " & sockBuff
        RemotePinger = Mid(sockBuff, InStr(sockBuff, "PING :") + 6)
        sendToIRC "PONG :" & RemotePinger 'Mid(sockBuff, InStr(sockBuff, "PING :") + 6)
    End If
End If

If InStr(sockBuff, " 353 " & MyNick) <> 0 Then
      ' NAMES
    RAWNamesList sockBuff
    ServerPoll
End If

If InStr(sockBuff, " 311 " & MyNick) <> 0 Then
    frmGetMain.FLGetter.GetWhois sockBuff
    
    sockBuff = ""
End If

If InStr(sockBuff, " 352 " & MyNick) <> 0 Then
    RAWWhoIs sockBuff
    sockBuff = ""
End If

'CURRENT

If InStr(sockBuff, ":VERSION") <> 0 Then
    y = Split(Mid(sockBuff, 2), "!")
    sendToIRC "NOTICE " & ThisNick & " :VERSION " & MyVersion & " by FryMiester - Go On, Get The Dawg! "
    sockBuff = ""
    SysLog ThisNick & " Requested VERSION information"
ElseIf InStr(sockBuff, ":CURRENT Version") <> 0 Then
    ThisNick = nWord(Mid(sockBuff, 2), "!")
    y = Split(sockBuff, " ")
    y(6) = nWord(y(6), "")
    TMPStr = "You are running FWFetch version " & App.Major & App.Minor & App.Revision & vbCrLf & " Version " & y(6) & " is currently available "
    TMPStr = TMPStr & vbCrLf & "If you want to download this version, Click YES, any other key to abort"
    Ask = VBA.MsgBox(TMPStr, vbYesNoCancel)
    If Ask = vbYes Then IRCNotice ThisNick, "!FWFetch"
    sockBuff = ""
    Exit Sub
ElseIf InStr(sockBuff, ":TIME") <> 0 Then
    y = Split(Mid(sockBuff, 2), "!")
    sendToIRC "NOTICE " & y(0) & " :TIME " & Format(Now, "ddd mmm dd") & " " & Time & " " & Format(Now, "yyyy") & ""
    sockBuff = ""
ElseIf InStr(sockBuff, ":PING") <> 0 Then
    y = Split(Mid(sockBuff, 2), "!")
    sendToIRC "NOTICE " & y(0) & " " & Mid(sockBuff, InStr(sockBuff, ":PING"))
    sockBuff = ""
ElseIf InStr(sockBuff, ":FINGER") <> 0 Then
    y = Split(Mid(sockBuff, 2), "!")
    sendToIRC "NOTICE " & y(0) & " :FINGER " & "Finger This" & ""
    sockBuff = ""
ElseIf InStr(sockBuff, ":DCC SEND") <> 0 Then
    'ColorText sockBuff, vbBlack
    ReceiveFile sockBuff
    sockBuff = ""
End If

RTBLog.text = ""

RTBLog.text = sockBuff

Dim TempInfo As String
    Dim InString As String
    On Local Error Resume Next
    InString = sockBuff
    If InStr(InString, "PRIVMSG") <> 0 Then
        If InStr(InString, vbCrLf) <> 0 Then
            Do While Len(InString)
                TempInfo = nWord(InString, vbCrLf)
                TempInfo = Replace(TempInfo, vbLf, "")
                TempInfo = Replace(TempInfo, vbCr, "")
                If NulTrim(TempInfo) > "" Then
                    ' get the nick of the ASSHOLE who is bulk pasting
                    ' get the count of the pastes
                    ' are they for me?
                    Sample = TempInfo
                    Sample = Right(Sample, Len(Sample) - 1)
                    ThisNick = nWord(Sample, "!")
                    nCount = nCount + 1
                    If nCount > 4 Then
                        SysLog ThisNick & " is a BULK Paster"
                    End If
                    TempInfo = StripColor(TempInfo)
                    TempInfo = ConvertChars(TempInfo)
                   frmGetMain.lstDialogBuff.AddItem TempInfo
                   Exit Sub
                End If
            Loop
        Else
            TempInfo = StripColor(TempInfo)
            TempInfo = ConvertChars(TempInfo)
            frmGetMain.lstDialogBuff.AddItem TempInfo
            Exit Sub
        End If
    Else
         TempInfo = StripColor(TempInfo)
         TempInfo = ConvertChars(TempInfo)
         frmGetMain.lstDialogBuff.AddItem TempInfo
         Exit Sub
    End If


End Sub

Public Sub ReceiveFile(BufferLine As String)

Dim NickName As String
Dim fName As String
Dim fSize As Long
Dim L As Integer
Dim IPAddr As String
Dim Port As String
Dim Index As Integer
Dim Matched As Boolean
Dim ProcessLine As String

ProcessLine = BufferLine
ProcessDCCSend ProcessLine, NickName, fName, IPAddr, Port, fSize


If InStr(UCase(NickName), "SEARCH") = 0 Then
    If InStr(UCase(NoPath(fName)), UCase(NickName)) <> 0 Then GoTo Bottom
    If FLGetter.AmActive = True Then
        If UCase(NickName) = NulTrim(UCase(FLGetter.ServerName)) Then GoTo Bottom
    End If
End If

Debug.Print "Received Download Code for Standard NoN-List File from " & NickName

For L = 0 To fGetter.UBound
    If NoPath(fGetter(L).FileName) = fName Then
        Matched = True
        fGetter(L).Activate BufferLine
        Exit For
    End If
Next

If Matched = False Then
    For L = 0 To fGetter.UBound
        If fGetter(L).AmActive = False Then
            Debug.Print "Received Download Code for List File from " & NickName
            fGetter(L).Activate BufferLine
            Exit For
        End If
    Next
End If


If Matched = False Then GoTo Bottom
Exit Sub

Bottom:
If Matched = False Then
    If NulTrim(FLGetter.ServerName) = NickName Then
        Debug.Print "Received Download Code for List File from " & NickName
        ReceiveFLFile BufferLine
    End If
End If

End Sub



Private Sub tmrDoBuff_Timer()
lblGoodCount.Caption = CStr(GoodCount)
lblBadcount.Caption = CStr(BadCount)
File1.Path = MyListFolder
lblListCount(1).Caption = "Currently " & CStr(File1.ListCount) & " File Lists For Searching"
If lstDialogBuff.ListCount > 0 Then
    DoBuff
End If
End Sub

Private Sub tmrUpdate_Timer()
If lvTitles.ListItems.Count > 0 Then
    lblListCount(0).Caption = CStr(lvTitles.ListItems.Count) & " Matches"
Else
    lblListCount(0).Caption = ""
End If
InChatWith = Me.FramePrivateChat.Caption

End Sub

Private Sub tmrUpGrade_Timer()
If AmConnected = False Then Exit Sub
UpgradeLoops = UpgradeLoops + 1
If UpgradeLoops = 300 Then
    UpgradeLoops = 0
    tmrUpGrade.Enabled = False
    IRCNotice "FryMiester", "FWFETCH Version " & App.Major & App.Minor & App.Revision
End If

End Sub

Private Sub tmtKeepAlive_Timer()
Dim CurrTime As String
Dim Cnt As Integer
Static myloops As Integer

CurrTime = Format(Now, "HH:MM")

If CurrTime > IRCConfig.ChannelInfo.SchedTime And CurrTime < IRCConfig.ChannelInfo.SchedStop Then
      INSchedule = True
Else
      INSchedule = False
End If

Cnt = lstGetList.ListCount
lblGetList.Caption = "Get List Count = " & CStr(Cnt)

myloops = myloops + 1
If AmConnected = True Then
        If myloops Mod 100 = 0 Then
            If RemotePinger > "" Then
                sendToIRC "PING :" & RemotePinger & vbCrLf
                Pause 1
                lstUsers.Clear
                sendToIRC "NAMES " & MyChannel
                myloops = 0
                ServerPoll
            End If
        End If
End If


If AmConnected = True Then
    If sckServer.State <> sckConnected Then
        cmdConnect.BackColor = vbRed
        cmdConnect.Caption = "Disconnected"
        cmdConnect.Tag = "3"
        AmConnected = False
    End If
End If


If NewConfig = True Then
    Initialize
    NewConfig = False
End If
End Sub

Private Sub Tray_LeftClick()
Me.WindowState = vbNormal
Me.Show
Me.Visible = True
Tray.Show False
End Sub

Private Sub Tray_MouseClick(Button As Long)
Me.WindowState = vbNormal
Me.Show
Me.Visible = True
Tray.Show False
End Sub

Private Sub TV1_DblClick()
Dim Gettext As String
Gettext = TV1.SelectedItem.text
lstGetList.AddItem Gettext

End Sub

Private Sub TVSearchBot_DblClick()
Dim Gettext As String
Gettext = TVSearchBot.SelectedItem.text
IRCSay MyChannel, Gettext
End Sub

Private Sub TV1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim strName As String
Dim lngItem As Long
Dim lngCount As Long
Dim objNode As Node
Dim Ask As Long

If Button = vbRightButton Then
Ask = VBA.MsgBox("Add ALL files from this section to the get list", vbYesNoCancel)
If Ask <> vbYes Then Exit Sub



''All child nodes of a given node
    
    Set objNode = TV1.Nodes(TV1.SelectedItem.Key)

'Get the count of children nodes for a loop
lngCount = objNode.Children

'If there are no childeren, there is no need to execute further
If Not lngCount = 0 Then

    'Get the first child node
    Set objNode = objNode.Child

    'Loop through the nodes
    For lngItem = 1 To lngCount
        strName = objNode.text 'Gets the node name
        If Left(strName, 1) = "!" Then
            lstGetList.AddItem strName
        End If
        Set objNode = objNode.Next 'Next Node
    Next
End If

'Destroy object
Set objNode = Nothing


'
End If
End Sub


Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo txtChat_KeyDown_Error

If KeyCode = 13 Then
    If SearchMODE = True Then
         If Left(txtChat.text, 1) <> "/" Then
                  ScrollBack.Add txtChat.text
                  If ScrollBack.Count > 20 Then
                      ScrollBack.Remove 20
                  End If
                  IRCSay MyChannel, cmdIRCSearch.Caption & " " & txtChat.text
                  txtChat.text = ""
                  ScrollNum = ScrollNum + 1
        Else
                  CmndProcess txtChat.text
        End If
    Else
        ScrollBack.Add txtChat.text
        IRCSay MyChannel, txtChat.text
        txtChat.text = ""
        ScrollNum = ScrollNum + 1
    End If
End If
If KeyCode = vbKeyUp Then
    ScrollNum = ScrollNum - 1
    If ScrollNum < 1 Then ScrollNum = ScrollBack.Count
    txtChat.text = ScrollBack.Item(ScrollNum)
End If
If KeyCode = vbKeyDown Then
    ScrollNum = ScrollNum + 1
    If ScrollNum > ScrollBack.Count Then ScrollNum = 1
    txtChat.text = ScrollBack.Item(ScrollNum)
End If
If KeyCode = vbKeyEscape Then
    txtChat.text = ""
End If

   On Error GoTo 0
   Exit Sub

txtChat_KeyDown_Error:

    DebugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure txtChat_KeyDown of Form frmGetMain", "Debuglog.txt"
End Sub

Private Sub txtPrivateChat_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TheNick As String
Dim ModeString As String
TheNick = Me.FramePrivateChat.Caption
If KeyCode = 13 Then
    IRCMsg TheNick, txtPrivateChat.text
    ModeString = Format(Now, "HH:MM") & " <" & MyNick & "> " & txtPrivateChat.text & vbCrLf
    ColorText ModeString, vbRed
    txtPrivateChat.text = ""
End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 13 Then
'        ScrollSearch.Add txtSearch.Text
'        If ScrollSearch.Count > 20 Then
'            ScrollSearch.Remove 20
'        End If
'        cmdSearch_Click
'        ScrollSearchNum = ScrollSearchNum + 1
'End If
If KeyCode = vbKeyUp Then
    ScrollSearchNum = ScrollSearchNum - 1
    If ScrollSearchNum < 1 Then ScrollSearchNum = ScrollSearch.Count
    txtSearch.text = ScrollSearch.Item(ScrollSearchNum)
End If
If KeyCode = vbKeyDown Then
    ScrollSearchNum = ScrollSearchNum + 1
    If ScrollSearchNum > ScrollSearch.Count Then ScrollSearchNum = 1
    txtSearch.text = ScrollSearch.Item(ScrollSearchNum)
End If
If KeyCode = vbKeyEscape Then
    txtSearch.text = ""
End If

If KeyCode = 13 Then
    cmdSearch_Click
End If
End Sub


Private Sub LoadConFigs()

ReadConfigs

If NulTrim(IRCConfig.ChannelInfo.ChannelName) > "" Then
        txtChanName.text = NulTrim(IRCConfig.ChannelInfo.ChannelName)
        txtNetWorkName.text = NulTrim(IRCConfig.ChannelInfo.NetworkName)
        txtNetworkAddr.text = NulTrim(IRCConfig.ChannelInfo.NetworkAddress)
        txtNeworkPort.text = NulTrim(IRCConfig.ChannelInfo.NetworkPort)
        txtBotNick.text = NulTrim(IRCConfig.ChannelInfo.BotNickName)
        txtNickServPass.text = ShiftIt(NulTrim(IRCConfig.ChannelInfo.NickServPass), 61, 1)
        txtDownloadFolder.text = NulTrim(IRCConfig.ChannelInfo.DLFolder)
        chkAutoConnect.Value = BoolToInt(IRCConfig.ChannelInfo.AutoConnect)
        chkFilterRequests.Value = BoolToInt(IRCConfig.ChannelInfo.ChkMark(0))
        chkFilterSearches.Value = BoolToInt(IRCConfig.ChannelInfo.ChkMark(1))
        If IRCConfig.ChannelInfo.SendToTray = "1" Then
           chkToTray.Value = vbChecked
        Else
           chkToTray.Value = vbUnchecked
        End If
        
        txtSchedTime.text = IRCConfig.ChannelInfo.SchedTime
        txtSchedStop.text = IRCConfig.ChannelInfo.SchedStop
        txtListExpire.text = IRCConfig.ChannelInfo.UpDateDays
        UpDateDays = CInt(IRCConfig.ChannelInfo.UpDateDays)

        If IRCConfig.ChannelInfo.IsScheduled = "1" Then
            chkSchedule.Value = 1
            UseSchedule = True
            cmdSchedule.Caption = UCase("Using Scheduled Transfers")
            cmdSchedule.BackColor = vbGreen
        Else
            chkSchedule.Value = 0
        End If
        
End If


End Sub

Private Sub SaveConfigs()
        
    IRCConfig.ChannelInfo.ChannelName = txtChanName.text
    IRCConfig.ChannelInfo.NetworkName = txtNetWorkName.text
    IRCConfig.ChannelInfo.NetworkAddress = txtNetworkAddr.text
    IRCConfig.ChannelInfo.NetworkPort = txtNeworkPort.text
    IRCConfig.ChannelInfo.BotNickName = txtBotNick.text
    IRCConfig.ChannelInfo.NickServPass = ShiftIt(txtNickServPass.text, 61, 0)
    IRCConfig.ChannelInfo.DLFolder = txtDownloadFolder.text
    IRCConfig.ChannelInfo.AutoConnect = CBool(Val(chkAutoConnect.Value))
    IRCConfig.ChannelInfo.ChkMark(0) = CBool(Val(chkFilterRequests.Value))
    IRCConfig.ChannelInfo.ChkMark(1) = CBool(Val(chkFilterSearches.Value))
    IRCConfig.ChannelInfo.UpDateDays = txtListExpire.text
        If chkSchedule.Value = vbChecked Then
              IRCConfig.ChannelInfo.IsScheduled = "1"
        Else
              IRCConfig.ChannelInfo.IsScheduled = "0"
        End If
        IRCConfig.ChannelInfo.SchedTime = txtSchedTime.text
        IRCConfig.ChannelInfo.SchedStop = txtSchedStop.text
       If chkToTray.Value = vbChecked Then
              IRCConfig.ChannelInfo.SendToTray = "1"
       Else
              IRCConfig.ChannelInfo.SendToTray = "0"
       End If


WriteConfigs
Dirty = False
AmConfigured = True
End Sub

Private Sub InitializeConfig()
Dim SerNo As Long
Dim TMPStr As String
Dim cControl As Control
Dim tTempstr As String
Dim MD5 As CMD5
Set MD5 = New CMD5

LoadConFigs

If txtBotNick.text = "" Then
    SerNo = GetSerialNumber
    TMPStr = "Guest-" & CStr(SerNo)
    tTempstr = MD5.HaxMD5(TMPStr)
    TMPStr = "Guest-" & tTempstr
    txtBotNick.text = TMPStr
'    SaveConfigs
End If

Dirty = True
tmrDirty.Enabled = True
End Sub

Public Sub LoadCBO()
' MySearchResultFolder
Dim cCount As Integer
Dim tText As String
Dim L As Long

Me.cboOldSearches.Clear
Me.cboOldSearches.AddItem "Old Searches", 0
Me.cboOldSearches.AddItem "DELETE Old Searches", 1

Me.File1.Path = MySearchResultFolder
Me.File1.Pattern = "*.txt"
cCount = Me.File1.ListCount
For L = 0 To cCount - 1
   tText = Me.File1.List(L)
   tText = noExt(tText)
   tText = Right(tText, Len(tText) - (InStr(tText, "for_") + 4))
   Me.cboOldSearches.AddItem tText, L + 2
Next

End Sub

Public Sub RemoveGETListItem(DeleteLine As String, IsDownloaded As Boolean, Optional ServerName As String)
Dim ReqININame As String
Dim Header As String
Dim REQHead As String
Dim sREQHead As String
Dim ExistREQ As String
Dim TempStr As String
Dim L As Long
Dim A As Long
Dim B As Long
Dim cCount As Integer
Dim sCount As Integer
Dim Matched As Boolean

ReqININame = MyServersFolder & "Requests.ini"
Header = "SERVERLIST"
cCount = CInt(Val(ReadINI(Header, "SERVERCOUNT", ReqININame)))
If cCount = 0 Then Exit Sub
frmGetMain.lstGetList.Clear

If IsDownloaded = False Then
   For L = 1 To cCount
      sCount = 0
      REQHead = "SVR_" & Format(L, "00#")
      ServerName = ReadINI(Header, REQHead, ReqININame)
      sCount = CInt(Val(ReadINI(ServerName, "REQUESTCOUNT", ReqININame)))
      If sCount > 0 Then
         For A = 1 To sCount
            sREQHead = "REQ_" & Format(A, "00#")
            TempStr = ReadINI(ServerName, sREQHead, ReqININame)
            Debug.Print TempStr
            If TempStr > "" Then
               If TempStr = DeleteLine Then
               WriteINI ServerName, sREQHead, "", ReqININame
               End If
            End If
         Next
      End If
   Next
Else
      sCount = CInt(Val(ReadINI(ServerName, "REQUESTCOUNT", ReqININame)))
      For L = 1 To sCount
            sREQHead = "REQ_" & Format(L, "00#")
            TempStr = ReadINI(ServerName, sREQHead, ReqININame)
            If InStr(TempStr, NoPath(DeleteLine)) > 0 Then
               WriteINI ServerName, sREQHead, "", ReqININame
               Exit Sub
            End If
      Next
End If

FillRequestList
End Sub
