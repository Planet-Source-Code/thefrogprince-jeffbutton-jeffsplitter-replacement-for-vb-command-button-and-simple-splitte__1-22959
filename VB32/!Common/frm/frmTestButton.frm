VERSION 5.00
Begin VB.Form frmTestButton 
   Caption         =   "Test Harness For jeffButton"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   Icon            =   "frmTestButton.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   3  'Windows Default
   Begin ControlTestHarness.jeffButton btnTest 
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   17
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      Caption         =   "Disabled &Test"
      Picture         =   "frmTestButton.frx":27A2
      CaptionAlignment=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Enabled         =   0   'False
   End
   Begin VB.PictureBox picToolbar 
      BackColor       =   &H8000000D&
      Height          =   585
      Left            =   90
      ScaleHeight     =   525
      ScaleWidth      =   4605
      TabIndex        =   6
      Top             =   300
      Width           =   4665
      Begin ControlTestHarness.jeffButton btnBack 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Back"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Back"
         Picture         =   "frmTestButton.frx":3094
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":35EE
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnForward 
         Height          =   375
         Left            =   510
         TabIndex        =   8
         ToolTipText     =   "Forward"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Forward"
         Picture         =   "frmTestButton.frx":3B48
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":40A2
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnRefresh 
         Height          =   375
         Left            =   960
         TabIndex        =   9
         ToolTipText     =   "Refresh "
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Refresh "
         Picture         =   "frmTestButton.frx":45FC
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":4B56
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnHome 
         Height          =   375
         Left            =   1410
         TabIndex        =   10
         ToolTipText     =   "Home"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Home"
         Picture         =   "frmTestButton.frx":50B0
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":560A
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnFavorites 
         Height          =   375
         Left            =   1860
         TabIndex        =   11
         ToolTipText     =   "Favorites"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Favorites"
         Picture         =   "frmTestButton.frx":5B64
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":60BE
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnGoUp 
         Height          =   375
         Left            =   2310
         TabIndex        =   12
         ToolTipText     =   "Go Up"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Go Up"
         Picture         =   "frmTestButton.frx":6618
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":6B72
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnHistory 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         ToolTipText     =   "History"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "History"
         Picture         =   "frmTestButton.frx":70CC
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":7626
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnSearch 
         Height          =   375
         Left            =   3210
         TabIndex        =   14
         ToolTipText     =   "Search"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Search"
         Picture         =   "frmTestButton.frx":7B80
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":80DA
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnGo 
         Height          =   375
         Left            =   3660
         TabIndex        =   15
         ToolTipText     =   "Go"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Go"
         Picture         =   "frmTestButton.frx":8634
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":8B8E
         ShowFocusRect   =   0   'False
      End
      Begin ControlTestHarness.jeffButton btnStop 
         Height          =   375
         Left            =   4110
         TabIndex        =   16
         ToolTipText     =   "Stop"
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Caption         =   ""
         Object.ToolTipText     =   "Stop"
         Picture         =   "frmTestButton.frx":8F28
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         HotPicture      =   "frmTestButton.frx":9482
         ShowFocusRect   =   0   'False
      End
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   1935
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "&Bert"
      Picture         =   "frmTestButton.frx":99DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12640511
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   1
      Left            =   660
      TabIndex        =   1
      ToolTipText     =   "TEST"
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "&Kermit"
      Object.ToolTipText     =   "TEST"
      Picture         =   "frmTestButton.frx":A2B6
      CaptionAlignment=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   2
      Left            =   2490
      TabIndex        =   2
      ToolTipText     =   "count"
      Top             =   2790
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "The &Count"
      Object.ToolTipText     =   "count"
      Picture         =   "frmTestButton.frx":AB90
      CaptionAlignment=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648384
      ButtonStyle     =   2
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   3
      Left            =   660
      TabIndex        =   3
      Top             =   2790
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "&Ernie"
      Picture         =   "frmTestButton.frx":B46A
      CaptionAlignment=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   4
      Left            =   2490
      TabIndex        =   4
      ToolTipText     =   "Big Bird"
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "B&ig Bird"
      Object.ToolTipText     =   "Big Bird"
      Picture         =   "frmTestButton.frx":BD44
      CaptionAlignment=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12583104
      BackColor       =   16761087
      ButtonStyle     =   2
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnMuppets 
      Height          =   795
      Index           =   5
      Left            =   2490
      TabIndex        =   5
      Top             =   1935
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      Caption         =   "Yummm... Me like cookies... cookies good!"
      Picture         =   "frmTestButton.frx":C61E
      CaptionAlignment=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761024
      ButtonStyle     =   2
      ShowFocusRect   =   0   'False
   End
   Begin ControlTestHarness.jeffButton btnTest 
      Height          =   615
      Index           =   1
      Left            =   2490
      TabIndex        =   18
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      Caption         =   "Enabled Test"
      Picture         =   "frmTestButton.frx":D1F0
      CaptionAlignment=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
   End
End
Attribute VB_Name = "frmTestButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

