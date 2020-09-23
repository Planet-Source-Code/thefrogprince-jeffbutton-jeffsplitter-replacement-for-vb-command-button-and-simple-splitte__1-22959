VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTestSplitter 
   Caption         =   "jeffSplitter Test Harness"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "frmTestSplitter.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv 
      Height          =   1365
      Left            =   3390
      TabIndex        =   4
      Top             =   390
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   2408
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Number"
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1545
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2725
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.TextBox txtBlahBlah 
      Height          =   1725
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmTestSplitter.frx":27A2
      Top             =   2370
      Width           =   6165
   End
   Begin ControlTestHarness.jeffSplitter splitterHorizontal 
      Height          =   135
      Left            =   180
      TabIndex        =   1
      Top             =   2070
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   238
      SplitterOrientation=   2
      SplitterMaxHeight=   4245
   End
   Begin ControlTestHarness.jeffSplitter splitterVertical 
      Height          =   1905
      Left            =   2910
      TabIndex        =   0
      Top             =   0
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3360
      SplitterOrientation=   1
      SplitterMaxWidth=   6375
   End
End
Attribute VB_Name = "frmTestSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.splitterHorizontal.BackColor = vbButtonFace
    Me.splitterVertical.BackColor = vbButtonFace
    LoadSamples
End Sub

Private Sub Form_Resize()
    
    With Me.splitterHorizontal
        .Left = 0
        .Width = Me.ScaleWidth
        .SplitterMaxHeight = Me.ScaleHeight
    End With
    With Me.splitterVertical
        .Top = 0
        .SplitterMaxWidth = Me.ScaleWidth
    End With
    With Me.txtBlahBlah
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
End Sub

Private Sub splitterHorizontal_SplitterMoved()
    With Me.tv
        .Top = 0
        .Height = splitterHorizontal.Top
    End With
    With Me.splitterVertical
        .Top = 0
        .Height = Me.tv.Height
    End With
    With Me.lv
        .Top = 0
        .Height = Me.tv.Height
    End With
    With Me.txtBlahBlah
        .Top = Me.splitterHorizontal.Top + Me.splitterHorizontal.Height
        .Height = Me.ScaleHeight - .Top
    End With
    
End Sub

Private Sub splitterVertical_SplitterMoved()
    With Me.tv
        .Left = 0
        .Width = Me.splitterVertical.Left
    End With
    With Me.lv
        .Left = Me.splitterVertical.Left + Me.splitterVertical.Width
        .Width = Me.ScaleWidth - .Left
    End With
    
End Sub

Public Function LoadSamples()
    
    Dim l As Long
    Dim m As Long
    Dim n As Long
    With Me.tv
        For l = 1 To 3
            .Nodes.Add , tvwNext, "topNode" & l, "Top Node " & l
            For m = 1 To 4
                .Nodes.Add "topNode" & l, tvwChild, "topNode" & l & "~secondNode" & m, "Second Node " & m
                For n = 1 To 2
                    .Nodes.Add "topNode" & l & "~secondNode" & m, tvwChild, "topNode" & l & "~secondNode" & m & "~thirdNode" & n, "Third Node " & n
                Next n
                Dim o As Long
                o = ((l - 1) * 4) + m
                lv.ListItems.Add(, "ITEM" & o, "Item" & o).SubItems(1) = CStr(o)
            Next m
        Next l
    End With
    
End Function
