VERSION 5.00
Begin VB.UserControl jeffButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   Begin VB.PictureBox picDisabled 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1680
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   1
      Top             =   270
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picNormal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1830
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picHot 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      DrawWidth       =   32
      Height          =   525
      Left            =   2700
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblSize"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   450
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgCurrent 
      Height          =   555
      Left            =   2640
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "This is a test of word wrap"
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   900
      TabIndex        =   0
      Top             =   420
      Width           =   390
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "jeffButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bHasFocus As Boolean

Public Enum enumJeffButtonStyles
    jbsStandard = 0
    jbsFlat = 1
    jbsSoft = 2
    jbsFlatSoft = jbsFlat + jbsSoft
    
End Enum

Public Enum enumCaptionAlignment
    ecaTop = 1
    ecaBottom = 2
    ecaLeft = 3
    ecaRight = 4
    ecaOverlayCenter = 5
End Enum

Private bMouseCaptured As Boolean
Private bMouseDown As Boolean
Private lPreviousCaptureHwnd As Long
'Default Property Values:
Const m_def_ButtonStyle = 0
Const m_def_CaptionAlignment = ecaBottom
Const m_def_ToolTipText = ""
'Property Variables:
Dim m_ButtonStyle As enumJeffButtonStyles
Dim m_CaptionAlignment As enumCaptionAlignment
Dim m_ShowFocusRect As Boolean

Dim m_ToolTipText As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_MemberFlags = "200"
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



Public Function SnapText()
    lblSize.AutoSize = False
    lblSize.AutoSize = True
End Function


Private Sub imgCurrent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = True
    RedrawControl

    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub





Private Sub lblCaption_Click()
'    RaiseEvent Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = True
    RedrawControl

    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = pixelsX(x)
    y = pixelsY(y)
    
    If x < 0 Or x > UserControl.ScaleWidth Or _
       y < 0 Or y > UserControl.ScaleHeight Then
        ReleaseMouse
    Else
        CaptureMouse
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
    
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseCaptured = False
    CaptureMouse
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    If bMouseDown Then
        bMouseDown = False
        RedrawControl
        RaiseEvent Click
    Else
        RedrawControl
    End If
        
        
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    bMouseCaptured = True
    bMouseDown = True
    RedrawControl
    RaiseEvent Click
    
    bMouseDown = False
    bMouseCaptured = False
    RedrawControl
    
End Sub


Private Sub UserControl_GotFocus()
    bHasFocus = True
    RedrawControl
End Sub


Public Function ClickKey(KeyCode As Integer, Optional ByVal Shift As Integer = 0) As Boolean
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        ClickKey = True
    End If
    If InStr(UCase(Me.Caption), "&" & UCase(Chr(KeyCode))) > 0 Then
        ClickKey = True
    End If
End Function


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If ClickKey(KeyCode) Then
        UserControl_MouseDown vbLeftButton, 0, 0, 0
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If ClickKey(KeyCode) Then
        UserControl_MouseUp vbLeftButton, 0, 0, 0
'        UserControl_Click
    End If
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    RedrawControl

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = True
    RedrawControl
    
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < 0 Or x > UserControl.ScaleWidth Or _
       y < 0 Or y > UserControl.ScaleHeight Then
        ReleaseMouse
    Else
        CaptureMouse
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Public Function CaptureMouse()
    If Not bMouseCaptured And (Me.ButtonStyle And jbsFlat) > 0 Then
        lPreviousCaptureHwnd = SetCapture(UserControl.hWnd)
        bMouseCaptured = True
        RedrawControl
        If GetCapture <> UserControl.hWnd Then
            Debug.Print "Issue"
        End If
    End If
    
End Function

Public Function ReleaseMouse()
    
    If bMouseCaptured And (Me.ButtonStyle And jbsFlat) > 0 Then
        bMouseDown = False
        ReleaseCapture
        If lPreviousCaptureHwnd <> 0 Then
            SetCapture lPreviousCaptureHwnd
                
        End If
        lPreviousCaptureHwnd = 0
        bMouseCaptured = False
        RedrawControl
        RaiseEvent MouseExit
    End If
End Function

Public Function LoadPicture()
'    Select Case True
'        Case Not UserControl.Enabled And UserControl.picDisabled.Picture.Height = 0
'            DrawState _
'                    picDisabled.hDC, _
'                    0, _
'                    0, _
'                    UserControl.picNormal.Picture.Handle, _
'                    0, _
'                    0, _
'                    0, _
'                    32, _
'                    32, _
'                    DST_ICON Or DSS_DISABLED
'            Set UserControl.imgCurrent.Picture = picDisabled.Image
'
'        Case Not UserControl.Enabled
'            Set UserControl.imgCurrent.Picture = picDisabled.Picture
'        Case Else
            Set UserControl.imgCurrent.Picture = Me.CurrentPic
            imgCurrent.Visible = False
'    End Select
End Function

Public Function RedrawControl()
    Static bRedrawingControl As Boolean
    If Not bRedrawingControl Then
        
        Dim lPicTop As Long
        Dim lPicLeft As Long
        Dim lPicHeight As Long
        Dim lPicWidth As Long
        lPicHeight = Me.CurrentPic.Picture.Height
        lPicWidth = Me.CurrentPic.Picture.Width
        
        bRedrawingControl = True
        
    '    FillPicture
'        ts.windowUpdate UserControl.hWnd, elwLOCK
        
        LoadPicture
        
'        UserControl.imgCurrent.Enabled = UserControl.Enabled
        lblCaption.Enabled = UserControl.Enabled
        
        
        ' Get Bitmap
        'UserControl.AutoRedraw = True
        
        Dim Pic As StdPicture
        
        Dim tr As RECT
    '    UserControl.Cls
        Dim eEdge As enumBorderEdges
        Dim lSpaceWidth As Long
            
        Select Case True
            Case m_CaptionAlignment = ecaBottom And lPicHeight > 0
                UserControl.lblSize.Width = UserControl.ScaleWidth - 12
                SnapText
                lSpaceWidth = UserControl.ScaleHeight - lblSize.Height ' - twipsY(4)
                lblSize.Move (UserControl.ScaleWidth - lblSize.Width) / 2, UserControl.ScaleHeight - 4 - lblSize.Height
                'lPicLeft = (UserControl.Width - lPicWidth) / 2
                'lPicTop = (lSpaceWidth - lPicHeight) / 2
                imgCurrent.Move (UserControl.ScaleWidth - imgCurrent.Width) / 2, (lSpaceWidth - imgCurrent.Height) / 2
                'FillPicture (UserControl.Width - imgCurrent.Width) / 2, (lSpaceWidth - imgCurrent.Height) / 2
            Case m_CaptionAlignment = ecaTop And imgCurrent.Picture.Height > 0
                UserControl.lblSize.Width = UserControl.ScaleWidth - 12
                SnapText
                lSpaceWidth = UserControl.ScaleHeight - lblSize.Height ' - twipsY(4)
                lblSize.Move (UserControl.ScaleWidth - lblSize.Width) / 2, 4
                imgCurrent.Move (UserControl.ScaleWidth - imgCurrent.Width) / 2, lblSize.Height + ((lSpaceWidth - imgCurrent.Height) / 2)
                'FillPicture (UserControl.Width - imgCurrent.Width) / 2, lblsize.Height + ((lSpaceWidth - imgCurrent.Height) / 2)
            Case m_CaptionAlignment = ecaLeft And imgCurrent.Picture.Width > 0
                On Error Resume Next
                UserControl.lblSize.Width = UserControl.ScaleWidth - imgCurrent.Width - 8
                If Err.Number <> 0 Then
                    UserControl.lblSize.Width = 1
                End If
                On Error GoTo 0
                SnapText
                
                'imgcurrent.Move lblsize.Left + lblsize.Width + ((lSpaceWidth - imgcurrent.Width) / 2), (UserControl.Height - imgcurrent.Height) / 2
                
                imgCurrent.Move (UserControl.ScaleWidth - imgCurrent.Width) - 4, (UserControl.ScaleHeight - imgCurrent.Height) / 2
                
                lSpaceWidth = imgCurrent.Left '- twipsX(4)
                
                lblSize.Move (lSpaceWidth - lblSize.Width) / 2, (UserControl.ScaleHeight - lblSize.Height) / 2
                
                'FillPicture (UserControl.Width - imgCurrent.Width) - twipsX(4), (UserControl.Height - imgCurrent.Height) / 2
                'lSpaceWidth = UserControl.Width - lblsize.Left - lblsize.Top
                'imgcurrent.Move lblsize.Left + lblsize.Width + ((lSpaceWidth - imgcurrent.Width) / 2), (UserControl.Height - imgcurrent.Height) / 2
                
            Case m_CaptionAlignment = ecaRight And imgCurrent.Picture.Width > 0
                On Error Resume Next
                UserControl.lblSize.Width = UserControl.ScaleWidth - imgCurrent.Width - 12
                If Err.Number <> 0 Then
                    UserControl.lblSize.Width = 1
                End If
                On Error GoTo 0
                SnapText
                imgCurrent.Move 4, (UserControl.ScaleHeight - imgCurrent.Height) / 2
                lblSize.Move imgCurrent.Left + imgCurrent.Width, (UserControl.ScaleHeight - lblSize.Height) / 2
                'FillPicture twipsX(4), (UserControl.Height - imgCurrent.Height) / 2
                
            Case m_CaptionAlignment = ecaOverlayCenter Or (imgCurrent.Picture.Height = 0 Or imgCurrent.Picture.Width = 0)
                imgCurrent.Move (UserControl.ScaleWidth - imgCurrent.Width) / 2, (UserControl.ScaleHeight - imgCurrent.Height) / 2
                UserControl.lblSize.Width = UserControl.ScaleWidth - 4
                lblSize.Move (UserControl.ScaleWidth - lblSize.Width) / 2, (UserControl.ScaleHeight - lblSize.Height) / 2
                'FillPicture (UserControl.Width - imgCurrent.Width) / 2, (UserControl.Height - imgCurrent.Height) / 2
                
        End Select
        lblCaption.Move lblSize.Left, lblSize.Top, lblSize.Width, lblSize.Height
        
        Select Case True
            Case bMouseDown And (m_ButtonStyle And jbsSoft) > 0
                eEdge = BDR_SUNKENINNER
            Case bMouseDown
                eEdge = EDGE_SUNKEN
            Case bMouseCaptured And (m_ButtonStyle And jbsFlat) > 0 And (m_ButtonStyle And jbsSoft) > 0
                eEdge = BDR_RAISEDINNER
            Case bMouseCaptured And (m_ButtonStyle And jbsFlat) > 0
                eEdge = EDGE_RAISED
            Case (m_ButtonStyle And jbsFlat) > 0
                eEdge = 0
            Case (m_ButtonStyle And jbsSoft) > 0
                eEdge = BDR_RAISEDINNER
            Case Else
                eEdge = EDGE_RAISED
        End Select

        tr.Right = UserControl.ScaleWidth
        tr.Bottom = UserControl.ScaleHeight
        
        ' Draw Pic
        Dim eText As Long
        If CurrentPic.Picture.Height > 0 Then
            Dim eDraw As Long
            


            Select Case True
                Case CurrentPic.Picture.Type = vbPicTypeIcon
                    eDraw = DST_ICON
                Case Else
                    eDraw = DST_BITMAP
            End Select
            
            Select Case True
                Case Not UserControl.Enabled And picDisabled.Picture.Height = 0
                    eDraw = eDraw Or DSS_DISABLED
'                    Debug.Print "Drawing Disabled Icon"
                Case Else
                    'eDraw = DST_ICON
'                    Debug.Print "Drawing Icon"
            End Select
            eDraw = eDraw Or DSS_RIGHT
            UserControl.Cls
            DrawState _
                    UserControl.hDC, _
                    0, _
                    0, _
                    CurrentPic.Picture.Handle, _
                    0, _
                    imgCurrent.Left, _
                    imgCurrent.Top, _
                    CurrentPic.Picture.Width, _
                    CurrentPic.Picture.Height, _
                    eDraw
        Else
            UserControl.Cls
        End If
        ' Draw Edge
        DrawEdge UserControl.hDC, tr, eEdge, BF_RECT
        
        eText = DST_TEXT
        If Not UserControl.Enabled Then
            eText = eText Or DSS_DISABLED
        End If

        
        tr.Bottom = tr.Bottom - 3
        tr.Left = tr.Left + 3
        tr.Top = tr.Top + 3
        tr.Right = tr.Right - 3
        
        If bHasFocus And m_ShowFocusRect Then
            If (m_ButtonStyle And jbsSoft) > 0 And Not bMouseDown Then
                tr.Bottom = tr.Bottom + 1
                tr.Right = tr.Right + 1
                tr.Left = tr.Left - 1
                tr.Top = tr.Top - 1
            End If
            DrawFocusRect UserControl.hDC, tr
        End If

        
        lblCaption.BackStyle = vbTransparent
        lblCaption.ForeColor = lblCaption.ForeColor
        
        
        bRedrawingControl = False
    End If
    
End Function

Public Function CurrentPic() As PictureBox
    Select Case True
        Case Not UserControl.Enabled And picDisabled.Picture.Height > 0
            Set CurrentPic = picDisabled
        Case UserControl.Enabled And bMouseCaptured And picHot.Picture.Height > 0
            Set CurrentPic = picHot
        Case Else
            Set CurrentPic = picNormal
    End Select
End Function



Public Function LoadNormal()
    Set picNormal.Picture = picNormal.Picture
End Function






Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseCaptured = False
    CaptureMouse
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    If bMouseDown Then
        bMouseDown = False
        RedrawControl
        RaiseEvent Click
    End If
        
    RedrawControl
    
End Sub


Private Sub UserControl_Paint()
    Static bPainting As Boolean
    If Not bPainting Then
        bPainting = True
        RedrawControl
        bPainting = False
    End If
End Sub

Private Sub UserControl_Resize()
    Me.RedrawControl
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property
'
Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    If InStr(New_Caption, "&") > 0 Then
        UserControl.AccessKeys = LCase(Mid(New_Caption, InStr(New_Caption, "&") + 1, 1))
    Else
        UserControl.AccessKeys = ""
    End If
    lblSize.Caption = New_Caption
    lblCaption.Visible = (New_Caption <> "")
    PropertyChanged "Caption"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
''Public Property Get ToolTipText() As String
''    ToolTipText = m_ToolTipText
''End Property
''
''Public Property Let ToolTipText(ByVal New_ToolTipText As String)
''    m_ToolTipText = New_ToolTipText
''    UserControl.Extender.ToolTipText = New_ToolTipText
''    UserControl.lblCaption.ToolTipText = New_ToolTipText
''    PropertyChanged "ToolTipText"
''End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picnormal,picnormal,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picNormal.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    If Not New_Picture Is Nothing Then
        If New_Picture.Type <> vbPicTypeIcon Then
            Exit Property
        End If
    End If
    
    Set picNormal.Picture = New_Picture
    PropertyChanged "Picture"
    Me.RedrawControl
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
    m_CaptionAlignment = m_def_CaptionAlignment
    m_ButtonStyle = m_def_ButtonStyle
    m_ShowFocusRect = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Me.Caption = PropBag.ReadProperty("Caption", "lblCaption")
'    lblSize.Caption = lblCaption.Caption
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    lblCaption.ToolTipText = m_ToolTipText
    Set picNormal.Picture = PropBag.ReadProperty("Picture", Nothing)
    Set picDisabled.Picture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Set picHot.Picture = PropBag.ReadProperty("HotPicture", Nothing)
    m_CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", m_def_CaptionAlignment)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Me.ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", -1)
'    RedrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Debug.Print "Storing pic: " & picNormal.Picture
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "lblCaption")
    Call PropBag.WriteProperty("ToolTipText", UserControl.Extender.ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("Picture", picNormal.Picture, Nothing)
    Call PropBag.WriteProperty("CaptionAlignment", m_CaptionAlignment, m_def_CaptionAlignment)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("DisabledPicture", picDisabled.Picture, Nothing)
    Call PropBag.WriteProperty("HotPicture", picHot.Picture, Nothing)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, -1)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionAlignment() As enumCaptionAlignment
    CaptionAlignment = m_CaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As enumCaptionAlignment)
    m_CaptionAlignment = New_CaptionAlignment
    PropertyChanged "CaptionAlignment"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Set lblSize.Font = New_Font
    PropertyChanged "Font"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    lblSize.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonStyle() As enumJeffButtonStyles
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As enumJeffButtonStyles)
    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled

    PropertyChanged "Enabled"
    RedrawControl
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picdisabled,picdisabled,-1,Picture
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set DisabledPicture = picDisabled.Picture
End Property
'
Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    If Not New_DisabledPicture Is Nothing Then
        If New_DisabledPicture.Type <> vbPicTypeIcon Then
            Exit Property
        End If
    End If
    Set picDisabled.Picture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
    Me.RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pichotTrack,pichotTrack,-1,Picture
Public Property Get HotPicture() As Picture
Attribute HotPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set HotPicture = picHot.Picture
End Property
'
Public Property Set HotPicture(ByVal New_HotPicture As Picture)
    If Not New_HotPicture Is Nothing Then
        If New_HotPicture.Type <> vbPicTypeIcon Then
            Exit Property
        End If
    End If
    Set picHot.Picture = New_HotPicture
    PropertyChanged "HotPicture"
    Me.RedrawControl
End Property


Public Property Get AccessKeys() As String
    AccessKeys = UserControl.AccessKeys
End Property


Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal bShow As Boolean)
    m_ShowFocusRect = bShow
    PropertyChanged "ShowFocusRect"
    Me.RedrawControl
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

