VERSION 5.00
Begin VB.UserControl jeffSplitter 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   9  'Size W E
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "jeffSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumSplitterOrientation
    esoVERTICAL = 1
    esoHORIZONTAL = 2
    
End Enum

Public bMouseCaptured As Boolean
'Default Property Values:
Const m_def_SplitterMinLeft = 0
Const m_def_SplitterMinTop = 0
Const m_def_SplitterMaxWidth = 0
Const m_def_SplitterMaxHeight = 0
Const m_def_SplitterOrientation = 0
Const m_def_SplitterLeft = 0
Const m_def_SplitterTop = 0
Const m_def_SplitterWidth = 0
Const m_def_SplitterHeight = 0
'Property Variables:
Dim m_SplitterMinLeft As Single
Dim m_SplitterMinTop As Single
Dim m_SplitterMaxWidth As Single
Dim m_SplitterMaxHeight As Single
Dim m_SplitterOrientation As enumSplitterOrientation
Dim m_SplitterLeft As Single
Dim m_SplitterTop As Single
Dim m_SplitterWidth As Single
Dim m_SplitterHeight As Single

Private lPreviousCaptureHwnd As Long

Public Event SplitterMoved()


Public Function CaptureMouse()
    
    If Not bMouseCaptured Then
        Debug.Print "Capturing mouse..."
        
        lPreviousCaptureHwnd = SetCapture(UserControl.hWnd)
        bMouseCaptured = True
        'RedrawControl
        If GetCapture <> UserControl.hWnd Then
            Debug.Print "Issue"
        End If
    End If
    
End Function

Public Function ReleaseMouse()
    
    If bMouseCaptured Then
        Debug.Print "Releasing mouse"
        ReleaseCapture
        If lPreviousCaptureHwnd <> 0 Then
            'SetCapture lPreviousCaptureHwnd
            lPreviousCaptureHwnd = 0
        End If
        bMouseCaptured = False
        
    End If
    
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SplitterOrientation() As enumSplitterOrientation
    SplitterOrientation = m_SplitterOrientation
End Property

Public Property Let SplitterOrientation(ByVal New_SplitterOrientation As enumSplitterOrientation)
    m_SplitterOrientation = New_SplitterOrientation
    Select Case True
        Case New_SplitterOrientation = esoHORIZONTAL
            UserControl.MousePointer = 7
        Case New_SplitterOrientation = esoVERTICAL
            UserControl.MousePointer = 9
    End Select
    PropertyChanged "SplitterOrientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterLeft() As Single
    SplitterLeft = UserControl.Extender.Left
End Property

Public Property Let SplitterLeft(ByVal New_SplitterLeft As Single)
    If m_SplitterOrientation = esoVERTICAL Then
        If New_SplitterLeft < Me.SplitterMinLeft Then
            New_SplitterLeft = m_SplitterMinLeft
        End If
        If New_SplitterLeft + UserControl.Extender.Width > m_SplitterMaxWidth Then
            New_SplitterLeft = m_SplitterMaxWidth - Me.SplitterWidth
        End If
    End If
    UserControl.Extender.Left = New_SplitterLeft
    PropertyChanged "SplitterLeft"
    If m_SplitterOrientation = esoVERTICAL Then
        RaiseEvent SplitterMoved
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterTop() As Single
    SplitterTop = UserControl.Extender.Top
End Property

Public Property Let SplitterTop(ByVal New_SplitterTop As Single)
    
    If m_SplitterOrientation = esoHORIZONTAL Then
        If New_SplitterTop < Me.SplitterMinTop Then
            New_SplitterTop = m_SplitterMinTop
        End If
        If New_SplitterTop + Me.SplitterHeight > m_SplitterMaxHeight Then
            New_SplitterTop = m_SplitterMaxHeight - Me.SplitterHeight
        End If
    End If
    UserControl.Extender.Top = New_SplitterTop
    PropertyChanged "SplitterTop"
    If m_SplitterOrientation = esoHORIZONTAL Then
        RaiseEvent SplitterMoved
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterWidth() As Single
    SplitterWidth = UserControl.Extender.Width
End Property

Public Property Let SplitterWidth(ByVal New_SplitterWidth As Single)
    UserControl.Extender.Width = New_SplitterWidth
    PropertyChanged "SplitterWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterHeight() As Single
    SplitterHeight = UserControl.Extender.Height
End Property

Public Property Let SplitterHeight(ByVal New_SplitterHeight As Single)
    UserControl.Extender.Height = New_SplitterHeight
    PropertyChanged "SplitterHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SplitterOrientation = m_def_SplitterOrientation
    m_SplitterLeft = m_def_SplitterLeft
    m_SplitterTop = m_def_SplitterTop
    m_SplitterWidth = m_def_SplitterWidth
    m_SplitterHeight = m_def_SplitterHeight
    m_SplitterMinLeft = m_def_SplitterMinLeft
    m_SplitterMinTop = m_def_SplitterMinTop
    m_SplitterMaxWidth = m_def_SplitterMaxWidth
    m_SplitterMaxHeight = m_def_SplitterMaxHeight
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CaptureMouse
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static bMouseMoving As Boolean
    If Not bMouseMoving Then
        bMouseMoving = True
        If bMouseCaptured Then
            Select Case True
                Case SplitterOrientation = esoVERTICAL
                    Me.SplitterLeft = Me.SplitterLeft + X
                    
                Case SplitterOrientation = esoHORIZONTAL
                    Me.SplitterTop = Me.SplitterTop + Y
                    
            End Select
            
        End If
        bMouseMoving = False
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseMouse
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Me.SplitterOrientation = PropBag.ReadProperty("SplitterOrientation", m_def_SplitterOrientation)
    m_SplitterLeft = PropBag.ReadProperty("SplitterLeft", m_def_SplitterLeft)
    m_SplitterTop = PropBag.ReadProperty("SplitterTop", m_def_SplitterTop)
    m_SplitterWidth = PropBag.ReadProperty("SplitterWidth", m_def_SplitterWidth)
    m_SplitterHeight = PropBag.ReadProperty("SplitterHeight", m_def_SplitterHeight)
    m_SplitterMinLeft = PropBag.ReadProperty("SplitterMinLeft", m_def_SplitterMinLeft)
    m_SplitterMinTop = PropBag.ReadProperty("SplitterMinTop", m_def_SplitterMinTop)
    m_SplitterMaxWidth = PropBag.ReadProperty("SplitterMaxWidth", m_def_SplitterMaxWidth)
    m_SplitterMaxHeight = PropBag.ReadProperty("SplitterMaxHeight", m_def_SplitterMaxHeight)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80C0FF)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SplitterOrientation", m_SplitterOrientation, m_def_SplitterOrientation)
    Call PropBag.WriteProperty("SplitterLeft", m_SplitterLeft, m_def_SplitterLeft)
    Call PropBag.WriteProperty("SplitterTop", m_SplitterTop, m_def_SplitterTop)
    Call PropBag.WriteProperty("SplitterWidth", m_SplitterWidth, m_def_SplitterWidth)
    Call PropBag.WriteProperty("SplitterHeight", m_SplitterHeight, m_def_SplitterHeight)
    Call PropBag.WriteProperty("SplitterMinLeft", m_SplitterMinLeft, m_def_SplitterMinLeft)
    Call PropBag.WriteProperty("SplitterMinTop", m_SplitterMinTop, m_def_SplitterMinTop)
    Call PropBag.WriteProperty("SplitterMaxWidth", m_SplitterMaxWidth, m_def_SplitterMaxWidth)
    Call PropBag.WriteProperty("SplitterMaxHeight", m_SplitterMaxHeight, m_def_SplitterMaxHeight)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80C0FF)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterMinLeft() As Single
    SplitterMinLeft = m_SplitterMinLeft
End Property

Public Property Let SplitterMinLeft(ByVal New_SplitterMinLeft As Single)
    m_SplitterMinLeft = New_SplitterMinLeft
    PropertyChanged "SplitterMinLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterMinTop() As Single
    SplitterMinTop = m_SplitterMinTop
End Property

Public Property Let SplitterMinTop(ByVal New_SplitterMinTop As Single)
    m_SplitterMinTop = New_SplitterMinTop
    PropertyChanged "SplitterMinTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterMaxWidth() As Single
    SplitterMaxWidth = m_SplitterMaxWidth
End Property

Public Property Get SplitterPosRatio() As Double
    On Error Resume Next
    Select Case True
        Case m_SplitterOrientation = esoVERTICAL
            SplitterPosRatio = (Me.SplitterLeft - m_SplitterMinLeft) / (Me.SplitterMaxWidth - m_SplitterMinLeft)
        Case m_SplitterOrientation = esoHORIZONTAL
            SplitterPosRatio = (Me.SplitterTop - m_SplitterMinTop) / (Me.SplitterMaxHeight - m_SplitterMinTop)
    End Select
End Property

Public Property Let SplitterPosRatio(ByVal dNew As Double)
    Select Case True
        Case Me.SplitterOrientation = esoVERTICAL
            Me.SplitterLeft = ((Me.SplitterMaxWidth - Me.SplitterMinLeft) * dNew) + Me.SplitterMinLeft
        Case Me.SplitterOrientation = esoHORIZONTAL
            Me.SplitterTop = ((Me.SplitterMaxHeight - Me.SplitterMinTop) * dNew) + Me.SplitterMinTop
    End Select
End Property

Public Property Let SplitterMaxWidth(ByVal New_SplitterMaxWidth As Single)
    ' Generate Current Percentage
    Dim dCurr As Double
    dCurr = Me.SplitterPosRatio
    m_SplitterMaxWidth = New_SplitterMaxWidth
    PropertyChanged "SplitterMaxWidth"
    Me.SplitterPosRatio = dCurr
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get SplitterMaxHeight() As Single
    SplitterMaxHeight = m_SplitterMaxHeight
End Property

Public Property Let SplitterMaxHeight(ByVal New_SplitterMaxHeight As Single)
    
    Dim dCurr As Double
    dCurr = Me.SplitterPosRatio
    m_SplitterMaxHeight = New_SplitterMaxHeight
    PropertyChanged "SplitterMaxHeight"
    Me.SplitterPosRatio = dCurr
    
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
End Property

