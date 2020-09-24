VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H80000018&
   ClientHeight    =   864
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   864
   ScaleWidth      =   3840
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   336
      TabIndex        =   0
      Text            =   "This is a constituent control"
      Top             =   168
      Width           =   2280
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' API
'=========================================================================

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_KEYUP          As Long = &H101
Private Const WM_CHAR           As Long = &H102

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_PROCESSTAB        As Boolean = False
Private Const DEF_PROCESSENTER      As Boolean = False
Private Const DEF_PROCESSESCAPE     As Boolean = False
Private Const DEF_LOCKED            As Boolean = False
Private Const DEF_CONSITUENT        As Boolean = False

Private m_uIPAO             As IPAOHookStruct
Private m_bProcessTab       As Boolean
Private m_bProcessEnter     As Boolean
Private m_bProcessEscape    As Boolean
Private m_bLocked           As Boolean
Private m_bConstituent      As Boolean

'=========================================================================
' Properties
'=========================================================================

Property Get ProcessTab() As Boolean
    ProcessTab = m_bProcessTab
End Property

Property Let ProcessTab(ByVal bValue As Boolean)
    m_bProcessTab = bValue
    Refresh
    PropertyChanged
End Property

Property Get ProcessEnter() As Boolean
    ProcessEnter = m_bProcessEnter
End Property

Property Let ProcessEnter(ByVal bValue As Boolean)
    m_bProcessEnter = bValue
    PropertyChanged
End Property

Property Get ProcessEscape() As Boolean
    ProcessEscape = m_bProcessEscape
End Property

Property Let ProcessEscape(ByVal bValue As Boolean)
    m_bProcessEscape = bValue
    PropertyChanged
End Property

Property Get Locked() As Boolean
    Locked = m_bLocked
End Property

Property Let Locked(ByVal bValue As Boolean)
    m_bLocked = bValue
    PropertyChanged
End Property

Property Get ConstituentControl() As Boolean
    ConstituentControl = m_bConstituent
End Property

Property Let ConstituentControl(ByVal bValue As Boolean)
    m_bConstituent = bValue
    Text1.Visible = bValue
    PropertyChanged
End Property

'=========================================================================
' Methods
'=========================================================================

Private Sub pvSetIPAO()
    Dim pOleObject          As IOleObject
    Dim pOleInPlaceSite     As IOleInPlaceSite
    Dim pOleInPlaceFrame    As IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
    Dim rcPos               As RECT
    Dim rcClip              As RECT
    Dim uFrameInfo          As OLEINPLACEFRAMEINFO
       
    On Error Resume Next
    Set pOleObject = Me
    Set pOleInPlaceSite = pOleObject.GetClientSite
    If Not pOleInPlaceSite Is Nothing Then
        pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)
        If Not pOleInPlaceFrame Is Nothing Then
            pOleInPlaceFrame.SetActiveObject m_uIPAO.ThisPointer, vbNullString
        End If
        If Not pOleInPlaceUIWindow Is Nothing Then '-- And Not m_bMouseActivate
            pOleInPlaceUIWindow.SetActiveObject m_uIPAO.ThisPointer, vbNullString
        Else
            pOleObject.DoVerb OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos)
        End If
    End If
End Sub

Private Function KeyIsPressed(lVirtKey As KeyCodeConstants) As Boolean
    KeyIsPressed = ((GetKeyState(lVirtKey) And &H8000) = &H8000)
End Function

Private Function GetShiftState() As Long
    GetShiftState = (-KeyIsPressed(vbKeyShift) * vbShiftMask) _
                Or (-KeyIsPressed(vbKeyControl) * vbCtrlMask) _
                Or (-KeyIsPressed(vbKeyMenu) * vbAltMask)
End Function

Friend Function frTranslateAccel(pMsg As MSG) As Boolean
    Dim pOleObject      As IOleObject
    Dim pOleControlSite As IOleControlSite
    
    On Error Resume Next
    Select Case pMsg.message
    Case WM_KEYDOWN, WM_KEYUP
        Select Case pMsg.wParam
        Case vbKeyTab
            '--- ctrl+tab ALWAYS moves focus away
            If (GetShiftState() And vbCtrlMask) <> 0 Then
                Set pOleObject = Me
                Set pOleControlSite = pOleObject.GetClientSite
                If Not pOleControlSite Is Nothing Then
                    pOleControlSite.TranslateAccelerator VarPtr(pMsg), GetShiftState() And vbShiftMask
                End If
            End If
            ' Check whether the control processes the Tab key
            If ProcessTab Then
                Debug.Print "vbKeyTab " & IIf(pMsg.message = WM_KEYDOWN, "pressed", "released") & "! "; Timer
                ' Ignore the message
                frTranslateAccel = True
            End If
        Case vbKeyReturn
            ' Check whether the control processes the Enter key
            If ProcessEnter Then
                Debug.Print "vbKeyReturn " & IIf(pMsg.message = WM_KEYDOWN, "pressed", "released") & "! "; Timer
                ' Ignore the message
                frTranslateAccel = True
            End If
        Case vbKeyEscape
            If ProcessEscape Then
                Debug.Print "vbKeyEscape " & IIf(pMsg.message = WM_KEYDOWN, "pressed", "released") & "! "; Timer
                ' Ignore the message
                frTranslateAccel = True
            End If
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, _
             vbKeyEnd, vbKeyPageDown, vbKeyPageUp
            ' Navigation keys filter
        Case Else
            ' If the control is read-only
            ' eat every other key
            If Locked Then frTranslateAccel = True
        End Select
    Case WM_CHAR
        ' If the control is read-only
        ' eat every char
        If Locked Then frTranslateAccel = True
    End Select
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub Text1_GotFocus()
    pvSetIPAO
    Refresh
End Sub

Private Sub Text1_LostFocus()
    Refresh
End Sub

Private Sub UserControl_GotFocus()
    pvSetIPAO
    Refresh
End Sub

Private Sub UserControl_LostFocus()
    Refresh
End Sub

Private Sub UserControl_Paint()
    Select Case GetFocus()
    Case hWnd, Text1.hWnd
        Cls
        If ProcessTab Then
            Print
            Print "   Use Tab to navigate columns :-))"
            Print "   Use Ctrl+Tab to move focus away"
        Else
            Print
            Print "   Use Tab to move to next control"
        End If
        Dim rc As RECT
        With rc
            .Left = 2
            .Top = 2
            .Right = ScaleWidth / Screen.TwipsPerPixelX - 2
            .Bottom = ScaleHeight / Screen.TwipsPerPixelY - 2
        End With
        DrawFocusRect hdc, rc
    End Select
End Sub

Private Sub UserControl_InitProperties()
    ProcessTab = DEF_PROCESSTAB
    ProcessEnter = DEF_PROCESSENTER
    ProcessEscape = DEF_PROCESSESCAPE
    Locked = DEF_LOCKED
    ConstituentControl = DEF_CONSITUENT
    If Ambient.UserMode Then
        InitIPAO m_uIPAO, Me
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        ProcessTab = .ReadProperty("ProcessTab", DEF_PROCESSTAB)
        ProcessEnter = .ReadProperty("ProcessEnter", DEF_PROCESSENTER)
        ProcessEscape = .ReadProperty("ProcessEscape", DEF_PROCESSESCAPE)
        Locked = .ReadProperty("Locked", DEF_LOCKED)
        ConstituentControl = .ReadProperty("ConstituentControl", DEF_CONSITUENT)
    End With
    If Ambient.UserMode Then
        InitIPAO m_uIPAO, Me
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Text1.Move (ScaleWidth - Text1.Width) \ 2, (ScaleHeight - Text1.Height) \ 2
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ProcessTab", ProcessTab, DEF_PROCESSTAB)
        Call .WriteProperty("ProcessEnter", ProcessEnter, DEF_PROCESSENTER)
        Call .WriteProperty("ProcessEscape", ProcessEscape, DEF_PROCESSESCAPE)
        Call .WriteProperty("Locked", Locked, DEF_LOCKED)
        Call .WriteProperty("ConstituentControl", ConstituentControl, DEF_CONSITUENT)
    End With
End Sub

Private Sub UserControl_Terminate()
    Debug.Print "UserControl_Terminate "; Timer
    TerminateIPAO m_uIPAO
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "UserControl_KeyDown "; KeyCode; Shift; Timer
End Sub


