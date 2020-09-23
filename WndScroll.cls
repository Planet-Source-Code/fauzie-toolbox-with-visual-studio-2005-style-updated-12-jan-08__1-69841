VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ScrollbarConstants
    sbNone = 0&
    sbHorizontal = 1&
    sbVertical = 2&
    sbBoth = 3&
End Enum

Public Enum MouseStatusConstants
    KeyShift = MK_SHIFT
    KeyCtrl = MK_CONTROL
    LButton = MK_LBUTTON
    RButton = MK_RBUTTON
    MButton = MK_MBUTTON
End Enum

Private m_HPage As Long
Private m_VPage As Long

Private m_HMax As Long
Private m_VMax As Long

Private m_HMin As Long
Private m_VMin As Long

Private m_hWnd As Long
Private m_OldProc As Long
Private m_WndProc As Long

Private m_Hooked As Boolean

Private m_UserMode As Boolean

Private m_DisableNoScroll As Boolean

Private m_Scrollbars As ScrollbarConstants

Private m_Tracking As Boolean

Implements ISubclass

Public Event PageLeft()
Public Event PageRight()
Public Event PageUp()
Public Event PageDown()

Public Event LineLeft()
Public Event LineRight()
Public Event LineUp()
Public Event LineDown()

Public Event HThumbPosition(ByVal Position As Long)
Public Event VThumbPosition(ByVal Position As Long)

Public Event HThumbTrack(ByVal Position As Long)
Public Event VThumbTrack(ByVal Position As Long)

Public Event MouseWheel(ByVal Shift As Long, ByVal Delta As Long, ByVal x As Long, ByVal y As Long)

Public Property Get DisableNoScroll() As Boolean
  DisableNoScroll = m_DisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal vData As Boolean)
  m_DisableNoScroll = vData
  
  SetRange sbHorizontal, m_HMin, m_HMax
  SetRange sbVertical, m_VMin, m_VMax
End Property

Public Property Get HPage() As Long
  Dim siData As SCROLLINFO
  
  siData.cbSize = Len(siData)
  siData.fMask = SIF_PAGE
  
  If (m_hWnd <> 0) Then
    GetScrollInfo m_hWnd, SB_HORZ, siData
    m_HPage = siData.nPage
  End If
  
  HPage = m_HPage
End Property

Public Property Let HPage(ByVal vData As Long)
  Dim siData As SCROLLINFO
  
  siData.cbSize = Len(siData)
  siData.fMask = SIF_PAGE Or (SIF_DISABLENOSCROLL And m_DisableNoScroll)
  
  m_HPage = vData
  
  If (m_hWnd <> 0) Then
    siData.nPage = m_HPage
    SetScrollInfo m_hWnd, SB_HORZ, siData, True
  End If
End Property

Public Property Get VPage() As Long
  Dim siData As SCROLLINFO
  
  siData.cbSize = Len(siData)
  siData.fMask = SIF_PAGE
    
  If (m_hWnd <> 0) Then
    GetScrollInfo m_hWnd, SB_VERT, siData
    m_VPage = siData.nPage
  End If
  
  VPage = m_VPage
End Property

Public Property Let VPage(ByVal vData As Long)
  Dim siData As SCROLLINFO
  
  siData.cbSize = Len(siData)
  siData.fMask = SIF_PAGE Or (SIF_DISABLENOSCROLL And m_DisableNoScroll)
  
  m_VPage = vData
  
  If (m_hWnd <> 0) Then
    siData.nPage = m_VPage
    SetScrollInfo m_hWnd, SB_VERT, siData, True
  End If
End Property

Public Sub SetRange(ByVal uScroll As ScrollbarConstants, _
                    ByVal uMin As Long, ByVal uMax As Long)
  Dim siData As SCROLLINFO
  
  siData.cbSize = Len(siData)
  siData.fMask = SIF_RANGE Or (SIF_DISABLENOSCROLL And m_DisableNoScroll)
  
  If ((uScroll And sbHorizontal) <> 0&) Then
    m_HMin = uMin&
    m_HMax = uMax&
    
    If (m_hWnd <> 0&) Then
      siData.nMax = m_HMax&
      siData.nMin = m_HMin&
      
      SetScrollInfo m_hWnd, SB_HORZ, siData, True
    End If
  End If
  
  If ((uScroll And sbVertical) <> 0&) Then
    m_VMin = uMin&
    m_VMax = uMax&
    
    If (m_hWnd <> 0&) Then
      siData.nMax = m_VMax&
      siData.nMin = m_VMin&
      
      SetScrollInfo m_hWnd, SB_VERT, siData, True
    End If
  End If
End Sub

Public Property Get HMaxValue() As Long
  If (m_hWnd <> 0&) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_RANGE
    
    GetScrollInfo m_hWnd, SB_HORZ, siData
            
    m_HMax = siData.nMax
    m_HMin = siData.nMin
  End If
  
  HMaxValue = m_HMax
End Property

Public Property Get VMaxValue() As Long
  If (m_hWnd <> 0&) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_RANGE
    
    GetScrollInfo m_hWnd, SB_VERT, siData
            
    m_VMax = siData.nMax
    m_VMin = siData.nMin
  End If
  
  VMaxValue = m_VMax
End Property

Public Property Get HMinValue() As Long
  If (m_hWnd <> 0&) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_RANGE
    
    GetScrollInfo m_hWnd, SB_HORZ, siData
            
    m_HMax = siData.nMin
    m_HMin = siData.nMin
  End If
  
  HMinValue = m_HMin
End Property

Public Property Get VMinValue() As Long
  If (m_hWnd <> 0&) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_RANGE
    
    GetScrollInfo m_hWnd, SB_VERT, siData
            
    m_VMax = siData.nMin
    m_VMin = siData.nMin
  End If
  
  VMinValue = m_VMin
End Property

Public Property Get Scrollbars() As ScrollbarConstants
  Scrollbars = m_Scrollbars
End Property

Public Property Let Scrollbars(ByVal vData As ScrollbarConstants)
  m_Scrollbars = vData
      
  If (m_hWnd <> 0) Then ShowScroll
End Property

Public Property Get HTrackPos() As Long
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_TRACKPOS
    
    GetScrollInfo m_hWnd, SB_HORZ, siData
            
    HTrackPos = siData.nTrackPos
  End If
End Property

Public Property Get VTrackPos() As Long
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_TRACKPOS
    
    GetScrollInfo m_hWnd, SB_VERT, siData
            
    VTrackPos = siData.nTrackPos
  End If
End Property

Public Property Get HPosition() As Long
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_POS
    
    GetScrollInfo m_hWnd, SB_HORZ, siData
            
    HPosition = siData.nPos
  End If
End Property

Public Property Get VPosition() As Long
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_POS
    
    GetScrollInfo m_hWnd, SB_VERT, siData
            
    VPosition = siData.nPos
  End If
End Property

Public Property Let HPosition(ByVal vData As Long)
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_POS
    siData.nPos = vData
    
    SetScrollInfo m_hWnd, SB_HORZ, siData, True
  End If
End Property

Public Property Let VPosition(ByVal vData As Long)
  If (m_hWnd <> 0) Then
    Dim siData As SCROLLINFO
    
    siData.cbSize = Len(siData)
    siData.fMask = SIF_POS
    siData.nPos = vData
    
    SetScrollInfo m_hWnd, SB_VERT, siData, True
  End If
End Property

Public Property Get Hooked() As Boolean
  Hooked = m_Hooked
End Property

Public Property Get hWnd() As Long
  hWnd = m_hWnd
End Property

Public Property Get OldProc() As Long
  OldProc = m_OldProc
End Property

Public Property Get WndProc() As Long
  WndProc = m_WndProc
End Property

Private Function HookScroll(ByVal hWnd As Long) As Boolean
  On Error Resume Next
  
  If (m_Hooked = True) Then
      UnhookScroll
  End If
  
  m_OldProc = GetWindowLong(hWnd, GWL_WNDPROC)
  If (m_OldProc = 0&) Then Exit Function
  
  m_hWnd = hWnd
  
  AttachMessage Me, m_hWnd, WM_HSCROLL
  AttachMessage Me, m_hWnd, WM_VSCROLL
  AttachMessage Me, m_hWnd, WM_MOUSEWHEEL
  
  m_WndProc = GetWindowLong(hWnd, GWL_WNDPROC)
  m_Hooked = True
  
  HookScroll = True
End Function

Public Function Hook(ByVal hWnd As Long, ByVal UserMode As Boolean) As Boolean
  m_UserMode = UserMode
  
  If (m_UserMode = True) Then
    Hook = HookScroll(hWnd)
  Else
    m_hWnd = hWnd
  End If
  
  EnableScrollBar m_hWnd, SB_BOTH, 0&
  ShowScroll
End Function

Private Function UnhookScroll()
  On Error Resume Next
  
  If (m_hWnd <> 0&) And (m_OldProc <> 0&) Then
  
    If (IsBadCodePtr(m_OldProc) = 0&) And (IsWindow(m_hWnd) <> 0&) Then
        SetWindowLong m_hWnd, GWL_WNDPROC, OldProc
    End If
    
    m_OldProc = 0&
    m_WndProc = 0&
    m_hWnd = 0&
  End If
  
  m_Hooked = False
End Function

Public Sub Unhook()
  UnhookScroll
End Sub

Private Sub EventRaised(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Static HThumb As Boolean, VThumb As Boolean
  
  Dim nCode As Long, iPos As Integer
  
  nCode = (wParam And &HFFFF&)
  
  Select Case nCode
  Case SB_LINELEFT:
    If (uMsg = WM_VSCROLL) Then
      RaiseEvent LineUp
    Else
      RaiseEvent LineLeft
    End If
  
  Case SB_LINERIGHT:
    If (uMsg = WM_VSCROLL) Then
      RaiseEvent LineDown
    Else
      RaiseEvent LineRight
    End If
  
  Case SB_PAGELEFT:
    If (uMsg = WM_VSCROLL) Then
      RaiseEvent PageUp
    Else
      RaiseEvent PageLeft
    End If
  
  Case SB_PAGERIGHT:
    If (uMsg = WM_VSCROLL) Then
      RaiseEvent PageDown
    Else
      RaiseEvent PageRight
    End If
  
  Case SB_THUMBPOSITION:
    If (uMsg = WM_HSCROLL) Then
      If HThumb = True Then
        HThumb = False
        RaiseEvent HThumbPosition(HTrackPos)
      End If
    Else
      If VThumb = True Then
        VThumb = False
        RaiseEvent VThumbPosition(VTrackPos)
      End If
    End If
  
  Case SB_THUMBTRACK:
    If (uMsg = WM_HSCROLL) Then
      RaiseEvent HThumbTrack(HTrackPos)
    Else
      RaiseEvent VThumbTrack(VTrackPos)
    End If
   
  End Select
  
  If uMsg = WM_MOUSEWHEEL Then
    Dim lpPoint As POINTAPI, iDelta As Long, iShift As Long
    
    PointFromLong lParam, lpPoint
    iShift = (wParam And &HFFFF)
    
    iShift = (iShift And (MK_SHIFT + MK_CONTROL))
     
    iDelta = ((wParam And &HFFFF0000) / &H10000) And &HFFFF
    RaiseEvent MouseWheel(iShift, iDelta, lpPoint.x, lpPoint.y)
  
  End If
End Sub

Private Sub ShowScroll()
  If (m_hWnd <> 0) Then
    If (m_Scrollbars And sbHorizontal) Then
      ShowScrollBar hWnd, SB_HORZ, True
    Else
      ShowScrollBar hWnd, SB_HORZ, False
    End If
    
    If (m_Scrollbars And sbVertical) Then
      ShowScrollBar hWnd, SB_VERT, True
    Else
      ShowScrollBar hWnd, SB_VERT, False
    End If
  End If
End Sub

Private Sub Class_Terminate()
  If m_Hooked = True Then
    UnhookScroll
  End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  '
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case iMsg
  Case WM_HSCROLL, WM_VSCROLL, WM_MOUSEWHEEL
    EventRaised hWnd, iMsg, wParam, lParam
    ISubclass_WindowProc = 0
  
  Case Else
    ISubclass_WindowProc = CallWindowProc(m_OldProc, hWnd, iMsg, wParam, lParam)
  
  End Select
End Function
