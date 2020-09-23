VERSION 5.00
Begin VB.UserControl ToolBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   FillStyle       =   0  'Solid
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ToolboxBitmap   =   "ToolBox.ctx":0000
End
Attribute VB_Name = "ToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WM_TIMER = &H113
Private Const WM_MOUSEMOVE = &H200

Private Type tChildsRect
  Rct() As RECT
  CaptRct() As RECT
End Type

Public Enum eBorderStyle
  [None] = 0
  [Fixed Single] = 1
End Enum

Public Enum eAppearance
  [Flat] = 0
  [3D] = 1
End Enum

Private Const ChildHeight = 22
Private Const NodeHeight = 15

' Private variables
Dim PNT As POINTAPI
Dim NodeColor() As Long
Dim HColor() As Long
Dim LChild, LSel
Dim SelectedNode As Integer
Dim NodesRct() As RECT
Dim NodesCaptRct() As RECT
Dim ChildsRct() As tChildsRect
Dim TheHeight As Long
Private m_bMouseOver       As Boolean

Private WithEvents WS As WndScroll
Attribute WS.VB_VarHelpID = -1

Implements ISubclass

'Constants
Const m_def_ForeColor = vbButtonText
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event NodeClick(Node As Nodes)
Event ChildClick(ParentNode As Nodes, Child As Child)
Event MouseOverItem(ParentNode As Nodes, Child As Child)
'Property Variables:
Dim m_NodeCount As Integer
Dim m_ForeColor As OLE_COLOR
Dim m_Nodes() As Nodes

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
  frmAbout.Show vbModal
End Sub

Public Function AddNode(nCaption As String, Optional nKey As String, Optional nDescription As String, Optional nToolTipText As String) As Integer
  m_NodeCount = m_NodeCount + 1
  ReDim Preserve m_Nodes(m_NodeCount) As Nodes
  Set m_Nodes(m_NodeCount) = New Nodes
  With m_Nodes(m_NodeCount)
    .Caption = nCaption
    .Description = nDescription
    .Key = nKey
    .ToolTipText = nToolTipText
  End With
  PropertyChanged "Nodes" & m_NodeCount
  AddNode = m_NodeCount
  DrawAll
End Function

Private Sub DrawAll()
  Dim i As Integer, j As Integer
  LChild = 0
  If m_NodeCount > 0 Then
    UserControl.Cls
    SetRects
    SetScroll
    For i = 1 To m_NodeCount
      ' Draw nodes
      If NodesRct(i).Top < UserControl.ScaleHeight And NodesRct(i).Bottom > 0 Then
        DrawNode i
      End If
      
      ' Draw node childs
      If Not m_Nodes(i).Expanded Then m_Nodes(i).SelectedChild = 0
      If m_Nodes(i).ChildCount > 0 And m_Nodes(i).Expanded Then
        For j = 1 To m_Nodes(i).ChildCount
          If ChildsRct(i).Rct(j).Top < UserControl.ScaleHeight And ChildsRct(i).Rct(j).Bottom > 0 Then
            DrawChild i, j
          End If
        Next
      End If
    Next
    UserControl.Refresh
  End If
End Sub

Private Sub DrawNode(nNode As Integer)
  Dim a As Integer, PlusMin As RECT
  
  If nNode <> SelectedNode Then
    For a = 0 To NodeHeight - 1
      UserControl.Line (NodesRct(nNode).Left, NodesRct(nNode).Top + a)-(NodesRct(nNode).Right, NodesRct(nNode).Top + a), NodeColor(a)
    Next
    UserControl.Line (NodesRct(nNode).Left, NodesRct(nNode).Bottom + 1)-(NodesRct(nNode).Right, NodesRct(nNode).Bottom + 1), NodeColor(0)
  Else
    UserControl.ForeColor = vbHighlight
    UserControl.FillColor = HColor(13)
    Rectangle UserControl.hdc, NodesRct(nNode).Left, NodesRct(nNode).Top, NodesRct(nNode).Right, NodesRct(nNode).Bottom + 2
  End If
  
  SetRect PlusMin, NodesRct(nNode).Left + 5, NodesRct(nNode).Top + 3, NodesRct(nNode).Left + 14, NodesRct(nNode).Top + 12
  If Not DrawTheme(UserControl.hdc, "TreeView", 2, IIf(m_Nodes(nNode).Expanded, 2, 1), PlusMin) Then _
  PE.PaintTransparentPicture UserControl.hdc, LoadResPicture(IIf(m_Nodes(nNode).Expanded, "MINUS", "PLUS"), vbResBitmap), NodesRct(nNode).Left + 5, NodesRct(nNode).Top + 3, 9, 9
  UserControl.ForeColor = m_ForeColor
  UserControl.FontBold = True
  DrawText UserControl.hdc, m_Nodes(nNode).Caption, Len(m_Nodes(nNode).Caption), NodesCaptRct(nNode), DT_SINGLELINE Or DT_VCENTER
End Sub

Private Sub DrawChild(nNode As Integer, nChild As Integer, Optional nHover As Boolean = False)
  Dim PX As Single, PY As Single
  Dim PW As Single, PH As Single
  
  If nHover Then
    UserControl.ForeColor = vbHighlight
    If m_Nodes(nNode).SelectedChild = nChild Then
      UserControl.FillColor = HColor(8)
    Else
      UserControl.FillColor = HColor(11)
    End If
  Else
    If m_Nodes(nNode).SelectedChild = nChild Then
      UserControl.ForeColor = vbHighlight
      UserControl.FillColor = HColor(13)
    Else
      UserControl.ForeColor = BackColor
      UserControl.FillColor = BackColor
    End If
  End If
  
  Rectangle UserControl.hdc, ChildsRct(nNode).Rct(nChild).Left, ChildsRct(nNode).Rct(nChild).Top, ChildsRct(nNode).Rct(nChild).Right, ChildsRct(nNode).Rct(nChild).Bottom
  
  UserControl.ForeColor = m_ForeColor
  UserControl.FontBold = False
  DrawText UserControl.hdc, m_Nodes(nNode).Childs(nChild).Caption, Len(m_Nodes(nNode).Childs(nChild).Caption), ChildsRct(nNode).CaptRct(nChild), DT_SINGLELINE Or DT_VCENTER
  
  With m_Nodes(nNode).Childs(nChild)
    ' Draw item icon
    If Not .Icon Is Nothing And m_Nodes(nNode).Expanded Then
      PW = ScaleX(.Icon.Width, vbHimetric, vbPixels)
      PH = ScaleY(.Icon.Height, vbHimetric, vbPixels)
      PX = ChildsRct(nNode).Rct(nChild).Left + (25 - PW) / 2
      PY = ChildsRct(nNode).Rct(nChild).Top + (ChildHeight - PH) / 2
      If .Icon.Type = vbPicTypeIcon Then
        'DrawTransparentBitmap doesn't support icons
        PE.PaintStandardPicture UserControl.hdc, .Icon, PX, PY, PW, PH, 0, 0
      Else
        If .UseMaskColor Then
          PE.PaintTransparentPicture UserControl.hdc, .Icon, PX, PY, PW, PH, 0, 0, .MaskColor
        Else
          PE.PaintStandardPicture UserControl.hdc, .Icon, PX, PY, PW, PH, 0, 0
        End If
      End If
    End If
  End With
End Sub

Private Sub SetRects()
  Dim i As Integer, j As Integer
  If m_NodeCount > 0 Then
    ReDim Preserve NodesRct(m_NodeCount) As RECT
    ReDim Preserve NodesCaptRct(m_NodeCount) As RECT
    ReDim Preserve ChildsRct(m_NodeCount) As tChildsRect
    For i = 1 To m_NodeCount
      ReDim Preserve ChildsRct(i).Rct(m_Nodes(i).ChildCount) As RECT
      ReDim Preserve ChildsRct(i).CaptRct(m_Nodes(i).ChildCount) As RECT
      
      If i > 1 Then
        If m_Nodes(i - 1).Expanded Then
          SetRect NodesRct(i), 2, NodesRct(i - 1).Bottom + 3 + (m_Nodes(i - 1).ChildCount * ChildHeight), UserControl.ScaleWidth - 2, 0
        Else
          SetRect NodesRct(i), 2, NodesRct(i - 1).Bottom + 3, UserControl.ScaleWidth - 2, 0
        End If
      ElseIf i = 1 Then
        SetRect NodesRct(i), 2, 2, UserControl.ScaleWidth - 2, 0
      End If
      NodesRct(i).Bottom = NodesRct(i).Top + NodeHeight
      
      For j = 1 To m_Nodes(i).ChildCount
        If m_Nodes(i).Expanded Then
          If j = 1 Then
            SetRect ChildsRct(i).Rct(j), 5, NodesRct(i).Bottom + 2, UserControl.ScaleWidth - 3, 0
          Else
            SetRect ChildsRct(i).Rct(j), 5, NodesRct(i).Bottom + 2 + (j - 1) * ChildHeight, UserControl.ScaleWidth - 3, 0
          End If
          ChildsRct(i).Rct(j).Bottom = ChildsRct(i).Rct(j).Top + ChildHeight
          SetRect ChildsRct(i).CaptRct(j), ChildsRct(i).Rct(j).Left + 25, ChildsRct(i).Rct(j).Top, ChildsRct(i).Rct(j).Right, ChildsRct(i).Rct(j).Bottom
        Else
          SetRect ChildsRct(i).Rct(j), 0, 0, 0, 0
          SetRect ChildsRct(i).CaptRct(j), 0, 0, 0, 0
        End If
      Next
      
      SetRect NodesCaptRct(i), NodesRct(i).Left + 20, NodesRct(i).Top, NodesRct(i).Right, NodesRct(i).Bottom
    Next
    If m_Nodes(m_NodeCount).ChildCount > 0 Then
      TheHeight = IIf(m_Nodes(m_NodeCount).Expanded, 0, NodesRct(m_NodeCount).Bottom) + 4 + ChildsRct(m_NodeCount).Rct(m_Nodes(m_NodeCount).ChildCount).Bottom + 1
    Else
      TheHeight = NodesRct(m_NodeCount).Bottom + 4
    End If
  End If
End Sub

Public Sub ExpandNode(Index As Variant)
  Index = KeyToIndex(Index)
  m_Nodes(Index).Expanded = True
  DrawAll
End Sub

Public Sub CollapseNode(Index As Variant)
  Index = KeyToIndex(Index)
  m_Nodes(Index).Expanded = False
  m_Nodes(Index).SelectedChild = 0
  DrawAll
End Sub

Private Function MatchRect(ByVal x As Long, ByVal y As Long)
  Dim Rct As RECT
  Dim i, j, Ok As Long
  
  Ok = GetClientRect(UserControl.hWnd, Rct)
  If PtInRect(Rct, x, y) <> 0 Then
    If m_NodeCount <> 0 Then
      For i = m_NodeCount To 1 Step -1
        If PtInRect(NodesRct(i), x, y) Then
          MatchRect = i
          Exit Function
        End If
        If m_Nodes(i).Expanded Then
          For j = m_Nodes(i).ChildCount To 1 Step -1
            If PtInRect(ChildsRct(i).Rct(j), x, y) Then
              MatchRect = Format(i, "00") & ", " & Format(j, "00")
              Exit Function
            End If
          Next
        End If
      Next
    End If
  
  End If
  
  MatchRect = -1
End Function

Private Function KeyToIndex(ByVal Index As Variant) As Integer
 ' Returns the integer index value of a button identified by either it's
 ' Key property or index.
 Dim i
 If Val(Index) = 0 Then
  For i = 1 To m_NodeCount
   If UCase$(m_Nodes(i).Key) = UCase$(Index) Then
    KeyToIndex = i
    Exit Function
   End If
  Next
 Else
  KeyToIndex = Val(Index)
 End If
 If KeyToIndex = 0 And Index <> 0 Then
  Err.Raise 35601, "Toolbox.KeyToIndex", "Element not found. Key is missing or illegal."
 End If
End Function

Public Property Get Appearance() As eAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
  Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eAppearance)
  UserControl.Appearance() = New_Appearance
  PropertyChanged "Appearance"
  UserControl.BackColor = UserControl.BackColor
End Property

Public Property Get BorderStyle() As eBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As eBorderStyle)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
   DrawAll
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  ' don't remove this comment
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  ' don't remove this comment
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If hWnd = UserControl.hWnd Then
    Select Case iMsg
    Case WM_MOUSEMOVE
      If Not (m_bMouseOver) Then
        m_bMouseOver = True
        ' Start checking to see if mouse is no longer over.
        SetTimer UserControl.hWnd, 1, 10, 0
      End If
      ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
      
    Case WM_TIMER
      OnTimer
    
    End Select
  End If
End Function

Private Sub OnTimer()
  Dim bOver As Boolean
  Dim rcItem As RECT
  Dim tp As POINTAPI

  bOver = True
  GetCursorPos tp
  GetWindowRect UserControl.hWnd, rcItem
  If (PtInRect(rcItem, tp.x, tp.y) = 0) Then
    bOver = False
    If LChild <> 0 Then DrawChild Val(Left(LChild, 2)), Val(Mid(LChild, 5, 2))
    LChild = 0 ': DrawAll
    UserControl.Refresh
  End If

  If Not (bOver) Then
    KillTimer UserControl.hWnd, 1
    m_bMouseOver = False
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim MR, a, B, i
  Dim R
  If Button = 1 Then
    MR = MatchRect(x, y)
    
    SelectedNode = 0
    
    If LSel <> -1 And LSel <> "" Then
      If Len(LSel) = 6 Then
        a = Val(Left(LSel, 2))
        B = Val(Mid(LSel, 5, 2))
        
        m_Nodes(a).SelectedChild = 0
        DrawChild Int(a), Int(B)
      Else
        m_Nodes(LSel).SelectedChild = 0
        DrawNode Int(LSel)
      End If
    End If
    
    If Len(MR) = 6 Then
      a = Val(Left(MR, 2))
      B = Val(Mid(MR, 5, 2))
      
      m_Nodes(a).SelectedChild = B
      
      If ChildsRct(a).Rct(B).Bottom > UserControl.ScaleHeight Then
        Do
          WS.VPosition = WS.VPosition + 2
          DrawAll
        Loop Until ChildsRct(a).Rct(B).Bottom < UserControl.ScaleHeight
      ElseIf ChildsRct(a).Rct(B).Top < 0 Then
        Do
          WS.VPosition = WS.VPosition - 2
          DrawAll
        Loop Until ChildsRct(a).Rct(B).Top > 0
      End If
      DrawChild Int(a), Int(B), True
      
      UserControl.Refresh
      RaiseEvent ChildClick(m_Nodes(a), m_Nodes(a).Childs(B))
    ElseIf MR > 0 Then
      SelectedNode = MR
      DrawNode Int(MR)
    End If
    LSel = MR
  End If
  UserControl.Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim MR, a, B, i
  
  If Button <> 1 And Ambient.UserMode Then
    MR = MatchRect(x, y)
    If MR = -1 Then Extender.ToolTipText = "": DrawAll: Exit Sub
      
    If Len(LChild) = 6 And LChild <> MR Then
      a = Val(Left(LChild, 2))
      B = Val(Mid(LChild, 5, 2))
      
      DrawChild Int(a), Int(B)
    End If
    If Len(MR) = 6 And MR <> LChild Then
      a = Val(Left(MR, 2))
      B = Val(Mid(MR, 5, 2))
      
      DrawChild Int(a), Int(B), True
      
      LChild = MR
      Extender.ToolTipText = m_Nodes(a).Childs(B).ToolTipText
      RaiseEvent MouseOverItem(m_Nodes(a), m_Nodes(a).Childs(B))
    ElseIf Len(MR) < 6 And MR > 0 Then
      Extender.ToolTipText = m_Nodes(MR).ToolTipText
      LChild = 0
    End If
    UserControl.Refresh
  End If
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If SelectedNode <> 0 Then
      m_Nodes(SelectedNode).Expanded = Not m_Nodes(SelectedNode).Expanded
      RaiseEvent NodeClick(m_Nodes(SelectedNode))
      DrawAll
'      If ChildsRct(SelectedNode).Rct(m_Nodes(SelectedNode).ChildCount).Bottom > UserControl.ScaleHeight And NodesRct(SelectedNode).Top > 0 Then
'        Do Until ChildsRct(SelectedNode).Rct(m_Nodes(SelectedNode).ChildCount).Bottom < UserControl.ScaleHeight Or NodesRct(SelectedNode).Top <= 2
'          WS.VPosition = WS.VPosition + 5
'          DrawAll
'          DoEvents
'        Loop
'      End If
    ElseIf Len(LChild) = 6 Then
      DrawChild Left(LChild, 2), Mid(LChild, 5, 2)
    End If
    LChild = 0
  End If
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_ExitFocus()
  SelectedNode = 0
  DrawAll
End Sub

Private Sub UserControl_Initialize()
  ' Hook the scrollbar
  Set WS = New WndScroll
  WS.Hook UserControl.hWnd, True
  
  InitPaintEffects
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
  hWnd = UserControl.hWnd
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim i As Integer
  
  UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set UserControl.Font = UserControl.Font
  m_NodeCount = PropBag.ReadProperty("NodeCount", 0)
  ReDim Preserve m_Nodes(m_NodeCount) As Nodes
  For i = 1 To m_NodeCount
    Set m_Nodes(i) = PropBag.ReadProperty("Nodes" & i, Nothing)
  Next

  ' Prepare the colors
  Dim Col1 As Long, Col2 As Long
  Col1 = BlendColor(vbWindowBackground, TranslateColor(UserControl.BackColor), 220) 'TranslateColor(UserControl.BackColor)
  Col2 = BlendColor(TranslateColor(UserControl.BackColor), Col1, 195) 'TranslateColor(UserControl.BackColor)
'  Col2 = Col1
'  Brighten2 Col1, 0.07
'  Brighten2 Col2, 0.05
'  Darken2 Col2, 0.05
  BlendColors Col1, Col2, NodeHeight, NodeColor()
  BlendColors TranslateColor(vbHighlight), TranslateColor(Col1), 15, HColor()

  If Ambient.UserMode Then
    AttachMessage Me, UserControl.hWnd, WM_MOUSEMOVE
    AttachMessage Me, UserControl.hWnd, WM_TIMER
  End If
End Sub

Private Sub UserControl_Resize()
  DrawAll
End Sub

Private Sub UserControl_Terminate()
  ' Unhook the scrollbar
  WS.Unhook
  Set WS = Nothing
  
  DetachMessage Me, UserControl.hWnd, WM_MOUSEMOVE
  DetachMessage Me, UserControl.hWnd, WM_TIMER
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Dim i As Integer
  
  Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("NodeCount", m_NodeCount)
  For i = 1 To m_NodeCount
    Call PropBag.WriteProperty("Nodes" & i, m_Nodes, Nothing)
  Next
End Sub

Public Property Get NodeCount() As Integer
Attribute NodeCount.VB_MemberFlags = "400"
  NodeCount = m_NodeCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get Nodes(ByVal Index As Variant) As Nodes
Attribute Nodes.VB_MemberFlags = "400"
  Index = KeyToIndex(Index)
  Set Nodes = m_Nodes(Index)
End Property

Public Property Set Nodes(ByVal Index As Variant, ByVal New_Nodes As Nodes)
  Index = KeyToIndex(Index)
  Set m_Nodes(Index) = New_Nodes
  PropertyChanged "Nodes" & Index
  DrawAll
End Property

Private Sub WS_LineDown()
  WS.VPosition = WS.VPosition + 20
  DrawAll
End Sub

Private Sub WS_LineUp()
  WS.VPosition = WS.VPosition - 20
  DrawAll
End Sub

Private Sub WS_MouseWheel(ByVal Shift As Long, ByVal Delta As Long, ByVal x As Long, ByVal y As Long)
  Dim wfp As Long
  If WS.Scrollbars = sbNone Then Exit Sub
  wfp = WindowFromPoint(x, y)
  If wfp = UserControl.hWnd Or wfp = UserControl.hWnd Then
    WS.VPosition = WS.VPosition + IIf(Delta > 0, -50, 50)
    DrawAll
  End If
End Sub

Private Sub WS_PageDown()
  WS.VPosition = WS.VPosition + (WS.VPage \ 2)
  DrawAll
End Sub '

Private Sub WS_PageUp()
  WS.VPosition = WS.VPosition - (WS.VPage \ 2)
  DrawAll
End Sub

Private Sub WS_VThumbTrack(ByVal Position As Long)
  WS.VPosition = Position
  DrawAll
End Sub

Private Sub OffsetAll(LastPos As Long, NewPos As Long)
  Dim i As Integer, j As Integer
  For i = 1 To m_NodeCount
    NodesRct(i).Right = ScaleWidth - 2
    OffsetRect NodesRct(i), 0, -(NewPos - LastPos)
    NodesCaptRct(i).Right = ScaleWidth - 2
    OffsetRect NodesCaptRct(i), 0, -(NewPos - LastPos)
    
    For j = 1 To m_Nodes(i).ChildCount
      If m_Nodes(i).Expanded Then
        ChildsRct(i).Rct(j).Right = ScaleWidth - 2
        ChildsRct(i).CaptRct(j).Right = ScaleWidth - 2
        OffsetRect ChildsRct(i).Rct(j), 0, -(NewPos - LastPos)
        OffsetRect ChildsRct(i).CaptRct(j), 0, -(NewPos - LastPos)
      End If
    Next
  Next
End Sub

Private Sub SetScroll()
  If UserControl.ScaleHeight < TheHeight Then
    WS.Scrollbars = sbVertical
    WS.SetRange sbVertical, 0, ((TheHeight - UserControl.ScaleHeight)) + (UserControl.ScaleHeight) - 1
    WS.VPage = UserControl.ScaleHeight
    OffsetAll 0, WS.VPosition
  Else
    WS.VPosition = 0
    WS.Scrollbars = sbNone
    SetRects
  End If
End Sub
