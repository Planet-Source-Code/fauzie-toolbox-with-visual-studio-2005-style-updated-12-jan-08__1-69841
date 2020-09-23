VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Option Explicit
DefInt A-Z

Public Caption$, Tag$, ToolTipText$
'Public Enabled As Boolean, Visible As Boolean
Public Expanded As Boolean
Public ClientLeft As Single, ClientHeight As Single
Public ClientWidth As Single, ClientTop As Single
Public Key$, Description$
Public SelectedChild As Integer

Public ChildCount As Integer
'local variable(s) to hold property value(s)
Private mvarChilds() As Child 'local copy

Public Property Set Childs(ByVal Index As Variant, ByVal vData As Child)
  Index = KeyToIndex(Index)
  Set mvarChilds(Index) = vData
End Property

Public Property Get Childs(ByVal Index As Variant) As Child
  Index = KeyToIndex(Index)
  Set Childs = mvarChilds(Index)
End Property

Private Function KeyToIndex(ByVal Index As Variant) As Integer
 ' Returns the integer index value of a button identified by either it's
 ' Key property or index.
 Dim I
 If Val(Index) = 0 Then
  For I = 1 To ChildCount
   If UCase$(mvarChilds(I).Key) = UCase$(Index) Then
    KeyToIndex = I
    Exit Function
   End If
  Next
 Else
  KeyToIndex = Val(Index)
 End If
 If KeyToIndex = 0 And Index <> 0 Then
  Err.Raise 35601, "Nodes.KeyToIndex", "Element not found. Key is missing or illegal."
 End If
End Function

Public Function AddChild(cCaption As String, Optional cKey As String, Optional cIcon As StdPicture = Nothing, Optional cToolTip As String, Optional cDescription As String, Optional cMaskColor As OLE_COLOR = vbMagenta, Optional cUseMaskColor As Boolean = True) As Integer
  ChildCount = ChildCount + 1
  ReDim Preserve mvarChilds(ChildCount) As Child
  Set mvarChilds(ChildCount) = New Child
  With mvarChilds(ChildCount)
    .Key = cKey
    Set .Icon = cIcon
    .ToolTipText = cToolTip
    .MaskColor = cMaskColor
    .UseMaskColor = cUseMaskColor
    .Caption = cCaption
  End With
  AddChild = ChildCount
End Function
