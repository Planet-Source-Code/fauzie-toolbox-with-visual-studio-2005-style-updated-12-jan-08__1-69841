VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Child"
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
Public Value As Boolean
Public Key$, Description$
Public UseMaskColor As Boolean
Public MaskColor As OLE_COLOR

Private mvarIcon As StdPicture 'local copy

Public Property Set Icon(ByVal vData As StdPicture)
  Set mvarIcon = vData
End Property

Public Property Get Icon() As StdPicture
  Set Icon = mvarIcon
End Property

