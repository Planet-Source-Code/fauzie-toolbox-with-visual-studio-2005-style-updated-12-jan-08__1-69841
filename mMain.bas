Attribute VB_Name = "mMain"
Option Explicit
DefInt A-Z

'Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
'Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4

Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, _
                pRect As RECT, pClipRect As RECT) As Long
Public Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, _
                pRect As RECT, pRegion As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public PE As cPaintEffects

Public Const ASMAIL$ = "fauzie811@yahoo.com"

Public Const INTERR$ = "An unexpected application error has occured!"
Public Const ERRTEXT$ = "If this problem continues, please contact me, at " + ASMAIL$ + ", quoting the above information."

Public Sub InitPaintEffects()
  If PE Is Nothing Then
    Set PE = New cPaintEffects
  End If
End Sub

Public Sub Highlight(c As Control)
  With c
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Public Function IsInControl(ByVal hWnd As Long) As Boolean
  Dim P As POINTAPI
  GetCursorPos P
  If hWnd = WindowFromPoint(P.x, P.y) Then IsInControl = -1
End Function

Public Function DrawTheme(ByVal ihDC As Long, _
                           sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           bRect As RECT) As Boolean
  
  Dim hTheme As Long
  Dim lResult As Long
  Dim hRgn As Long
  On Error GoTo NoXP

  hTheme = OpenThemeData(0&, StrPtr(sClass))

  If hTheme Then
    lResult = DrawThemeBackground(hTheme, ihDC, iPart, iState, bRect, bRect)
    DrawTheme = IIf(lResult, False, True)
  Else
    DrawTheme = False
  End If

  Exit Function

NoXP:
  DrawTheme = False

End Function


