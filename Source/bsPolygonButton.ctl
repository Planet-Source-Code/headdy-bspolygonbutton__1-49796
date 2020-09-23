VERSION 5.00
Begin VB.UserControl bsPolygonButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "bsPolygonButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : bsPolygonButton (User Control)
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : To provide a button control that takes the shape of a polygon
'             of almost any number of sides.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Updates
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type COORD
   X As Long
   Y As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Const WINDING = 2
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_NOCLIP = &H100
Private Const DT_VCENTER = &H4

Private m_iSides As Integer

Const m_def_iSides = 6
'Default Property Values:
Const m_def_ShowFocus = True
Const m_def_CaptionColour = vbButtonText
Const m_def_ButtonColour = vbButtonFace
Const m_def_LightestColour = vb3DHighlight
Const m_def_LightColour = vb3DLight
Const m_def_DarkColour = vb3DShadow
Const m_def_DarkestColour = vb3DDKShadow
Const m_def_iRotation = 0

'Property Variables:
Dim m_ShowFocus As Boolean
Dim m_CaptionColour As OLE_COLOR
Dim m_ButtonColour As OLE_COLOR
Dim m_Fount As Font
Dim m_LightestColour As OLE_COLOR
Dim m_LightColour As OLE_COLOR
Dim m_DarkColour As OLE_COLOR
Dim m_DarkestColour As OLE_COLOR
Dim m_Caption As String
Dim m_iRotation As Integer

'Event Declarations:
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event Click()
Event DblClick()

Const Pi# = 3.1415927
Const CLR_INVALID = &HFFFF

Dim hRegion As Long
Dim booGotFocus As Boolean


'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.Sides
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Gets/sets the number of sides the button has.
' Assuming  : Number of sides is between 3 and 100, inclusive.
'---------------------------------------------------------------------------------------
'
Public Property Get Sides() As Integer
Attribute Sides.VB_ProcData.VB_Invoke_Property = ";Appearance"
   Sides = m_iSides
End Property

Public Property Let Sides(ByVal iSides As Integer)
   If m_iSides < 3 Then
      m_iSides = 3
   ElseIf m_iSides > 100 Then
      m_iSides = 100
   End If
   m_iSides = iSides
   Call UserControl.PropertyChanged("Sides")
   DrawControl
End Property

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.DrawControl
' DateTime  : 09/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Draws the whole control (pressed if necessary).
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub DrawControl(Optional booPressed As Boolean)
   Dim X(0 To 1) As Single, Y(0 To 1) As Single
   Dim rctControl As RECT, lpOld As POINTAPI
   Dim I As Integer, iCounter As Integer
   Dim hBrush As Long
   
   Dim PolyCoord(100) As COORD
   
   SetWindowRgn UserControl.hWnd, 0, True
   ScaleMode = vbPixels
   AutoRedraw = True
   
   ' Clear the control (button colour)
   ' -------------------------------------------------------------------
   SetRect rctControl, 0, 0, ScaleWidth, ScaleHeight
   hBrush = CreateSolidBrush(TranslateColour(m_ButtonColour))
   FillRect UserControl.hdc, rctControl, hBrush
   DeleteObject hBrush
   
   ' Remember, X coordinate = Sin(angle) * X radius + X centre, and
   '           Y coordinate = Cos(angle) * Y radius + Y centre
   
   
   ' Draw text
   ' -------------------------------------------------------------------
   Set Font = m_Fount
   DrawText hdc, m_Caption, Len(m_Caption), rctControl, DT_CALCRECT
   If UserControl.Enabled Then
      ForeColor = TranslateColour(m_CaptionColour)
      OffsetRect rctControl, ScaleWidth / 2 - rctControl.Right / 2, _
         ScaleHeight / 2 - rctControl.Bottom / 2
      If booPressed Then
         OffsetRect rctControl, 1, 1
      End If
      DrawText hdc, m_Caption, Len(m_Caption), rctControl, DT_CENTER + DT_VCENTER + DT_NOCLIP
   Else
      ForeColor = TranslateColour(m_LightColour)
      OffsetRect rctControl, ScaleWidth / 2 - rctControl.Right / 2 + 1, _
         ScaleHeight / 2 - rctControl.Bottom / 2 + 1
      DrawText hdc, m_Caption, Len(m_Caption), rctControl, DT_CENTER + DT_VCENTER + DT_NOCLIP
      ForeColor = TranslateColour(m_DarkColour)
      OffsetRect rctControl, -1, -1
      DrawText hdc, m_Caption, Len(m_Caption), rctControl, DT_CENTER + DT_VCENTER + DT_NOCLIP
   End If
   
   ' Draw focus rectangle
   ' -------------------------------------------------------------------
   If booGotFocus And m_ShowFocus Then
      DrawFocusRect hdc, rctControl
   End If
   
   ' Draw the edges
   ' -------------------------------------------------------------------
   For I = 0 To 360 Step 360 / m_iSides
      X(0) = Sin(DegreesToRadians(I + m_iRotation)) * ((ScaleWidth - 1) / 2) + ((ScaleWidth - 1) / 2)
      Y(0) = Cos(DegreesToRadians(I + m_iRotation)) * ((ScaleHeight - 1) / 2) + ((ScaleHeight - 1) / 2)
      X(1) = Sin(DegreesToRadians(I + m_iRotation + 360 / m_iSides)) * ((ScaleWidth - 1) / 2) + ((ScaleWidth - 1) / 2)
      Y(1) = Cos(DegreesToRadians(I + m_iRotation + 360 / m_iSides)) * ((ScaleHeight - 1) / 2) + ((ScaleHeight - 1) / 2)

      ' first line
      DrawWidth = 2
      If booPressed Then
         ForeColor = TranslateColour(m_DarkestColour)
      Else
         If (ScaleHeight - (X(1) / ScaleWidth) * ScaleHeight <= Y(1)) Then
            ForeColor = TranslateColour(m_DarkColour)
         Else
            If ScaleHeight - (X(0) / ScaleWidth) * ScaleHeight <= Y(0) Then
               ForeColor = TranslateColour(m_DarkColour)
            Else
               ForeColor = TranslateColour(m_LightestColour)
            End If
         End If
      End If
      MoveToEx hdc, X(0), Y(0), lpOld
      LineTo hdc, X(1), Y(1)
      
      ' second line
      DrawWidth = 1
      If booPressed Then
         ForeColor = TranslateColour(m_DarkColour)
      Else
         If (ScaleHeight - (X(1) / ScaleWidth) * ScaleHeight <= Y(1)) Then
            ForeColor = TranslateColour(m_DarkestColour)
         Else
            If ScaleHeight - (X(0) / ScaleWidth) * ScaleHeight <= Y(0) Then
               ForeColor = TranslateColour(m_DarkestColour)
            Else
               ForeColor = TranslateColour(m_LightColour)
            End If
         End If
      End If
      MoveToEx hdc, X(0) + 1, Y(0) + 1, lpOld
      LineTo hdc, X(1) + 1, Y(1) + 1
   Next

   ' Create polygon region
   ' -------------------------------------------------------------------
   For I = 0 To 360 Step 360 / m_iSides
      PolyCoord(iCounter).X = Sin(DegreesToRadians(I + m_iRotation)) * ((ScaleWidth + 1) / 2) + ((ScaleWidth + 1) / 2)
      PolyCoord(iCounter).Y = Cos(DegreesToRadians(I + m_iRotation)) * ((ScaleHeight + 1) / 2) + ((ScaleHeight + 1) / 2)
      iCounter = iCounter + 1
   Next
   hRegion = CreatePolygonRgn(PolyCoord(0), m_iSides, WINDING)
   SetWindowRgn UserControl.hWnd, hRegion, True
   
   ' Because we've set AutoRedraw to True...
   Refresh
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_ExitFocus()
   booGotFocus = False
   DrawControl
End Sub

Private Sub UserControl_GotFocus()
   booGotFocus = True
   DrawControl
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      DrawControl True '(PtInRegion(hRegion, X, Y) <> 0)
   End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      DrawControl True
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      DrawControl
   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.UserControl_ReadProperties
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Reads the stored values for the properties.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_iSides = PropBag.ReadProperty("Sides", m_def_iSides)
   m_iRotation = PropBag.ReadProperty("Rotation", m_def_iRotation)
   m_LightestColour = PropBag.ReadProperty("LightestColour", m_def_LightestColour)
   m_LightColour = PropBag.ReadProperty("LightColour", m_def_LightColour)
   m_DarkColour = PropBag.ReadProperty("DarkColour", m_def_DarkColour)
   m_DarkestColour = PropBag.ReadProperty("DarkestColour", m_def_DarkestColour)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
   m_ButtonColour = PropBag.ReadProperty("ButtonColour", m_def_ButtonColour)
   Set m_Fount = PropBag.ReadProperty("Fount", Ambient.Font)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_ShowFocus = PropBag.ReadProperty("ShowFocus", m_def_ShowFocus)
End Sub

Private Sub UserControl_Resize()
   DrawControl
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.Rotation
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Allows the user to specify by how much the polygon is
'             "rotated".
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Public Property Get Rotation() As Integer
Attribute Rotation.VB_Description = "Specifies the rotation of the polygon."
Attribute Rotation.VB_ProcData.VB_Invoke_Property = ";Appearance"
   Rotation = m_iRotation
End Property

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.Rotation
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Allows the user to specify by how much the polygon is
'             "rotated".
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Public Property Let Rotation(ByVal New_Rotation As Integer)
   New_Rotation = New_Rotation Mod 360
   If New_Rotation < 0 Then
      New_Rotation = 360 - New_Rotation
   End If
   m_iRotation = New_Rotation
   PropertyChanged "Rotation"
   DrawControl
End Property

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.UserControl_InitProperties
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Sets the default values for the properties.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_InitProperties()
   m_iRotation = m_def_iRotation
   m_iSides = m_def_iSides
   m_LightestColour = m_def_LightestColour
   m_LightColour = m_def_LightColour
   m_DarkColour = m_def_DarkColour
   m_DarkestColour = m_def_DarkestColour
   m_Caption = Extender.Name
   m_ButtonColour = m_def_ButtonColour
   Set m_Fount = Ambient.Font
   m_CaptionColour = m_def_CaptionColour
   m_ShowFocus = m_def_ShowFocus
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.UserControl_Terminate
' DateTime  : 09/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Removes the region from memory, before the control is destroyed.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_Terminate()
   DeleteObject hRegion
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.UserControl_WriteProperties
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : "Saves" the properties for later use.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Sides", m_iSides, m_def_iSides)
   Call PropBag.WriteProperty("Rotation", m_iRotation, m_def_iRotation)
   Call PropBag.WriteProperty("LightestColour", m_LightestColour, m_def_LightestColour)
   Call PropBag.WriteProperty("LightColour", m_LightColour, m_def_LightColour)
   Call PropBag.WriteProperty("DarkColour", m_DarkColour, m_def_DarkColour)
   Call PropBag.WriteProperty("DarkestColour", m_DarkestColour, m_def_DarkestColour)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
   Call PropBag.WriteProperty("ButtonColour", m_ButtonColour, m_def_ButtonColour)
   Call PropBag.WriteProperty("Fount", m_Fount, Ambient.Font)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("ShowFocus", m_ShowFocus, m_def_ShowFocus)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LightestColour() As OLE_COLOR
Attribute LightestColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   LightestColour = m_LightestColour
End Property

Public Property Let LightestColour(ByVal New_LightestColour As OLE_COLOR)
   m_LightestColour = New_LightestColour
   PropertyChanged "LightestColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LightColour() As OLE_COLOR
Attribute LightColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   LightColour = m_LightColour
End Property

Public Property Let LightColour(ByVal New_LightColour As OLE_COLOR)
   m_LightColour = New_LightColour
   PropertyChanged "LightColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DarkColour() As OLE_COLOR
Attribute DarkColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   DarkColour = m_DarkColour
End Property

Public Property Let DarkColour(ByVal New_DarkColour As OLE_COLOR)
   m_DarkColour = New_DarkColour
   PropertyChanged "DarkColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DarkestColour() As OLE_COLOR
Attribute DarkestColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   DarkestColour = m_DarkestColour
End Property

Public Property Let DarkestColour(ByVal New_DarkestColour As OLE_COLOR)
   m_DarkestColour = New_DarkestColour
   PropertyChanged "DarkestColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,usercontrol.extender.name
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbbuttonface
Public Property Get ButtonColour() As OLE_COLOR
Attribute ButtonColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   ButtonColour = m_ButtonColour
End Property

Public Property Let ButtonColour(ByVal New_ButtonColour As OLE_COLOR)
   m_ButtonColour = New_ButtonColour
   PropertyChanged "ButtonColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Fount() As Font
Attribute Fount.VB_ProcData.VB_Invoke_Property = ";Font"
   Set Fount = m_Fount
End Property

Public Property Set Fount(ByVal New_Fount As Font)
   Set m_Fount = New_Fount
   PropertyChanged "Fount"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbbuttontext
Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowFocus() As Boolean
Attribute ShowFocus.VB_Description = "Determines whether or not the focus rectangle is shown."
Attribute ShowFocus.VB_ProcData.VB_Invoke_Property = ";Behavior"
   ShowFocus = m_ShowFocus
End Property

Public Property Let ShowFocus(ByVal New_ShowFocus As Boolean)
   m_ShowFocus = New_ShowFocus
   PropertyChanged "ShowFocus"
End Property


'---------------------------------------------------------------------------------------
' Procedure : bsPolygonButton.ShowAbout
' DateTime  : 09/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Shows the About screen.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Public Sub ShowAbout()
   frmAbout.Show vbModal
End Sub

'---------------------------------------------------------------------------------------
' Procedure : modUseful.DegreesToRadians
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Converts a value in degrees to radians, as used by Visual Basic.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function DegreesToRadians(ByVal sngAngle As Single) As Single
   DegreesToRadians = sngAngle * (Pi / 180)
End Function
'---------------------------------------------------------------------------------------
' Procedure : TranslateColour
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Used to convert Automation colours to a Windows (long) colour.
'---------------------------------------------------------------------------------------
'
Function TranslateColour(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If TranslateColor(oClr, hPal, TranslateColour) Then
       TranslateColour = CLR_INVALID
   End If
End Function
