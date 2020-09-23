VERSION 5.00
Begin VB.UserControl bsBuildBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
End
Attribute VB_Name = "bsBuildBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : bsBuildBox (User Control)
' DateTime  : 26/08/2000
' Author    : Drew (aka The Bad One)
' Purpose   : Have you ever seen a text box that has a button stuck inside it with "..."
'             written on it? This is that kind of control, custom-drawn to be
'             customisable.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Updates
'---------------------------------------------------------------------------------------
' 01/11/2003   The code has also been cleaned up, with many "magic numbers" being
'              removed, and most of the drawing commands replaced by API calls.
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_PUSHED = &H200
Private Const DFCS_FLAT = &H4000
Private Const PS_SOLID = 0
Private Const CLR_INVALID = &HFFFF
Private Const UNIT_TWIPS = 120

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_MIDDLE = &H800
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000

Private Const COLOR_BTNTEXT = 18
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_WINDOW = 5

Private Const DT_LEFT = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000

'Default property values for the control:
Const m_def_TextColour = 0
Const m_def_BorderStyle = 1
Const m_def_ThinBorderColour = &H999999
Const m_def_ButtonStyle = 2
Const m_def_Data = ""

'Property variables for the control:
Dim m_Text As String
Dim m_BackColour As OLE_COLOR
Dim m_Font As Font
Dim m_TextColour As OLE_COLOR
Dim m_BorderStyle As Integer
Dim m_ThinBorderColour As OLE_COLOR
Dim m_ButtonStyle As Integer
Dim m_Data As String

'Control's events!
Event Click(ByVal MouseButton As MouseButtonConstants)
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event NewData()

'Finally some variables as used by the control.
Private rctButton As RECT, txtRect As RECT, focusRect As RECT
Private btnDown As Boolean, btnState As Boolean
Private lBorderWidth As Long

' **** ENUMERATIONS ****
Enum bbEdgeStyle
   esNoBorder
   esThinBorder
   esRaisedThin
   esSunkenThin
   esRaised3D
   esSunken3D
   esEtched
   esBump
End Enum

Enum bbButtonStyle
   tsFlat
   tsThin
   tsNormal
End Enum

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_Click
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Sets the focus on the control.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_Click()
   UserControl.SetFocus
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_DblClick
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Double-click event.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_EnterFocus
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Redraws the control when it receives the focus.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_EnterFocus()
   
   Dim hBrush As Long
   
   'Drawing the highlight rectangle:
   hBrush = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
   SetRect txtRect, lBorderWidth + 2, lBorderWidth + 1, ScaleWidth - 18, ScaleHeight - (lBorderWidth + 1)
   FillRect hdc, txtRect, hBrush
   DeleteObject hBrush
   'Rectangle hdc, lBorderWidth + 1, lBorderWidth + 1, ScaleWidth - 18, ScaleHeight - (lBorderWidth + 1)
       
   'Adjust the focus rectangle so that it doesn't ruin the text.
   SetRect focusRect, lBorderWidth + 1, lBorderWidth + 1, ScaleWidth - 18, ScaleHeight - lBorderWidth - 1
   
   'Draw a focus rectangle so that we know the control has the focus.
   DrawFocusRect hdc, focusRect
   
   'It's important that this next line is where it is otherwise the focus rectangle will
   'appear to be a black border. Don't ask me why.
   SetTextColor hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
   
   'Now draw the text!
   DrawText UserControl.hdc, Text, Len(Text), txtRect, DT_LEFT + DT_VCENTER + DT_SINGLELINE Or DT_WORD_ELLIPSIS
       
   'Refresh the control or the changes will not be seen.
   UserControl.Refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_ExitFocus
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : To clean the control to get rid of the focus rectangle.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_ExitFocus()
   DrawControl
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_KeyPress
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Raises the Click event, as if the user had used the left mouse button.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent Click(vbLeftButton)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_MouseUp
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Raises the Click event if the button has been clicked.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If WithinButton(X, Y) Then
      DrawButton False
      Refresh
      RaiseEvent Click(Button)
   End If
   btnDown = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_MouseMove
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Responsible for drawing the button as pressed when necessary.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If btnDown Then
      btnState = WithinButton(X, Y)
      DrawButton btnState
      Refresh
   End If
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_MouseDown
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Monitors the mouse to see if it has been pressed on the button.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   btnDown = WithinButton(X, Y)
   If btnDown Then
      DrawButton True
      Refresh
   End If
End Sub

Private Sub UserControl_Paint()
   'DrawControl
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.UserControl_Resize
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Repaint the control if it has been resized.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_Resize()
   DrawControl
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.DrawControl
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Responsible for the rendering of the control; it also calls DrawButton().
' Assuming  : UserControl's ScaleMode is vbPixels.
'---------------------------------------------------------------------------------------
'
Private Sub DrawControl()
    
   Dim rctBorder As RECT
   Dim hBrush As Long, hPen As Long
   
   'Set the border to the size of the control (minus button).
   Call SetRect(rctBorder, 0, 0, ScaleWidth - (UNIT_TWIPS * 2) / Screen.TwipsPerPixelX, _
      UserControl.ScaleHeight)
   hBrush = CreateSolidBrush(TranslateColour(m_BackColour))
   FillRect hdc, rctBorder, hBrush
   DeleteObject hBrush

   Select Case BorderStyle
      Case esThinBorder
         'Thin border of one colour
         hPen = CreatePen(PS_SOLID, 1, TranslateColour(m_ThinBorderColour))
         DeleteObject SelectObject(hdc, hPen)
         DrawRect UserControl.hdc, rctBorder
         DeleteObject hPen
         lBorderWidth = 1
   
      Case esRaisedThin
         'A raised thin border
         DrawEdge UserControl.hdc, rctBorder, BDR_RAISEDINNER, BF_RECT
         lBorderWidth = 1

      Case esSunkenThin
         'A sunken thin border
         DrawEdge UserControl.hdc, rctBorder, BDR_SUNKENOUTER, BF_RECT
         lBorderWidth = 1

      Case esSunken3D
         'A sunken border
         DrawEdge UserControl.hdc, rctBorder, EDGE_SUNKEN, BF_RECT
         lBorderWidth = 2

      Case esRaised3D
         'A raised border
         DrawEdge UserControl.hdc, rctBorder, EDGE_RAISED, BF_RECT
         lBorderWidth = 2

      Case esEtched
         'An etched border
         DrawEdge UserControl.hdc, rctBorder, EDGE_ETCHED, BF_RECT
         lBorderWidth = 2

      Case esBump
         'A 'chiselled' border
         DrawEdge UserControl.hdc, rctBorder, EDGE_BUMP, BF_RECT
         lBorderWidth = 2
   End Select

   'For the text all we need to do really is draw the text.
   Set UserControl.Font = m_Font
   SetTextColor hdc, TranslateColour(m_TextColour)
   SetRect txtRect, lBorderWidth + 2, lBorderWidth + 1, ScaleWidth - 20, ScaleHeight - lBorderWidth
   DrawText UserControl.hdc, Text, Len(Text), txtRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS

   'Adjust the focus rectangle so that it doesn't ruin the text (or the border).
   SetRect focusRect, lBorderWidth + 2, lBorderWidth, ScaleWidth - 16 - lBorderWidth, ScaleHeight - lBorderWidth

   DrawButton False
   Refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TranslateColour
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Used to convert Automation colours to a Windows (long) colour.
' Assuming  : Colour passed is unsigned (ie. &Hxxxxxx&)
'---------------------------------------------------------------------------------------
'
Private Function TranslateColour(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If TranslateColor(oClr, hPal, TranslateColour) Then
      TranslateColour = CLR_INVALID
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.DrawRect
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Draws a rectangle from a RECT type.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub DrawRect(ByVal hdc As Long, Coords As RECT)
   Rectangle hdc, Coords.Left, Coords.Top, Coords.Right, Coords.Bottom
End Sub

Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "The colour of the text area of the control."
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
   m_BackColour = New_BackColour
   PropertyChanged "BackColour"
   DrawControl
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets the font used for the text in the BuildBox."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
   Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set m_Font = New_Font
   PropertyChanged "Font"
   DrawControl
End Property

Public Property Get TextColour() As OLE_COLOR
Attribute TextColour.VB_Description = "The colour of the text."
Attribute TextColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   TextColour = m_TextColour
End Property

Public Property Let TextColour(ByVal New_TextColour As OLE_COLOR)
   m_TextColour = New_TextColour
   PropertyChanged "TextColour"
   DrawControl
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_BackColour = GetSysColor(COLOR_WINDOW)
   Set m_Font = Ambient.Font
   m_TextColour = m_def_TextColour
   m_Text = UserControl.Extender.Name
   m_BorderStyle = m_def_BorderStyle
   m_ThinBorderColour = m_def_ThinBorderColour
   m_ButtonStyle = m_def_ButtonStyle
   m_Data = m_def_Data
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_BackColour = PropBag.ReadProperty("BackColour", GetSysColor(COLOR_WINDOW))
   Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_TextColour = PropBag.ReadProperty("TextColour", m_def_TextColour)
   m_Text = PropBag.ReadProperty("Text", UserControl.Extender.Name)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_ThinBorderColour = PropBag.ReadProperty("ThinBorderColour", m_def_ThinBorderColour)
   m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_Data = PropBag.ReadProperty("Data", m_def_Data)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("BackColour", m_BackColour, GetSysColor(COLOR_WINDOW))
   Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
   Call PropBag.WriteProperty("TextColour", m_TextColour, m_def_TextColour)
   Call PropBag.WriteProperty("Text", m_Text, UserControl.Extender.Name)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("ThinBorderColour", m_ThinBorderColour, m_def_ThinBorderColour)
   Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Data", m_Data, m_def_Data)
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Text held inside the BuildBox."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "200"
   Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
   m_Text = New_Text
   PropertyChanged "Text"
   DrawControl
End Property

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.DrawButton
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : This only draws the button on the control, in response to the user.
' Assuming  : UserControl's ScaleMode is vbPixels.
'---------------------------------------------------------------------------------------
'
Private Sub DrawButton(bIsDown As Boolean)
    
   Dim I As Integer
   Dim rctBuild As RECT
   Dim hPen As Long
   
   Const BUILD_SPACING = 4
   Const BUILD_WIDTH = 3
    
   'Define the button as a rectangle.
   SetRect rctButton, ScaleWidth - (UNIT_TWIPS * 2) / Screen.TwipsPerPixelX, _
      0, ScaleWidth, ScaleHeight
    
   'Find out the size of the selected border...
   Select Case m_BorderStyle
      Case esNoBorder:                       lBorderWidth = 0
      Case esThinBorder To esSunkenThin:     lBorderWidth = 1
      Case Else:                             lBorderWidth = SystemMetric(smDialogBorderWidth)
   End Select
    
   'Now to draw the button! How the button is drawn depends on the ButtonStyle.
   With rctButton
      If bIsDown Then
         DrawFrameControl hdc, rctButton, DFC_BUTTON, DFCS_BUTTONPUSH + DFCS_PUSHED
      Else
      Select Case m_ButtonStyle
         Case tsNormal
            DrawFrameControl hdc, rctButton, DFC_BUTTON, DFCS_BUTTONPUSH
         Case tsThin
            DrawEdge UserControl.hdc, rctButton, BDR_RAISEDINNER, BF_RECT + BF_SOFT + BF_MIDDLE
         Case tsFlat
            DrawFrameControl hdc, rctButton, DFC_BUTTON, DFCS_BUTTONPUSH + DFCS_FLAT
         End Select
      End If
   End With
    
   'Drawing the build icon.
   If Enabled Then
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNTEXT))
   Else
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_GRAYTEXT))
   End If
   DeleteObject SelectObject(hdc, hPen)
   For I = 0 To 2
      With rctBuild
         .Left = rctButton.Left + (BUILD_SPACING + BUILD_WIDTH * I + Abs(bIsDown))
         .Top = rctButton.Bottom - (6 + Abs(bIsDown))
         .Right = .Left + 2
         .Bottom = .Top + 2
         Rectangle hdc, .Left, .Top, .Right, .Bottom
      End With
   Next
   DeleteObject hPen
   'refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.WithinButton
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Finds out whether or not a point is inside the area of a Rect.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Function WithinButton(ByVal X As Long, ByVal Y As Long) As Boolean
   With rctButton
      WithinButton = isInLimit(X, .Left, .Right) And isInLimit(Y, .Top, .Bottom)
   End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : bsBuildBox.isInLimit
' DateTime  : 01/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Is true if value is between low and high, inclusive.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Function isInLimit(Value, Low, High) As Boolean
   isInLimit = Not (Value < Low Or Value > High)
End Function

Public Property Get BorderStyle() As bbEdgeStyle
Attribute BorderStyle.VB_Description = "The style of the edges of the control."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bbEdgeStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawControl
End Property

Public Property Get ThinBorderColour() As OLE_COLOR
Attribute ThinBorderColour.VB_Description = "If BorderStyle is set to esThinBorder, this is the border's colour."
Attribute ThinBorderColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   ThinBorderColour = m_ThinBorderColour
End Property

Public Property Let ThinBorderColour(ByVal New_ThinBorderColour As OLE_COLOR)
   m_ThinBorderColour = New_ThinBorderColour
   PropertyChanged "ThinBorderColour"
   DrawControl
End Property

Public Property Get ButtonStyle() As bbButtonStyle
Attribute ButtonStyle.VB_Description = "They style of the Build button."
Attribute ButtonStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As bbButtonStyle)
   m_ButtonStyle = New_ButtonStyle
   PropertyChanged "ButtonStyle"
   DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "The enabled state of the BuildBox."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Data() As String
Attribute Data.VB_Description = "A hidden value, like a combo box's ItemData property."
   Data = m_Data
End Property

Public Property Let Data(ByVal New_Data As String)
   m_Data = New_Data
   PropertyChanged "Data"
   RaiseEvent NewData
End Property
