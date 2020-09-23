VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\A..\..\INCOMP~1\BSPOLY~1\Source\prjPolygonButton.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bsPolygonButton demo"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin prjPolygonButton.bsPolygonButton bsPolygonButton 
      Height          =   2775
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgColour 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      TabIndex        =   11
      Top             =   1800
      Width           =   5295
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Highlight"
      End
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Light"
      End
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Shadow"
      End
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Dark Shadow"
      End
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   16
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Button"
      End
      Begin prjTest.bsBuildBox bsColour 
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   17
         Top             =   960
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Caption"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin MSComCtl2.UpDown updRotation 
         Height          =   285
         Left            =   4816
         TabIndex        =   8
         Top             =   795
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtRotation"
         BuddyDispid     =   196613
         OrigLeft        =   5040
         OrigTop         =   840
         OrigRight       =   5280
         OrigBottom      =   1095
         Max             =   359
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updSides 
         Height          =   285
         Left            =   2056
         TabIndex        =   5
         Top             =   795
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtSides"
         BuddyDispid     =   196614
         OrigLeft        =   2280
         OrigTop         =   840
         OrigRight       =   2520
         OrigBottom      =   1095
         Max             =   100
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkShowFocus 
         Caption         =   "Show focus rectangle"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtRotation 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   795
         Width           =   495
      End
      Begin VB.TextBox txtSides 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   795
         Width           =   495
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rotation"
         Height          =   195
         Left            =   3000
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Number of sides"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About this control..."
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   3000
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000016&
      Height          =   3000
      Index           =   1
      Left            =   128
      Top             =   128
      Width           =   3000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Const CLR_INVALID = &HFFFF

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

Private Sub bsColour_Click(Index As Integer, ByVal MouseButton As MouseButtonConstants)
   
   On Error GoTo bsColour_Click_Error

   dlgColour.Color = TranslateColour(bsColour(Index).BackColour)
   dlgColour.ShowColor
   
   bsColour(Index).BackColour = dlgColour.Color
   
   With bsPolygonButton
      Select Case Index
         Case 0
            .LightestColour = dlgColour.Color
         Case 1
            .lightColour = dlgColour.Color
         Case 2
            .darkColour = dlgColour.Color
         Case 3
            .DarkestColour = dlgColour.Color
         Case 4
            .ButtonColour = dlgColour.Color
         Case 5
            .CaptionColour = dlgColour.Color
      End Select
   End With
   ModifyColourFores

   On Error GoTo 0
   Exit Sub

bsColour_Click_Error:

End Sub

Private Sub chkEnabled_Click()
   bsPolygonButton.Enabled = chkEnabled.Value
End Sub

Private Sub chkShowFocus_Click()
   bsPolygonButton.ShowFocus = chkShowFocus.Value
End Sub

Private Sub cmdAbout_Click()
   bsPolygonButton.ShowAbout
End Sub

Private Sub Form_Load()
   With bsPolygonButton
      txtCaption.Text = .Caption
      updSides.Value = .Sides
      updRotation.Value = .Rotation
      chkEnabled.Value = Abs(.Enabled = True)
      chkShowFocus.Value = Abs(.ShowFocus = True)
      
      bsColour(0).BackColour = TranslateColour(.LightestColour)
      bsColour(1).BackColour = TranslateColour(.lightColour)
      bsColour(2).BackColour = TranslateColour(.darkColour)
      bsColour(3).BackColour = TranslateColour(.DarkestColour)
      bsColour(4).BackColour = TranslateColour(.ButtonColour)
      bsColour(5).BackColour = TranslateColour(.CaptionColour)
   End With
   ModifyColourFores
End Sub

Private Sub txtCaption_Change()
   bsPolygonButton.Caption = txtCaption.Text
End Sub

Private Sub updRotation_Change()
   bsPolygonButton.Rotation = updRotation.Value
End Sub

Private Sub updSides_Change()
   bsPolygonButton.Sides = updSides.Value
End Sub

Private Sub ModifyColourFores()
   Dim I As Integer
   For I = 0 To 5
      bsColour(I).TextColour = IIf(Abs(bsColour(I).BackColour) > &H666666, vbBlack, vbWhite)
   Next
End Sub
