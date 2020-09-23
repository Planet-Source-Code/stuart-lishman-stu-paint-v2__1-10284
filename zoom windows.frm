VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Zoom Window x1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3210
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider Slider1 
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5741
      _Version        =   327682
      MousePointer    =   1
      Orientation     =   1
      Min             =   1
      Max             =   15
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      X1              =   840
      X2              =   2400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      X1              =   1200
      X2              =   1200
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Image ZoomStretch 
      Height          =   1095
      Left            =   600
      MouseIcon       =   "zoom windows.frx":0000
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
zoomactive = True
Call StayOnTop
Me.left = 0
Me.top = (Screen.Height - Me.Height) - Form1.StatusBar1.Height
Form2.ZoomStretch.Width = ScaleX(200, vbPixels, vbTwips)
Form2.ZoomStretch.Height = ScaleY(200, vbPixels, vbTwips)
Form2.Line1.X1 = Form2.ZoomStretch.left + (Form2.ZoomStretch.Width / 2)
Form2.Line1.X2 = Form2.ZoomStretch.left + (Form2.ZoomStretch.Width / 2)
Form2.Line1.Y1 = Form2.ZoomStretch.top
Form2.Line1.Y2 = Form2.ZoomStretch.top + Form2.ZoomStretch.Height
Form2.Line2.Y1 = Form2.ZoomStretch.top + (Form2.ZoomStretch.Height / 2)
Form2.Line2.Y2 = Form2.ZoomStretch.top + (Form2.ZoomStretch.Height / 2)
Form2.Line2.X1 = Form2.ZoomStretch.left
Form2.Line2.X2 = Form2.ZoomStretch.left + Form2.ZoomStretch.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
zoomactive = False
Unload Me
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form2.hWnd, -1, 0, 0, 0, 0, &H1)
End Function
Private Sub Slider1_Change()
Form2.Caption = "Zoom Window x" & Slider1.Value
End Sub
Private Sub Slider1_Click()
Form2.Caption = "Zoom Window x" & Slider1.Value
End Sub
