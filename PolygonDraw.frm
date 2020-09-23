VERSION 5.00
Begin VB.Form PolygonDraw 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Draw Polygon"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5025
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Text            =   "5"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3060
      Left            =   120
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Radius:"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Draw at Angle:"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Sides:"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PolygonDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 NumSides = Val(Text1.Text)
 AtAngle = Val(Text2.Text)
 Call Draw_Polygon
 Unload Me
End Sub

Private Sub Command2_Click()
Form1.LineDr.Visible = False
Unload Me
End Sub

Private Sub Command3_Click()
Dim points(0 To 100) As POINTAPI
Dim retval As Long
an = (2 * Pi) / Text1.Text
d = an
anp = (Text2.Text / 360) * (2 * Pi)
an = an + anp
For a = 0 To Text1.Text
points(a).X = 100 + 90 * Cos(an)
points(a).Y = 100 + 90 * Sin(an)
an = an + d
Next a
Picture1.Cls
retval = Polygon(Picture1.hdc, points(0), Text1.Text)
End Sub

Private Sub Form_Load()
If Form1.LineDr.X1 > Form1.LineDr.X2 Then lentx = Form1.LineDr.X1 - Form1.LineDr.X2
If Form1.LineDr.X2 > Form1.LineDr.X1 Then lentx = Form1.LineDr.X2 - Form1.LineDr.X1
If Form1.LineDr.Y1 > Form1.LineDr.Y2 Then lenty = Form1.LineDr.Y1 - Form1.LineDr.Y2
If Form1.LineDr.Y2 > Form1.LineDr.Y1 Then lenty = Form1.LineDr.Y2 - Form1.LineDr.Y1
Label4.Caption = Sqr((lentx * lentx) + (lenty * lenty))
End Sub
