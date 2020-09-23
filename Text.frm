VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Text"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7860
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Raised"
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox ExamplePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   521
      TabIndex        =   20
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Frame Frame5 
      Caption         =   "3D Angle"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   7575
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6240
         TabIndex        =   19
         Text            =   "45"
         Top             =   240
         Width           =   855
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   135
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   238
         _Version        =   327682
         Min             =   1
         Max             =   30
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Angle"
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "3D effect"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Outlined"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Other Selections"
      Height          =   735
      Left            =   4800
      TabIndex        =   8
      Top             =   4200
      Width           =   2535
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   4
         Left            =   2040
         Picture         =   "Text.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   3
         Left            =   120
         Picture         =   "Text.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   2
         Left            =   1560
         Picture         =   "Text.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   1
         Left            =   1080
         Picture         =   "Text.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   0
         Left            =   600
         Picture         =   "Text.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Done"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type Text Here"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   7815
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Size"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Font"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Check2.Value = 0: Check3.Value = 0
Call draw_example
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then Frame5.Enabled = True: Check1.Value = 0: Check1.Enabled = False Else Frame5.Enabled = False: Check1.Enabled = True
Call draw_example
End Sub

Private Sub Check3_Click()
Check2.Value = 0: Check1.Value = 0
Call draw_example
End Sub

Private Sub Combo1_Click()
ExamplePic.FontName = Combo1
Form1.MainPic.FontName = ExamplePic.FontName
Form1.TextPic.FontName = ExamplePic.FontName
Form1.TextPic.Width = (Len(Text1.Text) * ExamplePic.FontSize)
Call draw_example
End Sub
Private Sub Combo2_Click()
ExamplePic.FontSize = Val(Combo2)
Form1.MainPic.FontSize = ExamplePic.FontSize
Form1.TextPic.FontSize = ExamplePic.FontSize
Form1.TextPic.Height = ExamplePic.FontSize + 30
Form1.TextPic.Width = (Len(Text1.Text) * ExamplePic.FontSize)
Call draw_example
End Sub
Private Sub Command1_Click()
Call Oldpic
Form1.TextPic.Caption = Text1.Text
xx = Form1.TextPic.left + 1
yy = Form1.TextPic.top + 1
If Check3.Value = 1 Then GoTo drwRaised
If Check2.Value = False Then GoTo ddd
chrs = 0
For a = 1 To Len(Text2.Text)
If Asc(Mid$(Text2.Text, a, 1)) < 48 Or Asc(Mid$(Text2.Text, a, 1)) > 57 Then chrs = 1
Next a
If chrs = 1 Then MsgBox "The angle text box should only contain numbers", vbCritical, "Error": Exit Sub
hh:
If Val(Text2.Text) > 360 Then Text2.Text = Val(Text2.Text) - 360: GoTo hh
VALU = Slider1.Value
ang = 360 / Val(Text2.Text)
ang = ((Pi * 2) / ang) + (Pi / 2)
For a = 1 To VALU
Form1.MainPic.ForeColor = rb
Form1.MainPic.DrawWidth = 1
Form1.MainPic.Line (xx + a * Cos(ang), yy + a * Sin(ang))-(xx + a * Cos(ang), yy + a * Sin(ang)), rb
Form1.MainPic.Print Text1.Text
Next a
GoTo ddff
drwRaised:
For X = xx - 1 To xx Step 1
For Y = yy - 1 To yy Step 1
Form1.MainPic.ForeColor = RGB(255, 255, 255)
Form1.MainPic.DrawWidth = 1: Form1.MainPic.Line (X, Y)-(X, Y), RGB(255, 255, 255): Form1.MainPic.Print Text1.Text
Next Y, X
For X = xx To xx + 1 Step 1
For Y = yy To yy + 1 Step 1
Form1.MainPic.ForeColor = RGB(0, 0, 0)
Form1.MainPic.DrawWidth = 1: Form1.MainPic.Line (X, Y)-(X, Y), RGB(0, 0, 0): Form1.MainPic.Print Text1.Text
Next Y, X
GoTo ddff
ddd:
If Check1.Value = False Then GoTo ddff
For X = xx - 1 To xx + 1 Step 1
For Y = yy - 1 To yy + 1 Step 1
Form1.MainPic.ForeColor = rb
Form1.MainPic.DrawWidth = 1: Form1.MainPic.Line (X, Y)-(X, Y), rb: Form1.MainPic.Print Text1.Text
Next Y, X
ddff:
Form1.MainPic.ForeColor = lb
Form1.MainPic.DrawWidth = 1: Form1.MainPic.Line (Form1.TextPic.left + 1, Form1.TextPic.top + 1)-(Form1.TextPic.left + 1, Form1.TextPic.top + 1), lb: Form1.MainPic.Print Text1.Text
Form1.MainPic.DrawWidth = Form1.Slider1.Value
Form1.TextPic.Caption = ""
Form1.TextPic.Visible = False
Unload Me
End Sub
Private Sub Command2_Click()
Form1.TextPic.Caption = ""
Form1.TextPic.Visible = False
Unload Me
End Sub
Private Sub Command3_Click(Index As Integer)
If Index = 3 Then ExamplePic.FontBold = False: ExamplePic.FontItalic = False: ExamplePic.FontUnderline = False: ExamplePic.FontStrikethru = False
If Index = 0 And ExamplePic.FontBold = True Then ExamplePic.FontBold = False: GoTo dnclk
If Index = 0 And ExamplePic.FontBold = False Then ExamplePic.FontBold = True: GoTo dnclk
If Index = 1 And ExamplePic.FontItalic = True Then ExamplePic.FontItalic = False: GoTo dnclk
If Index = 1 And ExamplePic.FontItalic = False Then ExamplePic.FontItalic = True: GoTo dnclk
If Index = 2 And ExamplePic.FontUnderline = True Then ExamplePic.FontUnderline = False: GoTo dnclk
If Index = 2 And ExamplePic.FontUnderline = False Then ExamplePic.FontUnderline = True: GoTo dnclk
If Index = 4 And ExamplePic.FontStrikethru = True Then ExamplePic.FontStrikethru = False: GoTo dnclk
If Index = 4 And ExamplePic.FontStrikethru = False Then ExamplePic.FontStrikethru = True: GoTo dnclk
dnclk:
If ExamplePic.FontBold = True Then Form1.MainPic.FontBold = True Else Form1.MainPic.FontBold = False
If ExamplePic.FontItalic = True Then Form1.MainPic.FontItalic = True Else Form1.MainPic.FontItalic = False
If ExamplePic.FontUnderline = True Then Form1.MainPic.FontUnderline = True Else Form1.MainPic.FontUnderline = False
If ExamplePic.FontStrikethru = True Then Form1.MainPic.FontStrikethru = True Else Form1.MainPic.FontStrikethru = False
Form1.TextPic.FontBold = Form1.MainPic.FontBold
Form1.TextPic.FontItalic = Form1.MainPic.FontItalic
Form1.TextPic.FontUnderline = Form1.MainPic.FontUnderline
Form1.TextPic.FontStrikethru = Form1.MainPic.FontStrikethru
Call draw_example
End Sub
Private Sub Form_Load()
Call StayOnTop
Me.left = Screen.Width - Me.Width
Me.top = (Screen.Height - Me.Height) - Form1.StatusBar1.Height
Form1.TextPic.ForeColor = lb
    Dim i As Integer
    With Combo1
        For i = 0 To Screen.FontCount - 1
            .AddItem Screen.Fonts(i)
        Next i
    End With
    With Combo2
        For i = 8 To 72 Step 2
            .AddItem i
        Next i
            .ListIndex = 1
    End With
For a = 0 To Combo1.ListCount
If Combo1.List(a) = Form1.MainPic.FontName Then Combo1.ListIndex = a
Next a
ExamplePic.FontName = Combo1
ExamplePic.FontSize = Val(Combo2)
End Sub
Private Sub Slider1_Change()
Call draw_example
End Sub
Private Sub Slider1_Scroll()
Call draw_example
End Sub
Private Sub Text1_Change()
Form1.TextPic.Caption = Text1.Text
Form1.TextPic.Width = (Len(Text1.Text) * ExamplePic.FontSize)
Form1.TextPic.ForeColor = lb
Call draw_example
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form3.hWnd, -1, 0, 0, 0, 0, &H1)
End Function
Private Sub draw_example()
ExamplePic.Cls
xx = 15
yy = 15
If Check3.Value = 1 Then GoTo raisedd
If Check2.Value = False Then GoTo ddd1
chrs = 0
For a = 1 To Len(Text2.Text)
If Asc(Mid$(Text2.Text, a, 1)) < 48 Or Asc(Mid$(Text2.Text, a, 1)) > 57 Then chrs = 1
Next a
If chrs = 1 Then MsgBox "The angle text box should only contain numbers", vbCritical, "Error": Exit Sub
hh1:
If Val(Text2.Text) > 360 Then Text2.Text = Val(Text2.Text) - 360: GoTo hh1
VALU = Slider1.Value
If Val(Text2.Text) <= 0 Then Text2.Text = "0": GoTo hhjh
ang = 360 / Val(Text2.Text)
ang = ((Pi * 2) / ang) + (Pi / 2)
hhjh:
For a = 1 To VALU
ExamplePic.ForeColor = rb
ExamplePic.DrawWidth = 1
ExamplePic.Line (xx + a * Cos(ang), yy + a * Sin(ang))-(xx + a * Cos(ang), yy + a * Sin(ang)), rb
ExamplePic.Print "Example"
Next a
GoTo ddff1
raisedd:
For X = xx - 1 To xx Step 1
For Y = yy - 1 To yy Step 1
ExamplePic.ForeColor = RGB(255, 255, 255)
ExamplePic.DrawWidth = 1: ExamplePic.Line (X, Y)-(X, Y), RGB(255, 255, 255): ExamplePic.Print "Example"
Next Y, X
For X = xx To xx + 1 Step 1
For Y = yy To yy + 1 Step 1
ExamplePic.ForeColor = RGB(0, 0, 0)
ExamplePic.DrawWidth = 1: ExamplePic.Line (X, Y)-(X, Y), RGB(0, 0, 0): ExamplePic.Print "Example"
Next Y, X
GoTo ddff1
ddd1:
If Check1.Value = False Then GoTo ddff1
For X = xx - 1 To xx + 1
For Y = yy - 1 To yy + 1
ExamplePic.ForeColor = rb
ExamplePic.DrawWidth = 1: ExamplePic.Line (X, Y)-(X, Y), rb: ExamplePic.Print "Example"
Next Y, X
ddff1:
ExamplePic.ForeColor = lb
ExamplePic.DrawWidth = 1: ExamplePic.Line (xx, yy)-(xx, yy), lb: ExamplePic.Print "Example"
ExamplePic.DrawWidth = Form1.Slider1.Value
End Sub
Private Sub Text2_Change()
Call draw_example
End Sub
