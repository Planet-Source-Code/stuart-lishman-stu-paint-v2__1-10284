VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Replacement"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      ScaleHeight     =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   201
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   2
         X1              =   -100.523
         X2              =   301.477
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         Index           =   2
         X1              =   97.455
         X2              =   97.455
         Y1              =   -10
         Y2              =   60
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   120
      ScaleHeight     =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   201
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         Index           =   1
         X1              =   -100.523
         X2              =   301.477
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         Index           =   1
         X1              =   97.455
         X2              =   97.455
         Y1              =   -10
         Y2              =   60
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         Index           =   0
         X1              =   97.455
         X2              =   97.455
         Y1              =   -10
         Y2              =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   0
         X1              =   -100.523
         X2              =   301.477
         Y1              =   24
         Y2              =   24
      End
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Blue:"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Green:"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Red:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "+100"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "-100"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cnt = 0
For a = 0 To 2
If Line2(a).X1 > 99 And Line2(a).X1 < 101 Then cnt = cnt + 1
Next a
If cnt = 3 Then GoTo hh
Call Oldpic
RepRed = Line2(0).X1
RepGre = Line2(1).X1
repBlu = Line2(2).X1
Call replace_routine
hh:
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Call StayOnTop
Me.left = Screen.Width - Me.Width
Me.top = 0
For a = 0 To 2
Label4(a) = 0
Line2(a).X1 = 100: Line2(a).X2 = 100
Next a
End Sub
Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then GoTo enn
If X < 0 Then X = 0
If X > 200 Then X = 200
Line2(Index).X1 = X
Line2(Index).X2 = X
Label4(Index) = Int(X) - 100
enn:
End Sub
Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then GoTo enn1
If X < 0 Then X = 0
If X > 200 Then X = 200
Line2(Index).X1 = X
Line2(Index).X2 = X
Label4(Index) = Int(X) - 100
enn1:
End Sub
Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then GoTo enn2
If X < 0 Then X = 0
If X > 200 Then X = 200
Line2(Index).X1 = X
Line2(Index).X2 = X
Label4(Index) = Int(X) - 100
enn2:
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form4.hWnd, -1, 0, 0, 0, 0, &H1)
End Function

