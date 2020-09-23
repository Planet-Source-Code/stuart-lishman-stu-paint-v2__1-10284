VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Trace"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7155
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Finished"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Oldpic
Form1.MainPic.Picture = Picture1.Image
Unload Me
End Sub

Private Sub Form_Load()
Call StayOnTop
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form7.hWnd, -1, 0, 0, 0, 0, &H1)
End Function

