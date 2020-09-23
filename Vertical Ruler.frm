VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   540
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15.61
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   0.953
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      X1              =   0
      X2              =   1.482
      Y1              =   3.387
      Y2              =   3.387
   End
   Begin VB.Line Line1 
      X1              =   0.25
      X2              =   0.25
      Y1              =   -0.185
      Y2              =   35.269
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call StayOnTop
Me.Height = Form1.MainPic.Height + 270
For a = 0 To 60
Line (0.1, a)-(0.4, a), RGB(0, 0, 0)
ofst = 0.15
Line (0.5, a - ofst)-(0.5, a - ofst), RGB(0, 0, 0)
Print a
Next a
For a = 0 To 60 Step 0.1
Line (0.15, a)-(0.35, a), RGB(0, 0, 0)
Next a
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form6.hWnd, -1, 0, 0, 0, 0, &H1)
End Function
Private Sub Form_Unload(Cancel As Integer)
Form1.Rulers.Checked = False
Unload Me
End Sub

