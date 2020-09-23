VERSION 5.00
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7995
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0.926
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   14.102
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      X1              =   2.117
      X2              =   2.117
      Y1              =   -9.999
      Y2              =   20.001
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   60.001
      Y1              =   0.25
      Y2              =   0.25
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop
Me.Width = Form1.MainPic.Width + 30
For a = 0 To 60
Line (a, 0.1)-(a, 0.4), RGB(0, 0, 0)
If a < 10 Then ofst = 0.13 Else ofst = 0.2
Line (a - ofst, 0.5)-(a - ofst, 0.5), RGB(0, 0, 0)
Print a
Next a
For a = 0 To 60 Step 0.1
Line (a, 0.15)-(a, 0.35), RGB(0, 0, 0)
Next a
End Sub
Private Function StayOnTop()
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(Form5.hWnd, -1, 0, 0, 0, 0, &H1)
End Function
Private Sub Form_Unload(Cancel As Integer)
Form1.Rilers.Checked = False
Unload Me
End Sub
