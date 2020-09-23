Attribute VB_Name = "Module1"
Option Explicit
Global PicName, lb, rb, zoomactive As Boolean, BrushType, RepRed, RepGre, repBlu, progress, NumSides, AtAngle, Rx, Ry, PolyX, PolyY, CopyX, CopyY
Global ImageArray(4, 1500, 1500) As Integer
Global X, Y As Integer
Global larrCol() As Long
Global Const CB_HEIGHT = 400
Global Const Pi = 3.14159265359
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42
Public Const ThisApp = "Stu Paint V2"
Public Const ThisKey = "Recent Files"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function CloseClipBoard Lib "user32" Alias "CloseClipboard" () As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As String) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal nCount As Long) As Long
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Function Oldpic()
Dim a
For a = 4 To 1 Step -1
Form1.UndoPicBox(a).Picture = Form1.UndoPicBox(a - 1).Image
Form1.UndoPicBox(a).Refresh
Next a
Form1.UndoPicBox(0).Picture = Form1.MainPic.Image
Form1.UndoPicBox(0).Refresh
End Function
Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function
Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function
Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Public Function replace_routine()
Dim r2, g2, b2, color1, r, g, b
r2 = RepRed / 100
g2 = RepGre / 100
b2 = repBlu / 100
For X = 0 To Form1.MainPic.ScaleWidth - 1
For Y = 0 To Form1.MainPic.ScaleHeight - 1
color1 = GetPixel(Form1.MainPic.hdc, X, Y)
r = (color1 Mod 256)
b = (Int(color1 / 65536))
g = ((color1 - (b * 65536) - r) / 256)
r = Abs(r * r2)
b = Abs(b * b2)
g = Abs(g * g2)
SetPixelV Form1.MainPic.hdc, X, Y, RGB(r, g, b)
Next Y
progress = (100 / (Form1.MainPic.ScaleWidth - 1)) * X
Call progressbar
Next X
Form1.MainPic.Refresh
End Function
Public Sub progressbar()
If Form1.Proview.Checked = False Then Exit Sub
Form1.Progpic(0).Cls
Form1.Progpic(0).ForeColor = RGB(192, 192, 192)
Form1.Progpic(0).Line (CByte(progress), 0)-(100, 200), , BF
Form1.Progpic(0).Line (45, 0)-(45, 0)
Form1.Progpic(0).ForeColor = RGB(0, 0, 0)
Form1.Progpic(0).Print CByte(progress)
Form1.StatusBar1.Panels.Item(1).Picture = Form1.Progpic(0).Image
End Sub
Public Function FileExist(sFileN As String) As Boolean
    Dim tmpRv As Long
    
    On Error Resume Next
    tmpRv = GetAttr(sFileN)
    If Err Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Sub GetRecentFiles()
    Dim i, j As Integer
    Dim varFiles As Variant
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    varFiles = GetAllSettings(ThisApp, ThisKey)
    For i = 0 To UBound(varFiles, 1)
        Form1.sepfil.Visible = True
        Form1.mnuRecentFile(0).Visible = True
        Form1.mnuRecentFile(i).Caption = varFiles(i, 1)
        Form1.mnuRecentFile(i).Visible = True
    Next i
End Sub
Sub UpdateFileMenu(Filename)
        Dim intRetVal As Integer
        intRetVal = OnRecentFilesList(Filename)
        If Not intRetVal Then
            WriteRecentFiles (Filename)
        End If
        GetRecentFiles
End Sub
Function OnRecentFilesList(Filename) As Integer
  Dim i
  For i = 1 To 4
    If Form1.mnuRecentFile(i).Caption = Filename Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function
Sub WriteRecentFiles(OpenFileName)
    Dim i, j As Integer
    Dim strFile, key As String
    For i = 3 To 1 Step -1
        key = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, key)
        If strFile <> "" Then
            key = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, key, strFile
        End If
    Next i
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

