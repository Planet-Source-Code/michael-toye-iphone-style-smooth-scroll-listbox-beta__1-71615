VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   225
   ScaleWidth      =   945
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Public Event Click()
Public Event DblClick()

Private HiLite&
Private HiLite2&
Private LoLite&
Private Greyed&
Private Shadow&

Private mIconDC&
Private mIconW&
Private mIconH&
Private mIconCX&
Private mIconCY&
Private mBackStyle&
Private mDC&
Private mEnabled As Boolean
Private mBitmap&
Private cw&, ch&
Private mCaption$
Private mTextloc%

Public Property Let BackStyle(mBS&)
mBackStyle = mBS
PropertyChanged BackStyle
DrawControlUP
End Property
Public Property Get BackStyle() As Long
    BackStyle = mBackStyle
End Property
Public Property Let IconDC(mDC&)
    mIconDC = mDC
    DrawControlUP
    UserControl.Refresh
End Property
Public Property Let IconW(m&)
    mIconW = m
    mIconCX = (((UserControl.Width \ Screen.TwipsPerPixelX) / 2) - (m / 2)) - 1
End Property
Public Property Let IconH(m&)
    mIconH = m
    mIconCY = (((UserControl.Height \ Screen.TwipsPerPixelY) / 2) - (m / 2))
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Let Caption(s$)
    mCaption = s
    PropertyChanged Caption
    
    DrawControlUP
End Property
Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Enabled(s As Boolean)
    mEnabled = s
    If mEnabled Then
        DrawControlUP
    Else
        DrawControlGreyed
    End If
    
    PropertyChanged Enabled
    
End Property
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property
Public Property Let TextLocation(s As Integer)
    mTextloc = s
    DrawControlUP
End Property
Public Property Get TextLocation() As Integer
    TextLocation = mTextloc
End Property

Private Sub SplitRGB(ByVal clr&, r&, G&, B&)
    r = clr And &HFF: G = (clr \ &H100&) And &HFF: B = (clr \ &H10000) And &HFF
End Sub
Private Sub Gradient(dc&, X&, Y&, dx&, dy&, ByVal c1&, ByVal c2&, v As Boolean)
Dim r1&, G1&, B1&, r2&, G2&, B2&, B() As Byte
Dim i&, lR!, lG!, lB!, dR!, dG!, dB!, BI&(9), xx&, yy&, dd&, hRPen&
    If dx = 0 Or dy = 0 Then Exit Sub
    If v Then xx = 1: yy = dy: dd = dy Else xx = dx: yy = 1: dd = dx
    SplitRGB c1, r1, G1, B1: SplitRGB c2, r2, G2, B2: ReDim B(dd * 4 - 1)
    dR = (r2 - r1) / (dd - 1): lR = r1: dG = (G2 - G1) / (dd - 1): lG = G1: dB = (B2 - B1) / (dd - 1): lB = B1
    For i = 0 To (dd - 1) * 4 Step 4: B(i + 2) = lR: lR = lR + dR: B(i + 1) = lG: lG = lG + dG: B(i) = lB: lB = lB + dB: Next
    BI(0) = 40: BI(1) = xx: BI(2) = -yy: BI(3) = 2097153: StretchDIBits dc, X, Y, dx, dy, 0, 0, xx, yy, B(0), BI(0), 0, vbSrcCopy
End Sub

Sub CreateBitmaps(ByRef dc&, ByRef BM&, dx&, dy&)
Static t1&, t2&
    DeleteDC t1: DeleteDC t2
    t1 = GetDC(0): t2 = GetDC(0)
    dc = CreateCompatibleDC(t1)
    BM = CreateCompatibleBitmap(t2, dx, dy)
    SelectObject dc, BM
End Sub
Sub UnloadContexts(ByRef dc&, ByRef BM&)
    DeleteDC dc
    DeleteObject BM
End Sub
Function textX()
    textX = ((UserControl.Width - UserControl.TextWidth(mCaption)) \ 2) \ Screen.TwipsPerPixelX
End Function
Function textY()
    textY = -1 + (((UserControl.Height - UserControl.TextHeight(mCaption)) \ 2) \ Screen.TwipsPerPixelY)
End Function


Private Sub UserControl_Click()
    If mEnabled Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If mEnabled Then
        DrawControlDown
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_Initialize()
    CreateBitmaps mDC, mBitmap, cw, ch
    
'    HiLite = RGB(215, 215, 215)
'    HiLite2 = RGB(255, 255, 255)
'    LoLite = RGB(165, 165, 165)
'    Shadow = RGB(150, 150, 150)
'    Greyed = RGB(190, 190, 190)
    
    HiLite = RGB(196, 196, 196)
    HiLite2 = RGB(255, 255, 255)
    LoLite = RGB(150, 150, 150)
    Shadow = RGB(114, 114, 114)
    Greyed = RGB(190, 190, 190)
    
    mBackStyle = 2
    
    
End Sub

'Gray
'RGB 196
'150
'
'darker 114
'
'lcd top (253,254,231) - (214,219,191)

Private Sub UserControl_InitProperties()
    mCaption = "Command"
    mEnabled = True
    mBackStyle = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled Then DrawControlDown
End Sub
Sub DrawControlDown()
    If mBackStyle = 1 Then
        Gradient UserControl.hdc, 0, 0, cw, (ch / 2) + 1, HiLite, HiLite2, True
        Gradient UserControl.hdc, 0, (ch / 2) + 1, cw, ch / 2, HiLite2, HiLite, True
        
        DrawLine UserControl.hdc, 0, 0, cw - 1, 0, Shadow
        DrawLine UserControl.hdc, 0, 0, 0, ch - 1, Shadow
    ElseIf mBackStyle = 2 Then
        Gradient UserControl.hdc, 0, 0, cw, ch, HiLite2, HiLite, True
        
        DrawLine UserControl.hdc, 0, 0, cw - 1, 0, LoLite
        DrawLine UserControl.hdc, 0, 0, 0, ch - 1, LoLite
    ElseIf mBackStyle = 3 Then
        Gradient UserControl.hdc, 0, 0, cw, (ch / 2) + 1, RGB(108, 168, 250), RGB(59, 109, 219), True
        Gradient UserControl.hdc, 0, (ch / 2) + 1, cw, ch / 2, RGB(59, 109, 219), RGB(118, 178, 255), True
        
        DrawLine UserControl.hdc, 0, 0, cw - 1, 0, RGB(39, 89, 199)
        DrawLine UserControl.hdc, 0, 0, 0, ch - 1, RGB(39, 89, 199)
    ElseIf mBackStyle = 4 Then
        Gradient UserControl.hdc, 0, 0, cw, (ch / 2) + 1, RGB(115, 115, 115), 0, True
        Gradient UserControl.hdc, 0, (ch / 2) + 1, cw, ch / 2, 0, RGB(115, 115, 115), True
        
        DrawLine UserControl.hdc, 0, 0, cw - 1, 0, 0
        DrawLine UserControl.hdc, 0, 0, 0, ch - 1, 0
    End If
    
    If mIconDC <> 0 Then
        BitBlt UserControl.hdc, mIconCX + 1, mIconCY + 1, mIconW, mIconH, mIconDC, 0, 0, vbSrcCopy
    Else
        If mBackStyle = 3 Or mBackStyle = 4 Then
            SetTextColor UserControl.hdc, vbWhite
            TextOut UserControl.hdc, textX + 2, textY + 1, mCaption, Len(mCaption)
        Else
        
            SetTextColor UserControl.hdc, RGB(220, 220, 220)
            TextOut UserControl.hdc, textX + 2, textY + 2, mCaption, Len(mCaption)
            SetTextColor UserControl.hdc, 0
            TextOut UserControl.hdc, textX + 2, textY + 1, mCaption, Len(mCaption)
    
        End If
    End If
    UserControl.Refresh
End Sub
Sub DrawControlUP()
    If mBackStyle = 1 Then
        Gradient UserControl.hdc, 0, 0, cw, ch / 2, HiLite, HiLite2, True
        Gradient UserControl.hdc, 0, ch / 2, cw, ch / 2, HiLite, LoLite, True
        
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, Shadow
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, Shadow
    ElseIf mBackStyle = 2 Then
        Gradient UserControl.hdc, 0, 0, cw, ch, HiLite2, HiLite, True
    
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, LoLite
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, LoLite
    ElseIf mBackStyle = 3 Then
        Gradient UserControl.hdc, 0, 0, cw, ch, RGB(108, 168, 250), RGB(59, 109, 219), True
        Gradient UserControl.hdc, 0, 0, cw, (ch / 2), RGB(59, 109, 219), RGB(118, 178, 255), True
        
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, RGB(39, 89, 199)
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, RGB(39, 89, 199)
    ElseIf mBackStyle = 4 Then
        Gradient UserControl.hdc, 0, 0, cw, ch / 2, 0, RGB(115, 115, 115), True
        Gradient UserControl.hdc, 0, ch / 2, cw, (ch / 2) + 1, 0, RGB(50, 50, 50), True
    End If
    If mIconDC <> 0 Then
        BitBlt UserControl.hdc, mIconCX, mIconCY, mIconW, mIconH, mIconDC, 0, 0, vbSrcCopy
    Else
        If mBackStyle = 3 Or mBackStyle = 4 Then
            SetTextColor UserControl.hdc, vbWhite
            TextOut UserControl.hdc, textX, textY, mCaption, Len(mCaption)
        Else
            SetTextColor UserControl.hdc, RGB(220, 220, 220)
            TextOut UserControl.hdc, textX, textY + 1, mCaption, Len(mCaption)
            SetTextColor UserControl.hdc, 0
            TextOut UserControl.hdc, textX, textY, mCaption, Len(mCaption)
        End If
    End If
    
    UserControl.Refresh

End Sub
Sub DrawControlGreyed()
    If mBackStyle = 1 Then
        Gradient UserControl.hdc, 0, 0, cw, ch / 2, HiLite, HiLite2, True
        Gradient UserControl.hdc, 0, ch / 2, cw, ch / 2, HiLite, LoLite, True
        
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, Shadow
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, Shadow
    ElseIf mBackStyle = 2 Then
        Gradient UserControl.hdc, 0, 0, cw, ch, HiLite2, HiLite, True
    
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, LoLite
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, LoLite
    ElseIf mBackStyle = 3 Then
        Gradient UserControl.hdc, 0, 0, cw, ch, RGB(108, 168, 250), RGB(59, 109, 219), True
        Gradient UserControl.hdc, 0, 0, cw, (ch / 2), RGB(59, 109, 219), RGB(118, 178, 255), True
        
        DrawLine UserControl.hdc, 0, ch - 1, cw, ch - 1, RGB(39, 89, 199)
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, RGB(39, 89, 199)
    ElseIf mBackStyle = 4 Then
        Gradient UserControl.hdc, 0, 0, cw, ch / 2, 0, RGB(115, 115, 115), True
        Gradient UserControl.hdc, 0, ch / 2, cw, (ch / 2) + 1, 0, RGB(50, 50, 50), True
    End If
    
    If mIconDC <> 0 Then
        BitBlt UserControl.hdc, mIconCX, mIconCY, mIconW, mIconH, mIconDC, 0, 0, vbSrcCopy
    Else
        'SetTextColor UserControl.hdc, HiLite
        'TextOut UserControl.hdc, textX + 1, textY + 1, mCaption, Len(mCaption)
        SetTextColor UserControl.hdc, IIf(mBackStyle = 3 Or mBackStyle = 4, HiLite, Shadow)
        TextOut UserControl.hdc, textX, textY, mCaption, Len(mCaption)
    End If

    
    UserControl.Refresh

End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled Then DrawControlUP
End Sub
Public Sub Refresh()
    UserControl_Paint
End Sub
Private Sub UserControl_Paint()
    If mEnabled Then
        DrawControlUP
    Else
        DrawControlGreyed
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
         mCaption = .ReadProperty("Caption", "Command")
         mEnabled = .ReadProperty("Enabled", True)
         mBackStyle = .ReadProperty("BackStyle", 1)
    End With
End Sub


Private Sub UserControl_Resize()

    UserControl.BackColor = UserControl.ParentControls.Item(0).BackColor

    cw = UserControl.Width \ Screen.TwipsPerPixelX
    ch = UserControl.Height \ Screen.TwipsPerPixelY
    If mEnabled Then
        DrawControlUP
    Else
        DrawControlGreyed
    End If
End Sub

Private Sub UserControl_Show()
    If mEnabled Then
        DrawControlUP
    Else
        DrawControlGreyed
    End If
End Sub

Private Sub UserControl_Terminate()
UnloadContexts mDC, mBitmap
End Sub
Sub DrawLine(ByRef dc&, X1&, Y1&, X2&, Y2&, c&)
Dim p&, Pt As POINTAPI
    p = CreatePen(0, 1, c): DeleteObject SelectObject(dc, p)
    Pt.X = X1: Pt.Y = Y1
    MoveToEx dc, X1, Y1, Pt: LineTo dc, X2, Y2
    DeleteDC p
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", mCaption
        .WriteProperty "Enabled", mEnabled
        .WriteProperty "BackStyle", mBackStyle
    End With
End Sub
