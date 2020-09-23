VERSION 5.00
Begin VB.UserControl mtListBox 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   2430
   ScaleWidth      =   1380
End
Attribute VB_Name = "mtListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Enum SortField
    Key
    Value1
    Value2
End Enum

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
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private oLine As LineGS

Public Event Click()
Public Event DblClick()
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private mCaption As String
Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Const Message = "Hello !"
Const OPAQUE = 2
Const TRANSPARENT = 1
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2


Private mListBoxBlank As cDIBSection
Private mListBoxBuffer As cDIBSection
Private mListBoxText As cDIBSection
Private mListBoxTextMask As cDIBSection
Private dis As Long
Private uW As Long
Private uH As Long
Private uHdrH As Long
Private uFtrH As Long
Private uIconW As Long
Private uTickW As Long
Private uTextAreaT As Long
Private uTextAreaH As Long
Private uTextAreaL As Long
Private uTextAreaW As Long
Private uColWidth(1) As Long
Private uColGap As Long
Private uCol2TxtStart As Long
Private uTextH As Long
Private uTextGp As Long
Private mDataItem() As cDataItem
Private MouseIsDown As Boolean
Private MouseIsUpStillScrolling As Boolean
Private MouseDownStop As Boolean
Private MouseDownStartY As Long
Private MouseDownStartTick As Long
Private mListYpos As Long
Private bFastScrollUP As Boolean
Private bFastScrollDOWN As Boolean
Private bOnlyOneColumn As Boolean
Private mListBoxCaption As String
Private bNoScrollingTooShort As Boolean
Private StartScrollTick As Long
Private mSort As SortField
Private ClickIndex As Long
Private ClickText As String
Private IndexToDataItem() As Long
Private DataItemIndex() As Long



Public Property Get ItemByIndex(v As Long) As String
ItemByIndex = mDataItem(v).Value
End Property
Public Property Get Item2ByIndex(v As Long) As String
Item2ByIndex = mDataItem(v).Value2
End Property
Public Property Get ItemByKey(v As String) As String
Dim i&
For i = 0 To UBound(mDataItem)
    If v = mDataItem(i).Key Then
        ItemByKey = mDataItem(i).Value
        Exit For
    End If
Next
End Property
Public Property Get Item2ByKey(v As String) As String
Dim i&
For i = 0 To UBound(mDataItem)
    If v = mDataItem(i).Key Then
        Item2ByKey = mDataItem(i).Value2
        Exit For
    End If
Next
End Property
Public Property Get KeyByValue(v As String) As String
Dim i&
For i = 0 To UBound(mDataItem)
    If v = mDataItem(i).Value Or v = mDataItem(i).Value2 Then
        KeyByValue = mDataItem(i).Key
        Exit For
    End If
Next
End Property



Public Property Get ItemCount() As Long
ItemCount = UBound(mDataItem) + 1
End Property

Public Property Let SortBy(v As SortField)
mSort = v
End Property
Public Property Get SortBy() As SortField
SortBy = mSort
End Property
Public Property Let Caption(B As String)
mListBoxCaption = B
CreateBlank
End Property

Public Property Get Caption() As String
Caption = mListBoxCaption
End Property

Public Property Let OnlyOneColumn(B As Boolean)
bOnlyOneColumn = B
SetColWidth
End Property
Public Property Get OnlyOneColumn() As Boolean
OnlyOneColumn = bOnlyOneColumn

End Property

Public Sub AddItem(sKey As String, sValue As String, Optional sValue2 As String)
Dim n&
    n = KeyExists(sKey)
    If n = -1 Then
        If Not mDataItem(UBound(mDataItem)) Is Nothing Then
            ReDim Preserve mDataItem(UBound(mDataItem) + 1)
            RecreateTextDIB UBound(mDataItem)
        End If
        Set mDataItem(UBound(mDataItem)) = New cDataItem
        n = UBound(mDataItem)
    End If
    
    mDataItem(n).Key = sKey
    mDataItem(n).Value = sValue
    mDataItem(n).Deleted = False
    
    If Not IsMissing(sValue2) Then mDataItem(n).Value2 = sValue2
    mDataItem(n).Ticked = False
    
    LinkSort n
    
    RepopulateTextDIB
    mListYpos = 0
    BlankToBuffer
    TextToBuffer mListYpos
    BufferToUsercontrol
End Sub
Public Sub DeleteItemByKey(sKey As String)
Dim i&, n&, bFound As Boolean
    If UBound(mDataItem) = 0 And mDataItem(0) Is Nothing Then Exit Sub
    bFound = False
    For i = 0 To UBound(mDataItem)
        If sKey = mDataItem(i).Key Then
            bFound = True
            Exit For
        End If
    Next
    If Not bFound Then Exit Sub
    n = i
    mDataItem(i).Deleted = True
    For n = i + 1 To UBound(mDataItem)
        Set mDataItem(n - 1) = mDataItem(n)
    Next
    ReDim Preserve mDataItem(UBound(mDataItem) - 1)
    
    RecreateTextDIB UBound(mDataItem)
    
    'must redo the linked list
    For n = 0 To UBound(mDataItem)
        mDataItem(n).NextSortItem = -1
        mDataItem(n).PrevSortItem = -1
        mDataItem(n).IsRoot = False
        LinkSort n, True
    Next
    
    RepopulateTextDIB
    
    mListYpos = 0
    BlankToBuffer
    TextToBuffer mListYpos
    BufferToUsercontrol

End Sub
Private Sub LinkSort(v&, Optional bForce As Boolean = False)
Dim i&, pr&, nx&, rt&, t_pr&, t_nx&

    If (UBound(mDataItem) = 0 And v = 0) Or (v = 0 And bForce) Then
        mDataItem(v).IsRoot = True
        mDataItem(v).NextSortItem = -1
        mDataItem(v).PrevSortItem = -1
        Exit Sub
    End If
    
    For rt = 0 To UBound(mDataItem)
        If mDataItem(rt).IsRoot Then Exit For
    Next
    i = rt
    Do
        nx = mDataItem(i).NextSortItem: pr = mDataItem(i).PrevSortItem
        
        If pr = -1 Then 'i is root
            If SortVal(i, mSort) > SortVal(v, mSort) Then 'v must be a new root!
                mDataItem(v).IsRoot = True: mDataItem(i).IsRoot = False
                mDataItem(v).NextSortItem = i: mDataItem(i).PrevSortItem = v: mDataItem(v).PrevSortItem = -1
                Exit Sub
            End If
        End If
        If nx = -1 Then 'i is end of list
            If SortVal(i, mSort) < SortVal(v, mSort) Then 'v must be a new end of list!
                mDataItem(v).PrevSortItem = i: mDataItem(v).NextSortItem = -1: mDataItem(i).NextSortItem = v
                Exit Sub
            End If
        End If
        'Debug.Print SortVal(i, mSort) & " < " & SortVal(v, mSort) & " And " & SortVal(v, mSort) & " <= " & SortVal(nx, mSort)
        If SortVal(i, mSort) < SortVal(v, mSort) And SortVal(v, mSort) <= SortVal(nx, mSort) Then
            mDataItem(i).NextSortItem = v: mDataItem(v).NextSortItem = nx
            mDataItem(v).PrevSortItem = i: mDataItem(nx).PrevSortItem = v
            Exit Sub
        End If
        
        pr = i: i = nx: nx = mDataItem(i).NextSortItem
        
    Loop Until i = -1
    
End Sub

Private Function SortVal$(n&, s As SortField)
If s = Key Then
    SortVal = mDataItem(n).Key
ElseIf s = Value1 Then
    SortVal = mDataItem(n).Value
Else
    SortVal = mDataItem(n).Value2
End If
End Function

Private Sub RecreateTextDIB(nLines&)
Dim h&
    h = uTextGp + (nLines * (uTextGp + uTextH))
    Set mListBoxText = New cDIBSection
    mListBoxText.Create uW, h
    SetBkMode mListBoxText.hdc, TRANSPARENT
    SetFont mListBoxText.hdc, "Segoe UI", 9
    
    Set mListBoxTextMask = New cDIBSection
    mListBoxTextMask.Create uW, h
    
    SetBkMode mListBoxTextMask.hdc, TRANSPARENT
    SetFont mListBoxTextMask.hdc, "Segoe UI", 9
    
    bNoScrollingTooShort = (h <= uTextAreaH)
    
End Sub
Private Sub RepopulateTextDIB()
Dim h&, Y&, i&, ColTxt1$, ColTxt2$, yOf&, n&
    yOf = 0
    h = uTextGp + ((UBound(mDataItem)) * (uTextGp + uTextH))
    Gradient mListBoxText.hdc, 0, 0, uW, h, RGB(255, 255, 255), RGB(230, 230, 230), False
    ReDim IndexToDataItem(UBound(mDataItem))
    ReDim DataItemIndex(UBound(mDataItem))
    For i = 0 To UBound(mDataItem)
        If mDataItem(i).IsRoot Then Exit For
    Next
    n = 0
    For Y = uTextGp To h Step (uTextGp + uTextH)
        IndexToDataItem(n) = i
        DataItemIndex(i) = n
        If mDataItem(i).Ticked Then
            '#### set 'ticked' banner gradient colors ####
            Gradient mListBoxText.hdc, 0, Y, uW - 12, uTextH, RGB(200, 200, 200), RGB(230, 230, 230), False
        End If
        
        ColTxt1 = SlimTextToCol(mDataItem(i).Value, 0)
        If Not bOnlyOneColumn Then ColTxt2 = SlimTextToCol(mDataItem(i).Value2, 1)
        TextOut mListBoxText.hdc, uColGap, Y, ColTxt1, Len(ColTxt1)
        If Not bOnlyOneColumn Then TextOut mListBoxText.hdc, uCol2TxtStart, Y, ColTxt2, Len(ColTxt2)

        n = n + 1
        i = mDataItem(i).NextSortItem
        If i = -1 Then Exit For
    Next
End Sub
Private Function SlimTextToCol$(sVal$, col&)
If Len(sVal) = 0 Then Exit Function
Dim tw&, i&
    tw = UserControl.TextWidth(sVal) \ Screen.TwipsPerPixelX
    If tw <= uColWidth(col) Then
        SlimTextToCol = sVal
        Exit Function
    End If
    For i = Len(sVal) - 3 To 1 Step -1
        SlimTextToCol = Left(sVal, i) + "..."
        If (UserControl.TextWidth(SlimTextToCol) \ Screen.TwipsPerPixelX) <= uColWidth(col) Then Exit Function
    Next
    
End Function
Private Function KeyExists(ByVal sKey As String) As Long
Dim i&
KeyExists = -1
For i = 0 To UBound(mDataItem)
    If Not mDataItem(i) Is Nothing Then
        If mDataItem(i).Key = sKey Then
            KeyExists = i
            Exit Function
        End If
    End If
Next

End Function

Private Sub UserControl_Initialize()
    Set mListBoxBlank = New cDIBSection
    Set mListBoxBuffer = New cDIBSection
    Set oLine = New LineGS
    
    ReDim mDataItem(0)
    RecreateTextDIB UBound(mDataItem) + 1
    
    '#### set font ####
    UserControl.FontName = "Segoe UI"
    UserControl.FontSize = 9
    uTextH = UserControl.TextHeight("HHttjjpq|!") \ Screen.TwipsPerPixelY
    uTextGp = 3
    
    mListYpos = 0
    
End Sub
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 And Not bFastScrollUP Then
        bFastScrollUP = True
        FastScroll
    ElseIf KeyCode = 40 And Not bFastScrollDOWN Then
        bFastScrollDOWN = True
        FastScroll
    Else
        If KeyCode >= 65 And KeyCode <= 90 Then
            FastScrollToLetter KeyCode
        End If
    End If

End Sub
Private Sub FastScrollToLetter(KC%)
Dim n&, iDir&
    n = DataItemIndex(FirstIndexOfLetter(KC)) * (uTextGp + uTextH)
    iDir = 2
    If mListYpos > n Then iDir = -2
    Do
        mListYpos = mListYpos + iDir
        BlankToBuffer
        TextToBuffer mListYpos
        BufferToUsercontrol
        If iDir < 0 And mListYpos <= 0 Then
            mListYpos = 0
            Exit Do
        End If
        If iDir > 0 And mListYpos >= (mListBoxText.Height - uTextAreaH) Then
            mListYpos = (mListBoxText.Height - uTextAreaH)
            Exit Do
        End If

    Loop Until Abs(mListYpos - n) < 3
    
End Sub
Private Function FirstIndexOfLetter&(KC%)
Dim i&

For i = 0 To UBound(mDataItem)
    If mDataItem(i).IsRoot Then Exit For
Next
FirstIndexOfLetter = -1
Do
    If Asc(UCase(Left(mDataItem(i).Value, 1))) >= KC Then
        FirstIndexOfLetter = i
        Exit Do
    End If
    i = mDataItem(i).NextSortItem
Loop Until i = -1 Or FirstIndexOfLetter >= 0
End Function

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If bFastScrollUP Then
    bFastScrollUP = False
End If
If bFastScrollDOWN Then
    bFastScrollDOWN = False
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bNoScrollingTooShort Then Exit Sub
    
    If MouseIsDown Then
        MouseDownStop = True
        MouseIsUpStillScrolling = False
        Exit Sub
    End If

Dim n&
    ClickText = ""
    n = Y \ Screen.TwipsPerPixelY
    If n < uHdrH Then
        bFastScrollUP = True
    ElseIf n > (uH - uFtrH) Then
        bFastScrollDOWN = True
    Else
        MouseIsDown = True
    End If
    If bFastScrollDOWN Or bFastScrollUP Then
        StartScrollTick = GetTickCount
        FastScroll
    End If
End Sub
Private Sub FastScroll()
Dim bFinish As Boolean, i&, t&
    Do
        t = GetTickCount - StartScrollTick
        If t < 1000 Then
            i = 2
        ElseIf t < 2000 Then
            i = 4
        ElseIf t < 4000 Then
            i = 8
        Else
            i = 16
        End If
        If Not (bFastScrollDOWN Or bFastScrollUP) Then bFinish = True
        If bFastScrollUP And mListYpos <= 0 Then
            mListYpos = 0
            bFinish = True
        End If
        If bFastScrollDOWN And mListYpos >= (mListBoxText.Height - uTextAreaH) Then
            mListYpos = (mListBoxText.Height - uTextAreaH)
            bFinish = True
        End If
        If Not bFinish Then
            If bFastScrollUP Then
                mListYpos = mListYpos - i
            Else
                mListYpos = mListYpos + i
            End If
            BlankToBuffer
            TextToBuffer mListYpos
            BufferToUsercontrol
        End If
        Sleep 2
        DoEvents

    Loop Until bFinish
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static prevY&


    If bNoScrollingTooShort Then Exit Sub

    If Not MouseIsDown Then
        prevY = -1
        Exit Sub
    End If

    If prevY = -1 Then prevY = Y
    dis = (Y - prevY) \ Screen.TwipsPerPixelY * -1
    prevY = Y
    mListYpos = mListYpos + dis
    If mListYpos < 0 Then
        mListYpos = 0
        Exit Sub
    End If
    If mListYpos >= (mListBoxText.Height - uTextAreaH) Then
        mListYpos = (mListBoxText.Height - uTextAreaH)
        Exit Sub
    End If
    TextToBuffer mListYpos
    BufferToUsercontrol

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t$, n&
    If dis = 0 And Y >= (uHdrH * Screen.TwipsPerPixelY) And Y <= ((uH - uHdrH) * Screen.TwipsPerPixelY) Then 'it's a fair bet user just clicked on an item
        ClickIndex = ((Y \ Screen.TwipsPerPixelY) - uHdrH + mListYpos) \ (uTextGp + uTextH)
        n = IndexToDataItem(ClickIndex)
        mDataItem(n).Ticked = Not mDataItem(n).Ticked
        RepopulateTextDIB
        BlankToBuffer
        TextToBuffer mListYpos
        BufferToUsercontrol
    End If
    
    If bNoScrollingTooShort Then Exit Sub
        
    bFastScrollDOWN = False
    bFastScrollUP = False
    If Not MouseIsDown Then Exit Sub
    Dim pos&
    MouseDownStop = False
    MouseIsUpStillScrolling = True
    MouseIsDown = False

    Do
        dis = dis + IIf(dis < 0, 1, -1)
        mListYpos = mListYpos + dis
        If mListYpos < 0 Then mListYpos = 0
        If mListYpos > (mListBoxText.Height - uTextAreaH) Then mListYpos = (mListBoxText.Height - uTextAreaH)
        BlankToBuffer
        TextToBuffer mListYpos
        BufferToUsercontrol
        Sleep 40
        DoEvents
    Loop Until dis = 0 Or mListYpos = 0 Or mListYpos = (mListBoxText.Height - uTextAreaH) Or MouseDownStop
    dis = 0
    MouseIsUpStillScrolling = True
    If MouseDownStop Then Exit Sub
    
End Sub

Private Sub UserControl_Paint()
    BlankToBuffer
    TextToBuffer mListYpos
    BufferToUsercontrol
End Sub
Private Sub SetColWidth()
Dim hp&

    '#### column 2 width is currently 30% of total ####
    
    If bOnlyOneColumn Then
        uColWidth(0) = uW - (2 * uColGap)
    Else
        hp = (uW - (uColGap * 3))
        uColWidth(0) = hp * 0.7
        uColWidth(1) = hp - uColWidth(0)
        uCol2TxtStart = (2 * uColGap) + uColWidth(0)
    End If
End Sub
Private Sub UserControl_Resize()

    uW = UserControl.Width \ Screen.TwipsPerPixelX: uH = UserControl.Height \ Screen.TwipsPerPixelY
    
    '#### header and footer heights ####
    uHdrH = 20
    uFtrH = 20
    
    
    uIconW = 20
    uTickW = 20
    uColGap = 3

    SetColWidth

    uTextAreaT = uHdrH
    uTextAreaH = uH - uHdrH - uFtrH
    uTextAreaL = uIconW
    uTextAreaW = uW - uIconW - uTickW
    
    With mListBoxBuffer
        .ClearUp
        .Create uW, uH
        SetBkMode .hdc, TRANSPARENT
        SetFont .hdc, "Segoe UI", 9
        SetTextColor .hdc, vbWhite
    End With

    CreateBlank
    
    BlankToBuffer
    TextToBuffer mListYpos
    BufferToUsercontrol
    
End Sub
Private Sub CreateBlank()
    With mListBoxBlank
        .ClearUp
        .Create uW, uH
        '#### header gradient colours ####
        If uHdrH > 0 Then Gradient .hdc, 0, 0, uW, uHdrH, RGB(108, 168, 250), RGB(59, 109, 219), True
        '#### footer gradient colours ####
        If uFtrH > 0 Then Gradient .hdc, 0, uH - uFtrH, uW, uFtrH, RGB(108, 168, 250), RGB(59, 109, 219), True
        '#### MAIN TEXT AREA gradient colours ####
        Gradient .hdc, 0, uTextAreaT, uW, uTextAreaH, RGB(255, 255, 255), RGB(230, 230, 230), False
    End With
    SetBkMode mListBoxBlank.hdc, TRANSPARENT
    SetFont mListBoxBlank.hdc, "Segoe UI", 9
    SetTextColor mListBoxBlank.hdc, vbWhite
    TextOut mListBoxBlank.hdc, 3, 2, mListBoxCaption, Len(mListBoxCaption)
End Sub
Private Sub LocationContextBarToBuffer(Y&)
Dim LineLen&, LineStart
    LineLen = ((uTextAreaH / mListBoxText.Height) * uTextAreaH) - 4

    LineStart = (Y / mListBoxText.Height) * uTextAreaH + uHdrH + 2
    oLine.LineGP mListBoxBuffer.hdc, uW - 4, LineStart, uW - 4, LineStart + LineLen, RGB(190, 190, 190)
End Sub
Private Sub TextToBuffer(Y&)
    BitBlt mListBoxBuffer.hdc, 0, uTextAreaT, uW, uTextAreaH, mListBoxText.hdc, 0, Y, vbSrcCopy
    If Not bNoScrollingTooShort Then LocationContextBarToBuffer Y
    If Len(ClickText) > 0 Then
        TextOut mListBoxBuffer.hdc, 3, uH - uTextH - 3, ClickText, Len(ClickText)
    End If
End Sub
Private Sub BlankToBuffer()
    BitBlt mListBoxBuffer.hdc, 0, 0, uW, uH, mListBoxBlank.hdc, 0, 0, vbSrcCopy
End Sub
Private Sub BufferToUsercontrol()
    BitBlt UserControl.hdc, 0, 0, uW, uH, mListBoxBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Private Function CreateMyFont(nSize&, sFontFace$, bBold As Boolean, bItalic As Boolean) As Long
Static r&, d&

    DeleteDC r: r = GetDC(0)
    d = GetDeviceCaps(r, LOGPIXELSY)
    CreateMyFont = CreateFont(-MulDiv(nSize, d, 72), 0, 0, 0, _
                              IIf(bBold, FW_BOLD, FW_NORMAL), bItalic, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                              CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, sFontFace) 'gdi 2
End Function

Private Sub SetFont(dc&, sFace$, nSize&)
Static c&
    ReleaseDC dc, c: DeleteDC c
    c = CreateMyFont(nSize, sFace, False, False)
    DeleteObject SelectObject(dc, c)
End Sub
