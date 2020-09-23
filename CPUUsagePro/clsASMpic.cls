VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASMpic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'************************************************
'* GioRock 2008                                 *
'* some parts of this class code is been erased *
'* to use rotation function only                *
'************************************************

'clsASMpic 1.0
'Buggy 2002/11
'use as you like.
'so far all assembler code by Robert Rayment, big thanks for your great work!

'clsASMpic is a simple class that makes possible REALTIME image manipulation of
'even large pictures in visual basic.
'this is accomplished using precompiled assembler code (machine code) in the
'time critical sections.
'it has a real wealth of functions, eg rotating and scaling with antialiasing,
'invertion and grayscaling, art effects like relief and smoothing,
'colour effects - each colour +/-, brightness and many more - plus some really
'nice and weird ideas.

'this is mostly a total conversion of one of Robert's former postings, to make
'using this great technique as easy as it is being fast:
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29034&lngWId=1
'also look there for explanation/additional info and more of his great work.

'I did my best to encapsulate everything into one single class and make calling
'its powerful functions as easy as possible. information hiding at its best.
'by using enums, the functions become self explanatory.
'decent error checking, to be improved (for example don't quit your prog until
'your loop of one degree rotation steps has finished...).
'even loading and saving pics with the standard dialog is handled.

'this could be made even faster if abandoning undo and reset function:
'I kept robert's original structure of 3 seperate copies in memory for each image:
'1 - original pic
'2 - current pic
'3 - undo pic
'if you need only parts of it, rip out the unneeded routines to keep the .exe small.

'/// D E C L A R A T I O N S ///////////////////////////////////////////////////

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
'To fill BITMAP structure
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
    (ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long
' APIs for getting DIB bits to PalBGR
Private Declare Function GetDIBits Lib "gdi32" _
    (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
    (ByVal hDC As Long) As Long
    
Private Declare Function CreateBitmap Lib "Gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "Gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetBkColor Lib "Gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "Gdi32.dll" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
    (ByVal hDC As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long
' For transferring drawing in an integer array to Form or PicBox
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal DesW As Long, ByVal DesH As Long, _
    ByVal SrcX As Long, ByVal SrcY As Long, _
    ByVal PICWW As Long, ByVal PICHH As Long, _
    lpBits As Any, lpBitsInfo As BITMAPINFO, _
    ByVal wUsage As Long, ByVal dwRop As Long) As Long
'For calling machine code
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpMCode As Long, _
    ByVal Long1 As Long, ByVal Long2 As Single, _
    ByVal Long3 As Single, ByVal Long4 As Long) As Long


'/// T Y P E S /////////////////////////////////////////////////////////////////

Private Type BITMAP
    bmType As Long    ' Type of bitmap
    bmWidth As Long    ' Pixel width
    bmHeight As Long    ' Pixel height
    bmWidthBytes As Long    ' Byte width = 3 x Pixel width
    bmPlanes As Integer    ' Color depth of bitmap
    bmBitsPixel As Integer    ' Bits per pixel, must be 16 or 24
    bmBits As Long    ' This is the pointer to the bitmap data  !!!
End Type
'NB PICTURE STORED IN MEMORY UPSIDE DOWN
'WITH INCREASING MEMORY GOING UP THE PICTURE
'bmp.bmBits points to the bottom left of the picture

' Structures for StretchDIBits
Private Type BITMAPINFOHEADER ' 40 bytes
    biSize As Long
    biwidth As Long
    biheight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiH As BITMAPINFOHEADER
'bmiH As RGBTRIPLE            'NB Palette NOT NEEDED for 16,24 & 32-bit
End Type
'MCode Structure for parameter passing to machine code functions
Private Type MCodeStruc
    PICW As Long
    PICH As Long
    PtrPalBGR As Long
    PtrPalLineCopy As Long
    Increment As Long
    QBLongColor As Long
    OpCode As Long
End Type

'/// E N U M S /////////////////////////////////////////////////////////////////

Private Enum eASMBinary
    PICROTATE = 1002
End Enum

'/// V A R S ///////////////////////////////////////////////////////////////////

Private bm As BITMAPINFO    'Info about pic (colordepth etc)
Private PalLineCopy(1, 1) As Byte    'For copying 1 line of PalBGR()
Private PicRotateMC() As Byte    'Array to hold machine code for Rotations
Private PIC As PictureBox    'Pointer to picbox to manipulate
Private PICW As Long, PICH As Long    'Picbox Width & Height (pixels)
Private PalBGR() As Byte    'To hold 3 full palettes (12 x PICW x PICH)
Private PalSize As Long    'Size of 1 palette (4 x PICW x PICH)
'finally made these private again as dimming in each subfunction costs time.
Private MCODE As MCodeStruc
Private ptrStruc As Long, ptMC As Long
'new private to switch autodrawing off, so that large filter arrays dont redraw after each step
Private m_AutoDraw As Boolean

Private Const BITSPIXEL = 12

Public Sub TransparentBlt(ByVal hDCDst As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal hDCSrc As Long, ByVal SrcX As Long, ByVal SrcY, ByVal SrcW As Long, ByVal SrcH, TransColor As Long)
'Parameter:

    'hDCDst- Device context into which image must be
    'drawn transparently
    
    'hDCSrc- Device context of source to be made transparent
    'in color TransColor
    
    'SrcX, SrcY, SrcW, SrcH - Rectangular region within
    'hDCSrc to be made transparent in terms of hDCDst,
    'and drawn to hDCDst
    
    'DstX, DstY - Coordinates in hDCDst (and hDCDst)
    'where the transparent bitmap must go.
    
    'TransColor - Transparent Color.
'-------------------------

Dim MonoMaskDC As Long, hMonoMask As Long
Dim MonoInvDC As Long, hMonoInv As Long
Dim ResulthDCDst As Long, hResultDst As Long
Dim ResulthDCSrc As Long, hResultSrc As Long
Dim hPrevMask As Long, hPrevInv As Long
Dim hPrevSrc As Long, hPrevDst As Long
Dim nRet As Long

    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(hDCDst)
    MonoInvDC = CreateCompatibleDC(hDCDst)
    hMonoMask = CreateBitmap(SrcW, SrcH, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(SrcW, SrcH, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
    'create keeper DCs and bitmaps
    ResulthDCDst = CreateCompatibleDC(hDCDst)
    ResulthDCSrc = CreateCompatibleDC(hDCDst)
    hResultDst = CreateCompatibleBitmap(hDCDst, SrcW, SrcH)
    hResultSrc = CreateCompatibleBitmap(hDCDst, SrcW, SrcH)
    hPrevDst = SelectObject(ResulthDCDst, hResultDst)
    hPrevSrc = SelectObject(ResulthDCSrc, hResultSrc)
   
    'copy src to monochrome mask
    Dim OldBC As Long
    OldBC = SetBkColor(hDCSrc, TransColor)
    nRet = BitBlt(MonoMaskDC, 0, 0, SrcW, SrcH, hDCSrc, SrcX, SrcY, vbSrcCopy)
    TransColor = SetBkColor(hDCSrc, OldBC)
   
    'create inverse of mask
    nRet = BitBlt(MonoInvDC, 0, 0, SrcW, SrcH, MonoMaskDC, 0, 0, vbNotSrcCopy)
   
    'get background
    nRet = BitBlt(ResulthDCDst, 0, 0, SrcW, SrcH, hDCDst, DstX, DstY, vbSrcCopy)
   
    'AND with Monochrome mask
    nRet = BitBlt(ResulthDCDst, 0, 0, SrcW, SrcH, MonoMaskDC, 0, 0, vbSrcAnd)
   
    'get overlapper
    nRet = BitBlt(ResulthDCSrc, 0, 0, SrcW, SrcH, hDCSrc, SrcX, SrcY, vbSrcCopy)
   
    'AND with inverse monochrome mask
    nRet = BitBlt(ResulthDCSrc, 0, 0, SrcW, SrcH, MonoInvDC, 0, 0, vbSrcAnd)
   
    'XOR these two
    nRet = BitBlt(ResulthDCDst, 0, 0, SrcW, SrcH, ResulthDCSrc, 0, 0, vbSrcInvert)
   
    'Display Transparent Image.
    nRet = BitBlt(hDCDst, DstX, DstY, SrcW, SrcH, ResulthDCDst, 0, 0, vbSrcCopy)
   
    'Free resource.
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask

    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv

    hResultDst = SelectObject(ResulthDCDst, hPrevDst)
    DeleteObject hResultDst

    hResultSrc = SelectObject(ResulthDCSrc, hPrevSrc)
    DeleteObject hResultSrc

    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResulthDCDst
    DeleteDC ResulthDCSrc
    
End Sub
Public Function UndoLast()
    If NoPicAssigned Then Exit Function
    CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 3), PalSize
    ShowPalBGR 2
End Function
Public Property Let PictureBox(ByRef PicBox As PictureBox)
Attribute PictureBox.VB_UserMemId = 0
    On Error GoTo InitializeError
    Set PIC = PicBox
    PICW = PIC.Width
    PICH = PIC.Height
    If Not SysBPPok(PIC) Then
        Set PIC = Nothing
        MsgBox "Not 16 or 24-bit color setting", vbCritical, "clsASM"
        Exit Property
    End If
    MCODE.PICW = PICW
    MCODE.PICH = PICH
    ptrStruc = VarPtr(MCODE.PICW)
'---------------------
    GeneratePalBGRs
    MCODE.PtrPalBGR = VarPtr(PalBGR(1, 1, 1, 1))
    MCODE.PtrPalLineCopy = VarPtr(PalLineCopy(1, 1))
'---------------------
    ShowPalBGR 3
'---------------------
    Exit Property
InitializeError:
    Set PIC = Nothing
    MsgBox "Couldn't assign picturebox", vbCritical, "clsASM"
End Property

Public Property Let AutoDraw(autoON As Boolean)
    m_AutoDraw = autoON
    If autoON Then
        ShowPalBGR 2 'if switched on again, draw last pic
    Else 'if switched off, update undo buffer
        CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 2), PalSize
    End If
End Property
Public Property Get AutoDraw() As Boolean
    AutoDraw = m_AutoDraw
End Property

Public Function DrawNow()
    ShowPalBGR 2
    DoEvents
End Function

Private Function NoPicAssigned() As Boolean
    If PIC Is Nothing Then
        NoPicAssigned = True
        MsgBox "You must assign a picturebox to the class first!", vbInformation, "clsASM"
    Else
        NoPicAssigned = False
    End If
End Function

Public Function ASM_Rotate(ByVal Angle As Integer, Optional ByVal AntiAlias As Boolean = False, Optional ByVal SourcePic As Boolean = True, Optional ByVal colorX As Long = -1, Optional ByVal centerX As Variant, Optional ByVal centerY As Variant)
    
    If NoPicAssigned Then Exit Function
    
    If Angle = 0 Then Exit Function

    If Not SourcePic Then
        If m_AutoDraw Then
            CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 2), PalSize
        End If
    End If

    If colorX < 0 Then colorX = PIC.BackColor

'centerX and centerY are passed as variant because the use of IsMissing is only possible with type variant
'if i used "optional centerx as long = -1" and checked for -1 as default value, rotating about this
'coordinate would be impossible...
    If IsMissing(centerX) Then centerX = PICW \ 2
    If IsMissing(centerY) Then centerY = PICH \ 2

    MCODE.QBLongColor = colorX
    ptMC = VarPtr(PicRotateMC(0))

    If SourcePic Then
        MCODE.OpCode = 2
    Else
        MCODE.OpCode = 0
    End If
    If AntiAlias Then
        MCODE.OpCode = MCODE.OpCode + 1
    End If
    
    CallWindowProc ptMC, ptrStruc, centerX, centerY, Angle + 360
    
    If m_AutoDraw Then
        ShowPalBGR 2
    End If

End Function

Private Function LoadMCodeFromString(asm_code_ARRAY() As Byte, ID_ASM As eASMBinary)
    asm_code_ARRAY = LoadResData(ID_ASM, "CUSTOM")
End Function

Private Function GeneratePalBGRs()
    GETDIBS PIC.Image
' Mem storage
'PalBGR for 32 bitcount
' BLUE     GREEN     RED       ALPHA
'(1,1,1,1)(2,1,1,1),(3,1,1,1),(4,1,1,1),,,(1,X,1,1),(2,X,1,1),(3,X,1,1),(4,X,1,1)
'(1,1,Y,1)(2,1,Y,1),(3,1,Y,1),(4,1,1,1),,,(1,X,Y,1),(2,X,Y,1),(3,X,Y,1),(4,X,Y,1)
'-----------------------------------------------------------------------------------
'(1,1,1,2)(2,1,1,2),(3,1,1,2),(4,1,1,2),,,(1,X,1,2),(2,X,1,2),(3,X,1,2),(4,X,1,2)
'(1,1,Y,2)(2,1,Y,1),(3,1,Y,1),(4,1,1,1),,,(1,X,Y,2),(2,X,Y,2),(3,X,Y,2),(4,X,Y,2)
'-----------------------------------------------------------------------------------
'(1,1,1,3)(2,1,1,3),(3,1,1,3),(4,1,1,3),,,(1,X,1,3),(2,X,1,3),(3,X,1,3),(4,X,1,3)
'(1,1,Y,3)(2,1,Y,3),(3,1,Y,3),(4,1,Y,3),,,(1,X,Y,3),(2,X,Y,3),(3,X,Y,3),(4,X,Y,3)
'-----------------------------------------------------------------------------------
' Save 2 copies of palette
    PalSize = 4 * PICW * PICH    ' Bytes
    CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 1), PalSize
    CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 1), PalSize
End Function

Private Function GETDIBS(ByVal PICIM As Long)
Dim NewDC As Long, OldH As Long
Dim BytesPerScanLine As Long, PadBytesPerScanLine As Long
Dim bmp As BITMAP

    On Error GoTo DIBError

'Get info on picture loaded into PIC
    GetObjectAPI PICIM, Len(bmp), bmp

    NewDC = CreateCompatibleDC(0&)
    OldH = SelectObject(NewDC, PICIM)

' Set up bm struc for GetDIBits & StretchDIBits
    With bm.bmiH
        .biSize = 40
        .biwidth = bmp.bmWidth
        .biheight = bmp.bmHeight
        .biPlanes = 1
        .biBitCount = 32    ' Sets up BGRA pixels
        .biCompression = 0
        BytesPerScanLine = ((((.biwidth * .biBitCount) + 31) \ 32) * 4)
        PadBytesPerScanLine = _
        BytesPerScanLine - (((.biwidth * .biBitCount) + 7) \ 8)
        .biSizeImage = BytesPerScanLine * Abs(.biheight)
    End With

' Set PalBGR to receive color bytes
    ReDim PalBGR(1 To 4, 1 To PICW, 1 To PICH, 1 To 3) As Byte
' Load color bytes to 1st third of PalBGR
    GetDIBits NewDC, PICIM, 0, PICH, PalBGR(1, 1, 1, 1), bm, 1
' Clear mem
    SelectObject NewDC, OldH

    DeleteDC NewDC
    
    Exit Function
    
'==========
DIBError:
    MsgBox "Error in function GETDIBS", vbCritical, "clsASM"
    On Error GoTo 0
End Function

Private Function ShowPalBGR(N)
' Blit PalBGR(N) to PIC
' N= 1,2 or 3
    PIC.Picture = LoadPicture()
    PIC.Visible = True

    Dim PalBGRPtr As Long    ' Pointer to PalBGR(1,1,1,1)
    PalBGRPtr = VarPtr(PalBGR(1, 1, 1, N))

    bm.bmiH.biwidth = PICW
    bm.bmiH.biheight = PICH
    If StretchDIBits(ByVal PIC.hDC, _
        0, 0, _
        PICW, PICH, _
        0, 0, _
        PICW, PICH, _
        ByVal PalBGRPtr, bm, _
        1, vbSrcCopy) = 0 Then
        Erase PalBGR
        MsgBox "Blit Error", vbCritical, "clsASM"
        End
    End If

    PIC.Refresh
    
End Function

Private Function SysBPPok(PIC As PictureBox) As Boolean
Dim SysBPP As Long
    SysBPP = GetDeviceCaps(PIC.hDC, BITSPIXEL)    'HORZRES = 8, VERTRES = 10, BITSPIXEL = 12
    If SysBPP <> 16 And SysBPP <> 32 Then    '16, 32 (24-bit BGR)
        SysBPPok = False
    Else
        SysBPPok = True
    End If
End Function

Public Function ResetPic()
    If NoPicAssigned Then Exit Function
    CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 1), PalSize
    CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 1), PalSize
    ShowPalBGR 1
End Function

Private Sub Class_Initialize()
    m_AutoDraw = True
'load all those mutthas...
    LoadMCodeFromString PicRotateMC, PICROTATE
End Sub

Private Sub Class_Terminate()
    Erase PalBGR
    Erase PicRotateMC
    Set PIC = Nothing
End Sub
