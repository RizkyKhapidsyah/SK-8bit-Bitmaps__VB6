VERSION 5.00
Begin VB.Form frm8BitBitmaps 
   AutoRedraw      =   -1  'True
   Caption         =   "Using and manipulating palettes"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBright 
      Caption         =   "&Brightness"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdRipple 
      Caption         =   "&Ripple"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtRipple 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtBright 
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "Green it"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdBlue 
      Caption         =   "Blue it"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Red It"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdGray 
      Caption         =   "Gray it"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
End
Attribute VB_Name = "frm8BitBitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Image Processing with 8-bit bitmaps
'

Option Explicit

Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTCOLOR As Long = &H0
Const LR_COLOR As Long = &H2
'****************************************
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Dim DC As Long
Dim Bitmaphandle As Long
Dim bm As BITMAP
'Original colors
Dim OriginalTable(1 To 256) As RGBQUAD
'color tables
Dim GrayTable(1 To 256) As RGBQUAD
Dim RedTable(1 To 256) As RGBQUAD
Dim BlueTable(1 To 256) As RGBQUAD
Dim GreenTable(1 To 256) As RGBQUAD
Dim InvertTable(1 To 256) As RGBQUAD

'Dimensions
Const BitmapWidth As Long = 200
Const BitmapHeight As Long = 200

Private Sub cmdBlue_Click()

SetDIBColorTable DC, 0, 256, BlueTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdBright_Click()
Dim TempValue As Long
Dim I As Long
Dim BrightColorTable(1 To 256) As RGBQUAD
'Brightness Table
Dim BrightTable(0 To 255) As Byte

'Build brightness lookup table
For I = 0 To 255
    TempValue = I * Val(txtBright.Text)
    
    If TempValue > 255 Then
        BrightTable(I) = 255
    Else
        BrightTable(I) = TempValue
    End If
Next I

'Build the actual color table
For I = 1 To 256
    
    BrightColorTable(I).rgbBlue = BrightTable(OriginalTable(I).rgbBlue)
    BrightColorTable(I).rgbRed = BrightTable(OriginalTable(I).rgbRed)
    BrightColorTable(I).rgbGreen = BrightTable(OriginalTable(I).rgbGreen)

Next I

SetDIBColorTable DC, 0, 256, BrightColorTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdGray_Click()

SetDIBColorTable DC, 0, 256, GrayTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdGreen_Click()

SetDIBColorTable DC, 0, 256, GreenTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdInvert_Click()

SetDIBColorTable DC, 0, 256, InvertTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdRed_Click()

SetDIBColorTable DC, 0, 256, RedTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdRestore_Click()

SetDIBColorTable DC, 0, 256, OriginalTable(1)

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub cmdRipple_Click()
Dim ByteArray() As Byte
Dim I As Long, J As Long
Dim TempValue As Long
Dim RippleTable() As Byte
Dim OriginalBits() As Byte

ReDim OriginalBits(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

GetBitmapBits Bitmaphandle, bm.bmWidthBytes * bm.bmHeight, OriginalBits(1, 1)


'Dimension the ripple lookup table
ReDim RippleTable(1 To BitmapWidth)

'Build ripple table
For I = 1 To BitmapWidth
    TempValue = I + Sin(I / 5) * Val(txtRipple.Text)
    If TempValue > BitmapWidth Then
        RippleTable(I) = BitmapWidth
    ElseIf TempValue < 1 Then
        RippleTable(I) = 1
    Else
        RippleTable(I) = TempValue
    End If
    
Next I

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = OriginalBits(I, RippleTable(J))
        
    Next J
Next I

SetBitmapBits Bitmaphandle, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy
Me.Refresh

'Reset the bits
SetBitmapBits Bitmaphandle, bm.bmWidthBytes * bm.bmHeight, OriginalBits(1, 1)

End Sub

Private Sub Form_Load()

DC = GenerateDC(App.Path & "\bitmap1.bmp", Bitmaphandle)

'Check if the bitmap is 8-bit
GetObjectAPI Bitmaphandle, Len(bm), bm

If bm.bmBitsPixel <> 8 Then 'not a usable format
    MsgBox "Must be an 8-bit bitmap"
    Unload Me
    Exit Sub
End If


'Save the original color table
GetDIBColorTable DC, 0, 256, OriginalTable(1)

'Create the gray color table, based on the original table
CreateColorTables

'draw the picture
BitBlt Me.hDC, 0, 0, BitmapWidth, BitmapHeight, DC, 0, 0, vbSrcCopy

Me.Refresh

End Sub

Private Sub CreateColorTables()
Dim I As Long
Dim TempValue As Long

For I = LBound(GrayTable) To UBound(GrayTable)
        
    'Create Gray Color table
    'Add the values together
    TempValue = OriginalTable(I).rgbBlue
    TempValue = TempValue + OriginalTable(I).rgbGreen
    TempValue = TempValue + OriginalTable(I).rgbRed
    
    'Get the medium value
    TempValue = TempValue / 3
    
    'Set the color in the gray table
    GrayTable(I).rgbBlue = TempValue
    GrayTable(I).rgbGreen = TempValue
    GrayTable(I).rgbRed = TempValue
        
    'Create the rest of the color tables
    'Red Table
    RedTable(I).rgbBlue = 0
    RedTable(I).rgbGreen = 0
    RedTable(I).rgbRed = OriginalTable(I).rgbRed
    
    'Green Table
    GreenTable(I).rgbBlue = 0
    GreenTable(I).rgbRed = 0
    GreenTable(I).rgbGreen = OriginalTable(I).rgbGreen
    
    'Blue table
    BlueTable(I).rgbBlue = OriginalTable(I).rgbBlue
    BlueTable(I).rgbGreen = 0
    BlueTable(I).rgbRed = 0
    
    'invert table
    InvertTable(I).rgbBlue = 255 - OriginalTable(I).rgbBlue
    InvertTable(I).rgbGreen = 255 - OriginalTable(I).rgbGreen
    InvertTable(I).rgbRed = 255 - OriginalTable(I).rgbRed

Next I
    
End Sub

'IN: FileName: The file name of the graphics
'    BitmapHandle: The receiver of the loaded bitmap handle
'OUT: The Generated DC
Public Function GenerateDC(FileName As String, ByRef Bitmaphandle As Long) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 1
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 2
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context and handle
Bitmaphandle = hBitmap
GenerateDC = DC

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
'Clean Up
DeleteGeneratedDC DC
DeleteObject Bitmaphandle

End Sub
