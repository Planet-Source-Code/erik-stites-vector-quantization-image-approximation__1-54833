VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Vector Quantization"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Tables"
      Height          =   555
      Left            =   45
      TabIndex        =   8
      Top             =   5625
      Width           =   1905
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "Adjust Tables"
      Height          =   555
      Left            =   2025
      TabIndex        =   6
      Top             =   5625
      Width           =   1905
   End
   Begin VB.PictureBox pbxRange 
      BackColor       =   &H00808080&
      Height          =   3900
      Left            =   7965
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   5
      Top             =   45
      Width           =   3900
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   510
      Left            =   2025
      TabIndex        =   4
      Top             =   4005
      Width           =   1905
   End
   Begin VB.PictureBox pbxTest 
      BackColor       =   &H00808000&
      Height          =   3900
      Left            =   4005
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      Top             =   45
      Width           =   3900
   End
   Begin VB.CommandButton cmdQuant 
      Caption         =   "Quantize"
      Height          =   510
      Left            =   45
      TabIndex        =   2
      Top             =   4005
      Width           =   1905
   End
   Begin VB.PictureBox pbxOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   3900
      Left            =   4005
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   3960
      Width           =   3900
   End
   Begin VB.PictureBox pbxVQ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   3900
      Left            =   45
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   45
      Width           =   3900
   End
   Begin VB.Label Label1 
      Caption         =   "Iterations"
      Height          =   240
      Left            =   2070
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "[status]"
      Height          =   285
      Left            =   2070
      TabIndex        =   7
      Top             =   5310
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Const REDCHANNEL As Integer = 2
Const GREENCHANNEL As Integer = 1
Const BLUECHANNEL As Integer = 0

Private Type RGBCode
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'Display handling
Const DIB_RGB_COLORS As Long = 0
Dim BMPInfo As BITMAPINFO
Dim VECInfo As BITMAPINFO
Dim Data() As Byte 'Original
Dim VecPairs() As Byte 'Approximated Output

'Iterations count
Dim Iter As Integer
Dim Steps As Integer

'Table
Const MAXIND As Integer = 127 '=(NumColor -1) Max is 255 unless you modify the steps for displaying the table
Dim VecTbl(0 To 2, 0 To MAXIND) As Integer 'Index of colors
Dim AvgTbl(0 To 2, 0 To MAXIND) As Long 'Summation of pixels for each index
Dim CntTbl(0 To MAXIND) As Long 'Number of pixels belonging to each index

'Finds and proportional distance, not the actual distance in the color cube
'Therefore, does not require exponents or square root as in: D=SQR(x^2+y^2+z^2)
Private Function FindNearest(x As Integer, y As Integer, z As Integer) As Integer
    Dim i As Integer
    Dim D As Integer
    Dim Shortest As Integer
    Dim Ind As Integer
    
    Shortest = 765 'Maximum distance
    
    For i = 0 To MAXIND
    
        D = Abs(x - VecTbl(0, i)) + Abs(y - VecTbl(1, i)) + Abs(z - VecTbl(2, i))
        
        If D < Shortest Then
            Shortest = D
            Ind = i
        End If
    
    Next
    
    FindNearest = Ind
    
End Function

Private Sub cmdAdjust_Click()
    Dim i As Integer
    
    For i = 0 To MAXIND
        
        'Resamples the index based on the input values
        If CntTbl(i) <> 0 Then
            VecTbl(0, i) = AvgTbl(0, i) \ CntTbl(i)

            VecTbl(1, i) = AvgTbl(1, i) \ CntTbl(i)

            VecTbl(2, i) = AvgTbl(2, i) \ CntTbl(i)
        End If
        
        AvgTbl(0, i) = 0
        AvgTbl(1, i) = 0
        AvgTbl(2, i) = 0
        CntTbl(i) = 0
        
        pbxRange.Line (i * Steps, 0)-(i * Steps + Steps - 1, 10), RGB(VecTbl(0, i), VecTbl(0, i), VecTbl(0, i)), BF
        pbxRange.Line (i * Steps, 11)-(i * Steps + Steps - 1, 20), RGB(VecTbl(1, i), VecTbl(1, i), VecTbl(1, i)), BF
        pbxRange.Line (i * Steps, 21)-(i * Steps + Steps - 1, 30), RGB(VecTbl(2, i), VecTbl(2, i), VecTbl(2, i)), BF
        pbxRange.Line (i * Steps, 31)-(i * Steps + Steps - 1, 40), RGB(VecTbl(0, i), VecTbl(1, i), VecTbl(2, i)), BF
    Next
    
    Iter = Iter + 1
    
    lblStatus.Caption = Iter
End Sub

Private Sub cmdPaste_Click()
    If Clipboard.GetData(vbCFBitmap) Then
        pbxVQ.Picture = Clipboard.GetData(vbCFBitmap)
    End If
End Sub

Private Sub cmdQuant_Click()
    Dim x As Integer, y As Integer
    Dim Col As Integer
    Dim Col1 As Integer, Col2 As Integer, Col3 As Integer
    Dim R As Integer, G As Integer, B As Integer
    
    GetDIBits pbxVQ.hdc, pbxVQ.Image.Handle, 0, BMPInfo.bmiHeader.biHeight, Data(0, 0, 0), BMPInfo, DIB_RGB_COLORS
    
    For x = 0 To 255
    
        For y = 0 To 255
            
            'Get the colors and find nearest index
            Col1 = Data(REDCHANNEL, x, y)
            Col2 = Data(GREENCHANNEL, x, y)
            Col3 = Data(BLUECHANNEL, x, y)
            Col = FindNearest(Col1, Col2, Col3)
            'Used for adjustments
            AvgTbl(0, Col) = AvgTbl(0, Col) + Col1
            AvgTbl(1, Col) = AvgTbl(1, Col) + Col2
            AvgTbl(2, Col) = AvgTbl(2, Col) + Col3
            CntTbl(Col) = CntTbl(Col) + 1
            'Generate output
            VecPairs(REDCHANNEL, x, y) = VecTbl(0, Col)
            VecPairs(GREENCHANNEL, x, y) = VecTbl(1, Col)
            VecPairs(BLUECHANNEL, x, y) = VecTbl(2, Col)
            'Display error
            R = 128 + (Col1 - VecTbl(0, Col))
            G = 128 + (Col2 - VecTbl(1, Col))
            B = 128 + (Col3 - VecTbl(2, Col))
            
            If R < 0 Then R = 0
            If G < 0 Then G = 0
            If B < 0 Then B = 0
            
            
            'Error map
            SetPixelV pbxTest.hdc, x, 255 - y, RGB(B, G, R)
            
        Next
        
    Next
    
    SetDIBits pbxOut.hdc, pbxOut.Image, 0, VECInfo.bmiHeader.biHeight, VecPairs(0, 0, 0), VECInfo, DIB_RGB_COLORS
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    Dim colorval As Long
    
    Randomize Rnd * Timer
    
    For i = 0 To MAXIND
        
        AvgTbl(0, i) = 0
        AvgTbl(1, i) = 0
        AvgTbl(2, i) = 0
        CntTbl(i) = 0
        
        'sample a random point on the original or each index in table
        colorval = pbxVQ.Point(Rnd * pbxVQ.ScaleWidth - 1, Rnd * pbxVQ.ScaleHeight - 1)
        
        VecTbl(0, i) = colorval And 255
        VecTbl(1, i) = (colorval \ 256) And 255
        VecTbl(2, i) = (colorval \ 65536) And 255
        
        'Display table
        pbxRange.Line (i * Steps, 0)-(i * Steps + Steps - 1, 10), RGB(VecTbl(0, i), VecTbl(0, i), VecTbl(0, i)), BF
        pbxRange.Line (i * Steps, 11)-(i * Steps + Steps - 1, 20), RGB(VecTbl(1, i), VecTbl(1, i), VecTbl(1, i)), BF
        pbxRange.Line (i * Steps, 21)-(i * Steps + Steps - 1, 30), RGB(VecTbl(2, i), VecTbl(2, i), VecTbl(2, i)), BF
        pbxRange.Line (i * Steps, 31)-(i * Steps + Steps - 1, 40), RGB(VecTbl(0, i), VecTbl(1, i), VecTbl(2, i)), BF
        
    Next
    
    Iter = 0
    lblStatus.Caption = Iter
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'Used for displaying color table data
    Steps = 256 \ (MAXIND + 1)
    
    With BMPInfo.bmiHeader
        .biWidth = pbxVQ.ScaleWidth
        .biHeight = pbxVQ.ScaleHeight
        .biSizeImage = .biWidth * .biHeight
        .biPlanes = 1
        .biBitCount = 32
        .biSize = 40
        .biCompression = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    
    With VECInfo.bmiHeader
        .biWidth = pbxVQ.ScaleWidth
        .biHeight = pbxVQ.ScaleHeight
        .biSizeImage = .biWidth * .biHeight
        .biPlanes = 1
        .biBitCount = 32
        .biSize = 40
        .biCompression = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    
    cmdReset_Click
    
    ReDim Data(0 To 3, 0 To BMPInfo.bmiHeader.biWidth - 1, 0 To BMPInfo.bmiHeader.biHeight - 1)
    ReDim VecPairs(0 To 3, 0 To VECInfo.bmiHeader.biWidth - 1, 0 To VECInfo.bmiHeader.biHeight - 1)
    
End Sub

Private Sub pbxVQ_Click()
    
    'Test input
    pbxVQ.Print "Vector Quantization"

End Sub
