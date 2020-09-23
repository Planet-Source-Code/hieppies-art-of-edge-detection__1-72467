VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Art of Edge Detection by Hieppies"
   ClientHeight    =   8985
   ClientLeft      =   2775
   ClientTop       =   1005
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   943
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   6360
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   19
      Top             =   6360
      Width           =   1920
   End
   Begin VB.PictureBox Pic9 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   4320
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   17
      Top             =   6360
      Width           =   1920
   End
   Begin VB.PictureBox Pic8 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   2280
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   15
      Top             =   6360
      Width           =   1920
   End
   Begin VB.PictureBox Pic7 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   4320
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   12
      Top             =   3360
      Width           =   1920
   End
   Begin VB.PictureBox Pic6 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   2280
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   11
      Top             =   3360
      Width           =   1920
   End
   Begin VB.PictureBox Pic5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   8400
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   5
      Top             =   360
      Width           =   1920
   End
   Begin VB.PictureBox Pic4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   6360
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   4
      Top             =   360
      Width           =   1920
   End
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   4320
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   3
      Top             =   360
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      Height          =   8415
      Left            =   10560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   0
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   2280
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   1
      Top             =   360
      Width           =   1920
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   0
      Top             =   360
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canny Edge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   6360
      TabIndex        =   20
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canny Grayscale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4320
      TabIndex        =   18
      Top             =   6120
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canny Gaussian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2280
      TabIndex        =   16
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   144
      X2              =   696
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gaussian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4320
      TabIndex        =   14
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mean"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sobel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   8400
      TabIndex        =   10
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6360
      TabIndex        =   9
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prewit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Isotropic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   825
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   696
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line2 
      X1              =   696
      X2              =   696
      Y1              =   0
      Y2              =   608
   End
   Begin VB.Line Line1 
      X1              =   144
      X2              =   144
      Y1              =   0
      Y2              =   616
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuload 
         Caption         =   "Load Gambar..."
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Keluar"
      End
   End
   Begin VB.Menu mnufilter 
      Caption         =   "Filter"
      Begin VB.Menu mnuoperator 
         Caption         =   "Deteksi Tepi"
         Begin VB.Menu mnuisothropic 
            Caption         =   "&Metode Isotropic"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuprewit 
            Caption         =   "&Metode Prewit"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnurobert 
            Caption         =   "&Metode Robert"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnusobel 
            Caption         =   "&Metode Sobel"
            Shortcut        =   {F4}
         End
         Begin VB.Menu vbsep 
            Caption         =   "-"
         End
         Begin VB.Menu mnucanny 
            Caption         =   "&Metode Canny"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu vbsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnublur 
         Caption         =   "Blur"
         Begin VB.Menu mnusmooth 
            Caption         =   "Mean"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnugaussian 
            Caption         =   "Gaussian"
            Shortcut        =   {F7}
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim R As Integer, G As Integer, B As Integer
Dim Itensity As Long, GradX As Long, GradY As Long, Grad As Long
Dim PixelValue As Long

Sub DecTORGB(ByVal Col As Long, R As Integer, G As Integer, B As Integer)
    R = Col Mod 256
    G = ((Col - R) Mod 65536) / 256
    B = (Col - R - G) / 65536
    If R < 0 Then R = 0: If R >= 255 Then R = 255
    If G < 0 Then G = 0: If G >= 255 Then G = 255
    If B < 0 Then B = 0: If B >= 255 Then B = 255
End Sub


Private Sub mnugaussian_Click()
On Error GoTo ErrHandle
Const MaxData = 255
Const DataGranularity = 1
Const Delta = DataGranularity / (2 * MaxData)

Dim X As Integer, Y As Integer
Dim fBias As Integer, fScaleFactor As Single
Dim fRadius As Single, Sum As Single, RR2 As Single
Dim MaxGaussianSize As Integer, GaussianSize As Integer
Dim GaussianKernel() As Double, Kernel() As Single
Dim KernelSize As Long
Dim fKernel As Long, fWidth As Long, fHeight As Long, fCount As Long
Dim SF As Single, Rad As Single, W As Single, C As Single
Dim KWH As Long, KWL As Long, KHH As Long, KHL As Long

MaxGaussianSize = 50
fBias = 0: fScaleFactor = 1: fCount = 1
fRadius = InputBox("Masukkan Angka Tetha antara [0.5 - 2] : ", "Gaussian Blur Radius")
If fRadius < 0 Then fRadius = 0
If fRadius > 10 Then MsgBox "Masukkan angka antara [0.5 - 2]", vbInformation, "Angka Kelebihan": Exit Sub
ReDim GaussianKernel(-MaxGaussianSize To MaxGaussianSize, -MaxGaussianSize To MaxGaussianSize)


Sum = 0
RR2 = -2 * fRadius * fRadius

For Y = -MaxGaussianSize To MaxGaussianSize
    For X = -MaxGaussianSize To MaxGaussianSize
        GaussianKernel(Y, X) = Exp((X * X + Y * Y) / RR2)
        Sum = Sum + GaussianKernel(Y, X)
    Next X
Next Y

For Y = -MaxGaussianSize To MaxGaussianSize
    For X = -MaxGaussianSize To MaxGaussianSize
        GaussianKernel(Y, X) = GaussianKernel(Y, X) / Sum
    Next X
Next Y

Sum = 0
GaussianSize = MaxGaussianSize
Do While (GaussianSize > 1) And (Sum < Delta)
    Sum = Sum + 4 * GaussianKernel(0, GaussianSize)
    GaussianSize = GaussianSize - 1
Loop

For Y = -GaussianSize To GaussianSize
    For X = -GaussianSize To GaussianSize
        Sum = Sum + GaussianKernel(Y, X)
    Next X
Next Y

For Y = -GaussianSize To GaussianSize
    For X = -GaussianSize To GaussianSize
        GaussianKernel(Y, X) = GaussianKernel(Y, X) / Sum
    Next X
Next Y

KernelSize = (2 * GaussianSize) + 1
SF = 0: Rad = 1
Dim RKernel(99999) As Single, cnt As Integer
cnt = 0
For Y = -GaussianSize To GaussianSize
    C = 0
    For X = -GaussianSize To GaussianSize
        W = Round((1 / GaussianKernel(GaussianSize, GaussianSize)) * GaussianKernel(Y, X))
        RKernel(cnt) = W
        SF = SF + W
        C = C + 1
        cnt = cnt + 1
    Next X
    Rad = Rad + 1
Next Y
fScaleFactor = SF

Me.Cls
ReDim Kernel(1, Rad, C)
cnt = 0
Dim tmps As String
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text + "Hasil Kalkulasi Kernel untuk Gaussian Blur >>" & vbCrLf
Text1.SelStart = Len(Text1.Text)
ReDim Kernel(0 To 1, -GaussianSize To GaussianSize, -GaussianSize To GaussianSize)
For I = -GaussianSize To GaussianSize
    For J = -GaussianSize To GaussianSize
        Kernel(1, I, J) = RKernel(cnt)
        tmps = tmps & Format(Kernel(1, I, J), "000") & " "
        cnt = cnt + 1
    Next J
    Text1.Text = Text1.Text + tmps & vbCrLf
    tmps = ""
Next I

Text1.Text = Text1.Text + "Menggunakan Theta = " & fRadius & " dan Ukuran Kernel = " & KernelSize & " x " & KernelSize & " Pixel" & vbCrLf
Text1.Text = Text1.Text + "Mengunakan Fungsi gambar 2D" & vbCrLf
Text1.SelStart = Len(Text1.Text)

Pic7.Cls
Dim tmpIntR As Double, tmpIntG As Double, tmpIntB As Double
Dim CTotal As Single
Dim CDataR As Single, CDataG As Single, CDataB As Single

CTotal = (((GaussianSize * GaussianSize) + GaussianSize) * 4) + 1
Dim tmpC As Long
Dim CountClr As Long

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width
        For J = -GaussianSize To GaussianSize
            For I = -GaussianSize To GaussianSize
                 PixelValue = GetPixel(Pic1.hdc, X + I, Y + J)
                 DecTORGB PixelValue, R, G, B
                 CDataR = CDataR + (Kernel(1, J, I) * R)
                 CDataG = CDataG + (Kernel(1, J, I) * G)
                 CDataB = CDataB + (Kernel(1, J, I) * B)
                 If mX < C Then mX = mX + 1
            Next I
        Next J
        CDataR = CDataR / (fScaleFactor + fBias)
        CDataG = CDataG / (fScaleFactor + fBias)
        CDataB = CDataB / (fScaleFactor + fBias)
        SetPixel Pic7.hdc, X, Y, RGB(CDataR, CDataG, CDataB)
        Pic7.Refresh
        DoEvents
    Next X
    Pic7.Refresh
Next Y
Pic7.Refresh
Exit Sub
ErrHandle:
Exit Sub
End Sub

Private Sub mnuisothropic_Click()
Dim Op_X(-1 To 1, -1 To 1) As Integer, Op_Y(-1 To 1, -1 To 1) As Integer
Dim X As Integer, Y As Integer, I As Integer, J As Integer
Pic2.Cls
Grad = 0
Op_X(-1, -1) = -1: Op_X(0, -1) = -(Sqr(2)): Op_X(1, -1) = -1
Op_X(-1, 0) = 0: Op_X(0, 0) = 0: Op_X(1, 0) = 0
Op_X(-1, 1) = 1: Op_X(0, 1) = (Sqr(2)): Op_X(1, 1) = 1

Op_Y(-1, -1) = -1: Op_Y(0, -1) = 0: Op_Y(1, -1) = 1
Op_Y(-1, 0) = -(Sqr(2)): Op_Y(0, 0) = 0: Op_Y(1, 0) = (Sqr(2))
Op_Y(-1, 1) = -1: Op_Y(0, 1) = 0: Op_Y(1, 1) = 1

Dim tmps As String
tmps = "Menggunakan Kernel untuk Isotropic" & vbCrLf & "Horisontal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_X(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
tmps = tmps & "Vertikal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_Y(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
Text1.Text = Text1.Text & tmps & vbCrLf
DoEvents
For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width - 1
        GradX = 0: GradY = 0: Grad = 0
        If X = 0 Or Y = 0 Or X = Pic1.Width - 1 Or Y = Pic1.Height - 1 Then
            Grad = 0
        Else
            For I = -1 To 1
                For J = -1 To 1
                PixelValue = GetPixel(Pic1.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
                DecTORGB PixelValue, R, G, B 'fungsi proses mendapatkan nilai RGB
                Itensity = (R + G + B) / 3 'Itensitas / B & W
                GradX = GradX + (Itensity * Op_X(I, J))
                GradY = GradY + (Itensity * Op_Y(I, J))
                Next J
            Next I
            Grad = Round(Sqr(Abs(GradX * GradX) + Abs(GradY * GradY)))
        End If
        If Grad <= 0 Then Grad = 0: If Grad >= 255 Then Grad = 255
        SetPixel Pic2.hdc, X, Y, RGB(Grad, Grad, Grad)
        Pic2.Refresh
    Next X
    Pic2.Refresh
Next Y

End Sub

Private Sub mnuload_Click()
Dim Token As Long
CM.Filter = "Image|*.bmp;*.jpg"
CM.ShowOpen
If CM.FileName <> "" Then
Token = InitGDIPlus
Pic1.Picture = LoadPictureGDIPlus(CM.FileName, Pic1.Width, Pic1.Height, , False)
FreeGDIPlus Token
End If
Text1.Text = ""
End Sub

Private Sub mnuprewit_Click()
Dim Op_X(-1 To 1, -1 To 1) As Integer, Op_Y(-1 To 1, -1 To 1) As Integer
Dim X As Integer, Y As Integer, I As Integer, J As Integer
Pic3.Cls
Grad = 0
Op_X(-1, -1) = -1: Op_X(0, -1) = -1: Op_X(1, -1) = -1
Op_X(-1, 0) = 0: Op_X(0, 0) = 0: Op_X(1, 0) = 0
Op_X(-1, 1) = 1: Op_X(0, 1) = 1: Op_X(1, 1) = 1

Op_Y(-1, -1) = 1: Op_Y(0, -1) = 0: Op_Y(1, -1) = -1
Op_Y(-1, 0) = 0: Op_Y(0, 0) = 0: Op_Y(1, 0) = 0
Op_Y(-1, 1) = 1: Op_Y(0, 1) = 0: Op_Y(1, 1) = -1

Dim tmps As String
tmps = "Menggunakan Kernel untuk Prewit" & vbCrLf & "Horisontal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_X(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
tmps = tmps & "Vertikal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_Y(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
Text1.Text = Text1.Text & tmps & vbCrLf
DoEvents

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width - 1
        GradX = 0: GradY = 0: Grad = 0
        If X = 0 Or Y = 0 Or X = Pic1.Width - 1 Or Y = Pic1.Height - 1 Then
            Grad = 0
        Else
            For I = -1 To 1
                For J = -1 To 1
                PixelValue = GetPixel(Pic1.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
                DecTORGB PixelValue, R, G, B 'fungsi proses mendapatkan nilai RGB
                Itensity = (R + G + B) / 3 'Itensitas / B & W
                GradX = GradX + (Itensity * Op_X(I, J))
                GradY = GradY + (Itensity * Op_Y(I, J))
                Next J
            Next I
            Grad = Round(Sqr(Abs(GradX * GradX) + Abs(GradY * GradY)))
        End If
        If Grad <= 0 Then Grad = 0: If Grad >= 255 Then Grad = 255
        SetPixel Pic3.hdc, X, Y, RGB(Grad, Grad, Grad)
        Pic3.Refresh
    Next X
    Pic3.Refresh
Next Y

End Sub

Private Sub mnurobert_Click()
Dim Op_X(-1 To 0, -1 To 0) As Integer, Op_Y(-1 To 0, -1 To 0) As Integer
Dim X As Integer, Y As Integer, I As Integer, J As Integer
Pic4.Cls
Grad = 0
Op_X(-1, -1) = 1: Op_X(0, -1) = 0
Op_X(-1, 0) = 0: Op_X(0, 0) = -1
Op_Y(-1, -1) = 0: Op_Y(0, -1) = -1
Op_Y(-1, 0) = 1: Op_Y(0, 0) = 0

Dim tmps As String
tmps = "Menggunakan Kernel untuk Robert" & vbCrLf & "Horisontal >>" & vbCrLf
For X = -1 To 0
    For Y = -1 To 0
        tmps = tmps & Format(Op_X(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
tmps = tmps & "Vertikal >>" & vbCrLf
For X = -1 To 0
    For Y = -1 To 0
        tmps = tmps & Format(Op_Y(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
Text1.Text = Text1.Text & tmps & vbCrLf
DoEvents

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width - 1
        GradX = 0: GradY = 0: Grad = 0
        If X = 0 Or Y = 0 Or X = Pic1.Width - 1 Or Y = Pic1.Height - 1 Then
            Grad = 0
        Else
            For I = -1 To 0
                For J = -1 To 0
                PixelValue = GetPixel(Pic1.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
                DecTORGB PixelValue, R, G, B 'fungsi proses mendapatkan nilai RGB
                Itensity = (R + G + B) / 3 'Itensitas / B & W
                GradX = GradX + (Itensity * Op_X(I, J))
                GradY = GradY + (Itensity * Op_Y(I, J))
                Next J
            Next I
            Grad = Round(Sqr(Abs(GradX * GradX) + Abs(GradY * GradY)))
        End If
        If Grad <= 0 Then Grad = 0: If Grad >= 255 Then Grad = 255
        SetPixel Pic4.hdc, X, Y, RGB(Grad, Grad, Grad)
        Pic4.Refresh
    Next X
    Pic4.Refresh
Next Y

End Sub

Private Sub mnusmooth_Click()
On Error GoTo ErrHandle
Dim X As Integer, Y As Integer
Dim mR As Integer, mG As Integer, mB As Integer
Dim mR1 As Integer, mR2 As Integer, mR3 As Integer, mR4 As Integer, mR5 As Integer
Dim mG1 As Integer, mG2 As Integer, mG3 As Integer, mG4 As Integer, mG5 As Integer
Dim mB1 As Integer, mB2 As Integer, mB3 As Integer, mB4 As Integer, mB5 As Integer
Dim mPixel1 As Long, mPixel2 As Long, mPixel3 As Long, mPixel4 As Long, mPixel5 As Long
Dim inpNum As Integer
Pic6.Cls
inpNum = 4

For Y = 1 To Pic1.Height - 2
    For X = 1 To Pic1.Width - 2
        mPixel1 = GetPixel(Pic1.hdc, X, Y)
        mPixel2 = GetPixel(Pic1.hdc, X + 1, Y)
        mPixel2 = GetPixel(Pic1.hdc, X - 1, Y)
        mPixel3 = GetPixel(Pic1.hdc, X, Y + 1)
        mPixel4 = GetPixel(Pic1.hdc, X, Y - 1)
        DecTORGB mPixel1, mR1, mG1, mB1
        DecTORGB mPixel2, mR2, mG2, mB2
        DecTORGB mPixel3, mR3, mG3, mB3
        DecTORGB mPixel4, mR4, mG4, mB4
        DecTORGB mPixel5, mR5, mG5, mB5
        If mR1 < 0 Then mR1 = 0
        mR = mR1 + mR2 + mR3 + mR4 + mR5
        mR = mR / inpNum
        mG = mG1 + mG2 + mG3 + mG4 + mG5
        mG = mG / inpNum
        mB = mB1 + mB2 + mB3 + mB4 + mB5
        mB = mB / inpNum
        
        SetPixel Pic6.hdc, X, Y, RGB(mR, mG, mB)
    Next X
    Pic6.Refresh
    DoEvents
Next Y
Exit Sub
ErrHandle:
'MsgBox "Bukan Angka!.", vbCritical, "Error"
Exit Sub
End Sub

Private Sub mnusobel_Click()
Dim Op_X(-1 To 1, -1 To 1) As Integer, Op_Y(-1 To 1, -1 To 1) As Integer
Dim X As Integer, Y As Integer, I As Integer, J As Integer
Pic5.Cls
Grad = 0
Op_X(-1, -1) = -1: Op_X(0, -1) = -2: Op_X(1, -1) = -1
Op_X(-1, 0) = 0: Op_X(0, 0) = 0: Op_X(1, 0) = 0
Op_X(-1, 1) = 1: Op_X(0, 1) = 2: Op_X(1, 1) = 1

Op_Y(-1, -1) = -1: Op_Y(0, -1) = 0: Op_Y(1, -1) = 1
Op_Y(-1, 0) = -2: Op_Y(0, 0) = 0: Op_Y(1, 0) = 2
Op_Y(-1, 1) = -1: Op_Y(0, 1) = 0: Op_Y(1, 1) = 1

Dim tmps As String
tmps = "Menggunakan Kernel untuk Sobel" & vbCrLf & "Horisontal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_X(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
tmps = tmps & "Vertikal >>" & vbCrLf
For X = -1 To 1
    For Y = -1 To 1
        tmps = tmps & Format(Op_Y(X, Y), "0") & " "
    Next Y
    tmps = tmps & vbCrLf
Next X
Text1.Text = Text1.Text & tmps & vbCrLf
DoEvents

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width - 1
        GradX = 0: GradY = 0: Grad = 0
        If X = 0 Or Y = 0 Or X = Pic1.Width - 1 Or Y = Pic1.Height - 1 Then
            Grad = 0
        Else
            For I = -1 To 1
                For J = -1 To 1
                PixelValue = GetPixel(Pic1.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
                DecTORGB PixelValue, R, G, B 'fungsi proses mendapatkan nilai RGB
                Itensity = (R + G + B) / 3 'Itensitas / B & W
                GradX = GradX + (Itensity * Op_X(I, J))
                GradY = GradY + (Itensity * Op_Y(I, J))
                Next J
            Next I
            Grad = Round(Sqr(Abs(GradX * GradX) + Abs(GradY * GradY)))
        End If
        If Grad <= 0 Then Grad = 0: If Grad >= 255 Then Grad = 255
        SetPixel Pic5.hdc, X, Y, RGB(Grad, Grad, Grad)
        Pic5.Refresh
    Next X
    Pic5.Refresh
Next Y
End Sub

Private Sub mnucanny_Click()
On Error GoTo ErrHandle
Const MaxData = 255
Const DataGranularity = 1
Const Delta = DataGranularity / (2 * MaxData)

Dim Op_X(-1 To 1, -1 To 1) As Integer, Op_Y(-1 To 1, -1 To 1) As Integer

Op_X(-1, -1) = -1: Op_X(0, -1) = -2: Op_X(1, -1) = -1
Op_X(-1, 0) = 0: Op_X(0, 0) = 0: Op_X(1, 0) = 0
Op_X(-1, 1) = 1: Op_X(0, 1) = 2: Op_X(1, 1) = 1

Op_Y(-1, -1) = -1: Op_Y(0, -1) = 0: Op_Y(1, -1) = 1
Op_Y(-1, 0) = -2: Op_Y(0, 0) = 0: Op_Y(1, 0) = 2
Op_Y(-1, 1) = -1: Op_Y(0, 1) = 0: Op_Y(1, 1) = 1

Dim X As Integer, Y As Integer
Dim fBias As Integer, fScaleFactor As Single
Dim fRadius As Single, Sum As Single, RR2 As Single
Dim MaxGaussianSize As Integer, GaussianSize As Integer
Dim GaussianKernel() As Double, Kernel() As Single
Dim KernelSize As Long
Dim fKernel As Long, fWidth As Long, fHeight As Long, fCount As Long
Dim SF As Single, Rad As Single, W As Single, C As Single
Dim KWH As Long, KWL As Long, KHH As Long, KHL As Long

Pic8.Cls: Pic9.Cls: Pic10.Cls
MaxGaussianSize = 50
fBias = 0: fScaleFactor = 1: fCount = 1
fRadius = InputBox("Masukkan Angka Tetha antara [0.5 - 2] : ", "Gaussian Blur Radius")
If fRadius < 0 Then fRadius = 0
If fRadius > 10 Then MsgBox "Masukkan angka antara [0.5 - 2]", vbInformation, "Angka Kelebihan": Exit Sub
ReDim GaussianKernel(-MaxGaussianSize To MaxGaussianSize, -MaxGaussianSize To MaxGaussianSize)


Sum = 0
RR2 = -2 * fRadius * fRadius

For Y = -MaxGaussianSize To MaxGaussianSize
    For X = -MaxGaussianSize To MaxGaussianSize
        GaussianKernel(Y, X) = Exp((X * X + Y * Y) / RR2)
        Sum = Sum + GaussianKernel(Y, X)
    Next X
Next Y

For Y = -MaxGaussianSize To MaxGaussianSize
    For X = -MaxGaussianSize To MaxGaussianSize
        GaussianKernel(Y, X) = GaussianKernel(Y, X) / Sum
    Next X
Next Y

Sum = 0
GaussianSize = MaxGaussianSize
Do While (GaussianSize > 1) And (Sum < Delta)
    Sum = Sum + 4 * GaussianKernel(0, GaussianSize)
    GaussianSize = GaussianSize - 1
Loop

For Y = -GaussianSize To GaussianSize
    For X = -GaussianSize To GaussianSize
        Sum = Sum + GaussianKernel(Y, X)
    Next X
Next Y

For Y = -GaussianSize To GaussianSize
    For X = -GaussianSize To GaussianSize
        GaussianKernel(Y, X) = GaussianKernel(Y, X) / Sum
    Next X
Next Y

KernelSize = (2 * GaussianSize) + 1
SF = 0: Rad = 1
Dim RKernel(99999) As Single, cnt As Integer
cnt = 0
For Y = -GaussianSize To GaussianSize
    C = 0
    For X = -GaussianSize To GaussianSize
        W = Round((1 / GaussianKernel(GaussianSize, GaussianSize)) * GaussianKernel(Y, X))
        RKernel(cnt) = W
        SF = SF + W
        C = C + 1
        cnt = cnt + 1
    Next X
    Rad = Rad + 1
Next Y
fScaleFactor = SF

Me.Cls
ReDim Kernel(1, Rad, C)
cnt = 0
Dim tmps As String
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & "Mencari Kernel untuk Metode Canny" & vbCrLf
Text1.Text = Text1.Text & "Langkah 1. Konversi gambar ke Gaussian Blur..." & vbCrLf
Text1.Text = Text1.Text + "Hasil Kalkulasi Kernel dari Gaussian Blur >>" & vbCrLf
Text1.SelStart = Len(Text1.Text)
ReDim Kernel(0 To 1, -GaussianSize To GaussianSize, -GaussianSize To GaussianSize)
For I = -GaussianSize To GaussianSize
    For J = -GaussianSize To GaussianSize
        Kernel(1, I, J) = RKernel(cnt)
        tmps = tmps & Format(Kernel(1, I, J), "000") & " "
        cnt = cnt + 1
    Next J
    Text1.Text = Text1.Text + tmps & vbCrLf
    tmps = ""
Next I

Text1.Text = Text1.Text + "Menggunakan Theta = " & fRadius & " dan Ukuran Kernel = " & KernelSize & " x " & KernelSize & " Pixel" & vbCrLf
Text1.Text = Text1.Text + "Mengunakan Fungsi gambar 2D"
Text1.SelStart = Len(Text1.Text)

Dim tmpIntR As Double, tmpIntG As Double, tmpIntB As Double
Dim CTotal As Single
Dim CDataR As Single, CDataG As Single, CDataB As Single
Dim Intensity As Long, Grad As Single

CTotal = (((GaussianSize * GaussianSize) + GaussianSize) * 4) + 1
Dim tmpC As Long
Dim CountClr As Long

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width
        For J = -GaussianSize To GaussianSize
            For I = -GaussianSize To GaussianSize
                 PixelValue = GetPixel(Pic1.hdc, X + I, Y + J)
                 DecTORGB PixelValue, R, G, B
                 CDataR = CDataR + (Kernel(1, J, I) * R)
                 CDataG = CDataG + (Kernel(1, J, I) * G)
                 CDataB = CDataB + (Kernel(1, J, I) * B)
                 If mX < C Then mX = mX + 1
            Next I
        Next J
        CDataR = CDataR / (fScaleFactor + fBias)
        CDataG = CDataG / (fScaleFactor + fBias)
        CDataB = CDataB / (fScaleFactor + fBias)

        SetPixel Pic8.hdc, X, Y, RGB(CDataR, CDataG, CDataB)
        Pic8.Refresh
        DoEvents
    Next X
Next Y
Pic8.Refresh
Text1.Text = Text1.Text & vbCrLf & vbCrLf & "Langkah 2. Konversi gambar ke Grayscale..." & vbCrLf
Text1.SelStart = Len(Text1.Text)
Dim Greycolor As Integer, PixNum As Integer
'PixNum = InputBox("Masukkan Nilai Untuk Grayscale antara [0 - 255] :", "Input Grayscale")
PixNum = 90
If PixNum > 255 Then PixNum = 255
Text1.Text = Text1.Text & "Nilai Warna Grayscale = " & PixNum & vbCrLf
Text1.SelStart = Len(Text1.Text)
DoEvents
For Y = 0 To Pic8.Height
    For X = 0 To Pic8.Width
        PixelValue = GetPixel(Pic8.hdc, X, Y)
        DecTORGB PixelValue, R, G, B
        Greycolor = Greyscale(PixelValue, PixNum)
        SetPixel Pic9.hdc, X, Y, RGB(Greycolor, Greycolor, Greycolor)
        Pic9.Refresh
    Next X
    Pic9.Refresh
Next Y

Dim Itensity As Long, GradX As Long
Dim ThresholdMin As Integer, ThresholdMax As Integer
Dim sR As Long, sG As Long, sB As Long

Text1.Text = Text1.Text & vbCrLf & "Langkah 3. Filtering gambar mengunakan 2 Thresholds..." & vbCrLf
Text1.SelStart = Len(Text1.Text)
ThresholdMin = InputBox("Masukkan Nilai Min Threshold :", "Nilai Min Threshold")
If ThresholdMin > 255 Then ThresholdMin = 255
Text1.Text = Text1.Text & "Min Threshold = " & ThresholdMin & vbCrLf
Text1.SelStart = Len(Text1.Text)
ThresholdMax = InputBox("Masukkan Nilai Max Threshold :", "Nilai Max Threshold")
If ThresholdMax > 255 Then ThresholdMax = 255
Text1.Text = Text1.Text & "Max Threshold = " & ThresholdMax & vbCrLf
Text1.SelStart = Len(Text1.Text)
'fScaleFactor = GaussianSize * GaussianSize
DoEvents
'ThresholdMin = 0
'ThresholdMax = 0

For Y = 0 To Pic1.Height - 1
    For X = 0 To Pic1.Width - 1
        GradX = 0: GradY = 0: Grad = 0
        If X = 0 Or Y = 0 Or X = Pic1.Width - 1 Or Y = Pic1.Height - 1 Then
            Grad = 0
        Else
            For I = -1 To 1
                For J = -1 To 1
                PixelValue = GetPixel(Pic8.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
                DecTORGB PixelValue, R, G, B 'fungsi proses mendapatkan nilai RGB
                Itensity = (R + G + B) / 3 'Itensitas / B & W
                GradX = GradX + (Itensity * Op_X(I, J))
                GradY = GradY + (Itensity * Op_Y(I, J))
                Next J
            Next I
            Grad = Round(Sqr(Abs(GradX * GradX) + Abs(GradY * GradY)))
        End If
        If Grad >= ThresholdMin And Grad <= ThresholdMax Then
            Grad = Grad - Abs(((ThresholdMax + ThresholdMin) / KernelSize))
            If Grad <= ThresholdMax Then Grad = 0: If Grad >= ThresholdMax Then Grad = 255
            SetPixel Pic10.hdc, X, Y, RGB(Grad, Grad, Grad)
        Else
            If Grad <= 0 Then Grad = 0: If Grad >= 255 Then Grad = 255
            SetPixel Pic10.hdc, X, Y, RGB(Grad, Grad, Grad)
        End If
        Pic10.Refresh
    Next X
    Pic10.Refresh
Next Y
Exit Sub
ErrHandle:
Exit Sub
End Sub

Public Function Greyscale(ByVal Colr As Long, PixelNum As Integer) As Integer
    Dim R As Long, G As Long, B As Long
    R = Colr Mod 256
    G = R Mod 256
    B = G Mod 256
    If R < 0 Then R = 0: If R > 255 Then R = 255
    If G < 0 Then G = 0: If G > 255 Then G = 255
    If B < 0 Then B = 0: If B > 255 Then B = 255
    Greyscale = PixelNum * (R / 255 + G / 255 + B / 255)
End Function

Private Sub Pic10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PixVal As Long, rR As Integer, rG As Integer, rB As Integer
PixVal = GetPixel(Pic10.hdc, X + I, Y + J) ' dapatkan pixel dari posisi x + i dan y + j
DecTORGB PixVal, rR, rG, rB
Text2.Text = rR & "-" & rG & "-" & rB
End Sub
