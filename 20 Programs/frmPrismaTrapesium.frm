VERSION 5.00
Begin VB.Form frmPrismaTrapesium 
   Caption         =   "Kalkulator Volume Prisma Trapesium"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnKembali 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtTinggiP 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtTinggiT 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtAlas2 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung Volume"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtAlas1 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label tinggiP 
      Alignment       =   1  'Right Justify
      Caption         =   "Masukkan Tinggi Prisma"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label tinggiT 
      Caption         =   "Masukkan Tinggi Trapesium"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label alas2 
      Caption         =   "Masukkan Panjang Alas Bawah"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label labelHasil 
      Caption         =   "Hasil: "
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Tag             =   "Hasil"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label alas1 
      Caption         =   "Masukkan Panjang Alas Atas"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label labelJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hitung Volume  Prisma Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Tag             =   "Judul"
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmPrismaTrapesium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnHitung_Click()
    Dim Kalkulator As New KalkulatorGeometri.Hitung
    
    If Not IsNumeric(txtAlas1.Text) Or Not IsNumeric(txtAlas2.Text) Or Not IsNumeric(txtTinggiP.Text) Or Not IsNumeric(txtTinggiT.Text) Then
        MsgBox "Input tidak valid! Pastikan semua kolom terisi dengan angka.", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    Dim alas1 As Double
    Dim alas2 As Double
    Dim tinggiT As Double
    Dim tinggiP As Double
    Dim HasilVolume As Double
    
    alas1 = Val(txtAlas1.Text)
    alas2 = Val(txtAlas2.Text)
    tinggiT = Val(txtTinggiT.Text)
    tinggiP = Val(txtTinggiP.Text)
    
    If alas1 < 0 Or alas2 < 0 Or tinggiT < 0 Or tinggiP < 0 Then
        MsgBox "Nilai tidak boleh negatif!", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    HasilVolume = Kalkulator.VolPrismaTrapesium(alas1, alas2, tinggiT, tinggiP)
    
    labelHasil.Caption = "Hasil Volume: " & HasilVolume
End Sub

Private Sub btnKembali_Click()
    frmMenuUtama.Show
    Unload Me
End Sub

Private Sub Form_Load()
    TerapkanTema Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMenuUtama.Show
End Sub

