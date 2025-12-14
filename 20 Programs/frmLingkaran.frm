VERSION 5.00
Begin VB.Form frmLingkaran 
   Caption         =   "Kalkulator Luas Lingkaran"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnKembali 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtPanjangSisi 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label labelHasil 
      Caption         =   "Hasil: "
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Tag             =   "Hasil"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label jariJari 
      Caption         =   "Masukkan jari-jarii"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label labelJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hitung Luas Lingkaran"
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
      Left            =   15
      TabIndex        =   0
      Tag             =   "Judul"
      Top             =   360
      Width           =   6465
   End
End
Attribute VB_Name = "frmLingkaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnHitung_Click()
    Dim Kalkulator As New KalkulatorGeometri.Hitung
    
    If Not IsNumeric(txtPanjangSisi.Text) Then
        MsgBox "Input tidak valid! Pastikan semua kolom terisi dengan angka.", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    Dim Sisi As Double
    Dim HasilLuas As Double
    
    Sisi = Val(txtPanjangSisi.Text)
    
    If Sisi < 0 Then
        MsgBox "Nilai tidak boleh negatif!", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    HasilLuas = Kalkulator.LLingkaran(Sisi)
    
    labelHasil.Caption = "Hasil Luas: " & HasilLuas
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

