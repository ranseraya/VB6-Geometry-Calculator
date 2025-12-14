VERSION 5.00
Begin VB.Form frmJuringLingkaran 
   Caption         =   "Kalkulator Luas Juring Lingkaran"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnKembali 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtSudut 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtJariJari 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label sudut 
      Caption         =   "Masukkan sudut"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label labelHasil 
      Caption         =   "Hasil: "
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Tag             =   "Hasil"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label jariJari 
      Caption         =   "Masukkan jari-jari"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label labelJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hitung Luas Juring Lingkaran"
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
      Left            =   120
      TabIndex        =   0
      Tag             =   "Judul"
      Top             =   360
      Width           =   6555
   End
End
Attribute VB_Name = "frmJuringLingkaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnHitung_Click()
    Dim Kalkulator As New KalkulatorGeometri.Hitung
    
    If Not IsNumeric(txtJariJari.Text) Or Not IsNumeric(txtSudut.Text) Then
        MsgBox "Input tidak valid! Pastikan semua kolom terisi dengan angka.", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    Dim jariJari As Double
    Dim sudut As Double
    Dim HasilLuas As Double
    
    jariJari = Val(txtJariJari.Text)
    sudut = Val(txtSudut.Text)
    
    If jariJari < 0 Or sudut < 0 Or sudut > 360 Then
        MsgBox "Nilai tidak boleh negatif atau sudut melebihi 360!", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    HasilLuas = Kalkulator.LJuring(jariJari, sudut)
    
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

