VERSION 5.00
Begin VB.Form frmPersegiPanjang 
   Caption         =   "Kalkulator Luas Persegi Panjang"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnKembali 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtLebar 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtPanjang 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lebar 
      Caption         =   "Masukkan Lebar"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label labelHasil 
      Caption         =   "Hasil:"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Tag             =   "Hasil"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label panjang 
      Caption         =   "Masukkan Panjang"
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
      Caption         =   "Hitung Luas Persegi Panjang"
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
      Left            =   105
      TabIndex        =   0
      Tag             =   "Judul"
      Top             =   360
      Width           =   6510
   End
End
Attribute VB_Name = "frmPersegiPanjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnHitung_Click()
    Dim Kalkulator As New KalkulatorGeometri.Hitung
    
    If Not IsNumeric(txtPanjang.Text) Or Not IsNumeric(txtLebar.Text) Then
        MsgBox "Input tidak valid! Pastikan semua kolom terisi dengan angka.", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    Dim panjang As Double
    Dim lebar As Double
    Dim HasilLuas As Double
    
    panjang = Val(txtPanjang.Text)
    lebar = Val(txtLebar.Text)
    
    If panjang < 0 Or lebar < 0 Then
        MsgBox "Nilai tidak boleh negatif!", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    HasilLuas = Kalkulator.LPersegiPanjang(panjang, lebar)
    
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

