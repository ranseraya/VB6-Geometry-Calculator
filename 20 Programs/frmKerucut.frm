VERSION 5.00
Begin VB.Form frmKerucut 
   Caption         =   "Kalkulator Volume Kerucut"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnKembali 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtTinggi 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung Volume"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtJariJari 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label tinggi 
      Caption         =   "Masukkan Tinggi"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label labelHasil 
      Caption         =   "Hasil: "
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Tag             =   "Hasil"
      Top             =   2640
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
      Caption         =   "Hitung Volume Kerucut"
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
      Left            =   45
      TabIndex        =   0
      Tag             =   "Judul"
      Top             =   360
      Width           =   6525
   End
End
Attribute VB_Name = "frmKerucut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnHitung_Click()
    Dim Kalkulator As New KalkulatorGeometri.Hitung
    
    If Not IsNumeric(txtJariJari.Text) Or Not IsNumeric(txtTinggi.Text) Then
        MsgBox "Input tidak valid! Pastikan semua kolom terisi dengan angka.", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    Dim jariJari As Double
    Dim tinggi As Double
    Dim HasilVolume As Double
    
    jariJari = Val(txtJariJari.Text)
    tinggi = Val(txtTinggi.Text)
    
    If tinggi < 0 Or jariJari < 0 Then
        MsgBox "Nilai tidak boleh negatif!", vbCritical, "Input Salah"
        Exit Sub
    End If
    
    HasilVolume = Kalkulator.VolKerucut(jariJari, tinggi)
    
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

