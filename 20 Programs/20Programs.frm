VERSION 5.00
Begin VB.Form frmMenuUtama 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnLimasSegiempat 
      Caption         =   "Volume Limas Segiempat"
      Height          =   735
      Index           =   9
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnPrismaSegiempat 
      Caption         =   "Volume Prisma Segiempat"
      Height          =   735
      Index           =   8
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton btnPrismaTrapesium 
      Caption         =   "Volume Prisma Trapesium"
      Height          =   735
      Index           =   7
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton btnTrapesium 
      Caption         =   "Luas Trapesium"
      Height          =   735
      Index           =   6
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnElips 
      Caption         =   "Luas Elips"
      Height          =   735
      Index           =   5
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton btnJuring 
      Caption         =   "Luas Juring Lingkaran"
      Height          =   735
      Index           =   4
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton btnBola 
      Caption         =   "Volume Bola"
      Height          =   735
      Index           =   3
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnLimasSegitiga 
      Caption         =   "Volume Limas Segitiga"
      Height          =   735
      Index           =   2
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton btnLayang 
      Caption         =   "Luas Layang-Layang"
      Height          =   735
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton btnBelahKetupat 
      Caption         =   "Luas Belah Ketupat"
      Height          =   735
      Index           =   0
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnKerucut 
      Caption         =   "Volume Kerucut"
      Height          =   735
      Index           =   9
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton btnPrismaSegitiga 
      Caption         =   "Volume Prisma Segitiga"
      Height          =   735
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton btnTabung 
      Caption         =   "Volume Tabung"
      Height          =   735
      Index           =   7
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnBalok 
      Caption         =   "Volume Balok"
      Height          =   735
      Index           =   6
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton btnKubus 
      Caption         =   "Volume Kubus"
      Height          =   735
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnLingkaran 
      Caption         =   "Luas Lingkaran"
      Height          =   735
      Index           =   4
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton btnJajarGenjang 
      Caption         =   "Luas Jajar Genjang"
      Height          =   735
      Index           =   3
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton btnSegitiga 
      Caption         =   "Luas Segitiga"
      Height          =   735
      Index           =   2
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnPersegiPanjang 
      Caption         =   "Luas Persegi Panjang"
      Height          =   735
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton btnPersegi 
      Caption         =   "Luas Persegi"
      Height          =   735
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label judul 
      Caption         =   "Selamat Datang di Menu Utama"
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Tag             =   "Judul"
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBalok_Click(Index As Integer)
    frmBalok.Show
    Me.Hide
End Sub

Private Sub btnBelahKetupat_Click(Index As Integer)
    frmBelahKetupat.Show
    Me.Hide
End Sub

Private Sub btnBola_Click(Index As Integer)
    frmBola.Show
    Me.Hide
End Sub

Private Sub btnElips_Click(Index As Integer)
    frmElips.Show
    Me.Hide
End Sub

Private Sub btnJajarGenjang_Click(Index As Integer)
    frmJajarGenjang.Show
    Me.Hide
End Sub

Private Sub btnJuring_Click(Index As Integer)
    frmJuringLingkaran.Show
    Me.Hide
End Sub

Private Sub btnKerucut_Click(Index As Integer)
    frmKerucut.Show
    Me.Hide
End Sub

Private Sub btnKubus_Click(Index As Integer)
    frmKubus.Show
    Me.Hide
End Sub

Private Sub btnLayang_Click(Index As Integer)
    frmLayangLayang.Show
    Me.Hide
End Sub

Private Sub btnLimasSegiempat_Click(Index As Integer)
    frmLimasSegiempat.Show
    Me.Hide
End Sub

Private Sub btnLimasSegitiga_Click(Index As Integer)
    frmLimasSegitiga.Show
    Me.Hide
End Sub

Private Sub btnLingkaran_Click(Index As Integer)
    frmLingkaran.Show
    Me.Hide
End Sub

Private Sub btnPersegi_Click(Index As Integer)
    frmPersegi.Show
    Me.Hide
End Sub

Private Sub btnPersegiPanjang_Click(Index As Integer)
    frmPersegiPanjang.Show
    Me.Hide
End Sub

Private Sub btnPrismaSegiempat_Click(Index As Integer)
    frmPrismaSegiempat.Show
    Me.Hide
End Sub

Private Sub btnPrismaSegitiga_Click(Index As Integer)
    frmPrismaSegitiga.Show
    Me.Hide
End Sub

Private Sub btnPrismaTrapesium_Click(Index As Integer)
    frmPrismaTrapesium.Show
    Me.Hide
End Sub

Private Sub btnSegitiga_Click(Index As Integer)
    frmSegitiga.Show
    Me.Hide
End Sub

Private Sub btnTabung_Click(Index As Integer)
    frmTabung.Show
    Me.Hide
End Sub

Private Sub btnTrapesium_Click(Index As Integer)
    frmTrapesium.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    TerapkanTema Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
