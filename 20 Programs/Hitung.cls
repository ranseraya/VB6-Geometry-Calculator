VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function LPersegi(ByVal s As Double) As Double
    LPersegi = s * s
End Function

Public Function LPersegiPanjang(ByVal p As Double, ByVal l As Double) As Double
    LPersegiPanjang = p * l
End Function

Public Function LSegitiga(ByVal a As Double, ByVal t As Double) As Double
    LSegitiga = 0.5 * a * t
End Function

Public Function LLingkaran(ByVal r As Double) As Double
    LLingkaran = 3.14 * r ^ 2
End Function

Public Function LJajarGenjang(ByVal a As Double, ByVal t As Double) As Double
    LJajarGenjang = a * t
End Function

Public Function LTrapesium(ByVal sisiA As Double, ByVal sisiB As Double, ByVal t As Double) As Double
    LTrapesium = 0.5 * (sisiA + sisiB) * t
End Function

Public Function LLayangLayang(ByVal d1 As Double, ByVal d2 As Double) As Double
    LLayangLayang = 0.5 * d1 * d2
End Function

Public Function LBelahKetupat(ByVal d1 As Double, ByVal d2 As Double) As Double
    LBelahKetupat = 0.5 * d1 * d2
End Function

Public Function LElips(ByVal r1 As Double, ByVal r2 As Double) As Double
    LElips = 3.14 * r1 * r2
End Function

Public Function LJuring(ByVal r As Double, ByVal sudut As Double) As Double
    LJuring = (sudut / 360) * 3.14 * r ^ 2
End Function

Public Function VolKubus(ByVal s As Double) As Double
    VolKubus = s ^ 3
End Function

Public Function VolBalok(ByVal p As Double, ByVal l As Double, ByVal t As Double) As Double
    VolBalok = p * l * t
End Function

Public Function VolTabung(ByVal r As Double, ByVal t As Double) As Double
    VolTabung = (3.14 * r ^ 2) * t
End Function

Public Function VolKerucut(ByVal r As Double, ByVal t As Double) As Double
    VolKerucut = (1 / 3) * (3.14 * r ^ 2) * t
End Function

Public Function VolBola(ByVal r As Double) As Double
    VolBola = (4 / 3) * 3.14 * r ^ 3
End Function

Public Function VolPrismaSegitiga(ByVal a As Double, ByVal tS As Double, ByVal tP As Double) As Double
    VolPrismaSegitiga = (0.5 * a * tS) * tP
End Function

Public Function VolLimasSegiempat(ByVal p As Double, ByVal l As Double, ByVal t As Double) As Double
    VolLimasSegiempat = (1 / 3) * (p * l) * t
End Function

Public Function VolLimasSegitiga(ByVal a As Double, ByVal tS As Double, ByVal tL As Double) As Double
    VolLimasSegitiga = (1 / 3) * (0.5 * a * tS) * tL
End Function

Public Function VolPrismaSegiempat(ByVal p As Double, ByVal tA As Double, ByVal tP As Double) As Double
    VolPrismaSegiempat = p * tA * tP
End Function

Public Function VolPrismaTrapesium(ByVal sisiA As Double, ByVal sisiB As Double, ByVal tA As Double, ByVal tP As Double) As Double
    VolPrismaTrapesium = (0.5 * (sisiA + sisiB) * tA) * tP
End Function

