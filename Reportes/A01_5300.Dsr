VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_5300 
   Caption         =   "Consulta Por Centro de Costo"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   14817
   SectionData     =   "A01_5300.dsx":0000
End
Attribute VB_Name = "ListA01_5300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pag As Integer

Private Sub Detail_Format()
    If Detail.BackStyle = ddBKNormal Then
        Detail.BackStyle = ddBKTransparent
    Else
        Detail.BackStyle = ddBKNormal
    End If
    TxtDesvArt = Format(TxtRTotal - TxtPTotal, "##,##0.00")
    If TxtPTotal = 0 Then
        TxtDesvPorcArt = Format("1", "0.00 %")
    Else
        TxtDesvPorcArt = Format(Round(TxtDesvArt / TxtPTotal, 4), "0.00 %")
    End If
End Sub

Private Sub GroupFooter2_BeforePrint()
    TxtDesvGral = Format(TxtTotGral - TxtTotPresGral, "##,##0.00")
    TxtDesvPorcGral = Format(Round(TxtDesvGral / TxtTotPresGral, 4), "0.00 %")
End Sub

Private Sub GroupCentro_Format()
    TxtDesvio = Format(TxtTot - TxtTotPres, "##,##0.00")
    If TxtTotPres = 0 Then
        TxtDesvPorc = Format("1", "0.00 %")
    Else
        TxtDesvPorc = Format(Round(TxtDesvio / TxtTotPres, 4), "0.00 %")
    End If
End Sub

Private Sub GroupSubCentro_Format()
    TxtDesvioD = Format(TxtTotD - TxtTotPresD, "##,##0.00")
    If TxtTotPresD = 0 Then
        TxtDesvPorcD = Format("1", "0.00 %")
    Else
        TxtDesvPorcD = Format(Round(TxtDesvioD / TxtTotPresD, 4), "0.00 %")
    End If
End Sub

Private Sub PageFooter_BeforePrint()
    pag = pag + 1
    TxtPagNro.Caption = "Página " & pag
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

