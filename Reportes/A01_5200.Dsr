VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_5200 
   Caption         =   "Consulta Por Cuenta Contable"
   ClientHeight    =   11115
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "A01_5200.dsx":0000
End
Attribute VB_Name = "ListA01_5200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pag As Integer
Private Total As Double
Private TotalPres As Double

Private Sub Detail_Format()
    If Detail.BackStyle = ddBKNormal Then
        Detail.BackStyle = ddBKTransparent
    Else
        Detail.BackStyle = ddBKNormal
    End If
    If TxtTotPresD <> "" Then
        Total = Total + TxtTotD
        TotalPres = TotalPres + TxtTotPresD
        TxtDesvD = Format(TxtTotD - TxtTotPresD, "##,##0.00")
        If TxtTotPresD = 0 Then
            TxtDesvPorcD = Format("1", "0.00 %")
        Else
            TxtDesvPorcD = Format(Round(TxtDesvD / TxtTotPresD, 4), "0.00 %")
        End If
    Else
        TxtDesvPorcD = ""
        TxtDesvD = ""
    End If
End Sub

Private Sub GroupFooter1_BeforePrint()
    txtTotGral = Format(Total, "##,##0.00")
    TxtTotPresGral = Format(TotalPres, "##,##0.00")
    TxtDesvGral = Format(txtTotGral - TxtTotPresGral, "##,##0.00")
    If TxtTotPresGral = 0 Then
       TxtDesvPorcGral = Format(100, "0.00") & " %"
    Else
       TxtDesvPorcGral = Format(Round(100 * TxtDesvGral / TxtTotPresGral, 2), "0.00") & " %"
    End If
End Sub

Private Sub PageFooter_Format()
    pag = pag + 1
    TxtPagNro.Caption = "Página " & pag
    
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

