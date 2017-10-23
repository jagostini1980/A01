VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_5900 
   Caption         =   "Reporte Presupuesto Financiero"
   ClientHeight    =   9660
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   11625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   20505
   _ExtentY        =   17039
   SectionData     =   "ListA01_5900.dsx":0000
End
Attribute VB_Name = "ListA01_5900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pag As Integer
Dim TotalGOF As Double

Private Sub DetailCuentas_Format()
    If TxtCentroEmisor.BackStyle = ddBKNormal Then
        TxtCentroEmisor.BackStyle = ddBKTransparent
        TxtImporte.BackStyle = ddBKTransparent
        TxtImporte2.BackStyle = ddBKTransparent
        TxtImporte3.BackStyle = ddBKTransparent
        TxtDesvio.BackStyle = ddBKTransparent
        TxtDesvio2.BackStyle = ddBKTransparent
    Else
        TxtCentroEmisor.BackStyle = ddBKNormal
        TxtImporte.BackStyle = ddBKNormal
        TxtImporte2.BackStyle = ddBKNormal
        TxtImporte3.BackStyle = ddBKNormal
        TxtDesvio.BackStyle = ddBKNormal
        TxtDesvio2.BackStyle = ddBKNormal
    End If
    If UCase(Mid(Trim(TxtCentroEmisor), 1, 5)) = "TOTAL" Then
        TxtCentroEmisor.Font.Bold = True
        TxtImporte.Font.Bold = True
        TxtImporte2.Font.Bold = True
        TxtImporte3.Font.Bold = True
        TxtDesvio.Font.Bold = True
        TxtDesvio2.Font.Bold = True
    Else
        TxtCentroEmisor.Font.Bold = False
        TxtImporte.Font.Bold = False
        TxtImporte2.Font.Bold = False
        TxtImporte3.Font.Bold = False
        TxtDesvio.Font.Bold = False
        TxtDesvio2.Font.Bold = False
    End If
End Sub

Private Sub GroupFooter1_BeforePrint()
    TotalGOF = TotalGOF - TxtTotImportes
    txtTotal.Visible = True
    TxtFondo.Visible = True
    txtTotal = TotalGOF
End Sub

Private Sub GroupHeader1_Format()
    If TxtTipo = "Ingresos" Then
        TxtNegRub.Text = "Unidad de Negocio"
        TxtCabImp.Text = ""
        TxtCabImp2.Text = "Imp. Proyección"
        TxtCabImp3.Text = "Imp. Real"
        TxtCabDesvio.Text = "Desvio Proyec./Real"
        TxtCabDesvio2.Text = ""
    Else
        TxtNegRub.Text = "Rubro"
        TxtCabImp.Text = "Pres. Financiero"
        TxtCabImp2.Text = "Pres. SGP"
        TxtCabImp3.Text = "Real Financiero"
        TxtCabDesvio.Text = "Pres. F/SGP"
        TxtCabDesvio2.Text = "Real Financiero/SGP"
    End If
End Sub

Private Sub PageFooter_Format()
    pag = pag + 1
    TxtPagNro.Caption = "Página " & pag
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

