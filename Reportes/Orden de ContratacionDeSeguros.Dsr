VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RepOrdenDeContratacionDeSeguros 
   Caption         =   "Orden de Contratación de Seguros"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   300
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "Orden de ContratacionDeSeguros.dsx":0000
End
Attribute VB_Name = "RepOrdenDeContratacionDeSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    If TxtCuotas.BackStyle = ddBKNormal Then
        TxtCuotas.BackStyle = ddBKTransparent
        TxtFVencimiento.BackStyle = ddBKTransparent
        TxtImporte.BackStyle = ddBKTransparent
        
        TxtInterno.BackStyle = ddBKTransparent
        TxtDominio.BackStyle = ddBKTransparent
        TxtPrima.BackStyle = ddBKTransparent
        TxtOtrosCostos.BackStyle = ddBKTransparent
        txtTotal.BackStyle = ddBKTransparent
    Else
        TxtCuotas.BackStyle = ddBKNormal
        TxtFVencimiento.BackStyle = ddBKNormal
        TxtImporte.BackStyle = ddBKNormal
        
        TxtInterno.BackStyle = ddBKNormal
        TxtDominio.BackStyle = ddBKNormal
        TxtPrima.BackStyle = ddBKNormal
        TxtOtrosCostos.BackStyle = ddBKNormal
        txtTotal.BackStyle = ddBKNormal
    End If
    
End Sub

Private Sub PageHeader_Format()
    TxtFechaImp.Text = Now

End Sub
