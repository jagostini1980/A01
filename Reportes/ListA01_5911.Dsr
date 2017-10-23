VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_5911 
   Caption         =   "Detalle Financiero por cuenta Contable"
   ClientHeight    =   11115
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "ListA01_5911.dsx":0000
End
Attribute VB_Name = "ListA01_5911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pag As Integer

Private Sub DetailCuentas_Format()
    If DetailCuentas.BackStyle = ddBKNormal Then
        DetailCuentas.BackStyle = ddBKTransparent
    Else
        DetailCuentas.BackStyle = ddBKNormal
    End If
End Sub

Private Sub PageFooter_Format()
    pag = pag + 1
    TxtPagNro.Caption = "Página " & pag
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

