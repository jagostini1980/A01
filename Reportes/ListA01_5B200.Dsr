VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_5B200 
   Caption         =   "Consulta Rubros Contables"
   ClientHeight    =   10155
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   11295
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   19923
   _ExtentY        =   17912
   SectionData     =   "ListA01_5B200.dsx":0000
End
Attribute VB_Name = "ListA01_5B200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalGOF As Double

Private Sub DetailCuentas_Format()
    If DetailCuentas.BackStyle = ddBKNormal Then
        DetailCuentas.BackStyle = ddBKTransparent
    Else
        DetailCuentas.BackStyle = ddBKNormal
    End If
    
End Sub

Private Sub PageFooter_Format()
    TxtPagNro.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

