VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListA01_1400 
   Caption         =   "Centros de Costos"
   ClientHeight    =   11115
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "A01_1400.dsx":0000
End
Attribute VB_Name = "ListA01_1400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pag As Integer

Private Sub DetailSubCentros_Format()
    If DetailSubCentros.BackStyle = ddBKNormal Then
        DetailSubCentros.BackStyle = ddBKTransparent
    Else
        DetailSubCentros.BackStyle = ddBKNormal
    End If

End Sub

Private Sub PageFooter_Format()
    pag = pag + 1
    TxtPagNro.Caption = "P�gina " & pag
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

