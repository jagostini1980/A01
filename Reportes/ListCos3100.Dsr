VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ListCos3100 
   Caption         =   "Costo por linea"
   ClientHeight    =   11115
   ClientLeft      =   -3945
   ClientTop       =   285
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "ListCos3100.dsx":0000
End
Attribute VB_Name = "ListCos3100"
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
    If TxtLinea.Text = "Total General ==>" Or TxtLinea.Text = "Total Linea ==>" Then
        Line2.Visible = True
        Line3.Visible = True
    Else
        Line2.Visible = False
        Line3.Visible = False
    End If

End Sub

Private Sub PageFooter_Format()
    pag = pag + 1
    TxtPagNro.Caption = "Página " & pag
    
End Sub

Private Sub PageHeader_Format()
    TxtFecha.Text = Date
End Sub

