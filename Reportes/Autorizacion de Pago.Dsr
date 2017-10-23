VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RepAutorizacionDePago 
   Caption         =   "Autorizacion de Carga Contable"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   300
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "Autorizacion de Pago.dsx":0000
End
Attribute VB_Name = "RepAutorizacionDePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Total As Double

Private Sub Detail_Format()
On Error Resume Next
    If Detail.BackStyle = ddBKNormal Then
        Detail.BackStyle = ddBKTransparent
    Else
        Detail.BackStyle = ddBKNormal
    End If
    
    Total = Total + import
End Sub

Private Sub PageHeader_Format()
    TxtFechaImp.Text = Now
End Sub

Private Sub ReportFooter_Format()
    tot.Text = Format(Total, "0.00##")
End Sub
