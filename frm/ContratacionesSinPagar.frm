VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ContratacionesSinPagar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Órdenes de Contración por Proveedor"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCrearRecepcion 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4012
      TabIndex        =   1
      Top             =   6795
      Width           =   1320
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5452
      TabIndex        =   2
      Top             =   6795
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6600
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   11642
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "ContratacionesSinPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemsChequeados As Integer
Public P_Codigo As Long
'Public Periodo As String

Private Sub CmdCrearRecepcion_Click()
  Dim i As Integer
    
    ReDim VecAutorizacionDePago(0)
    
    For i = 1 To LvListado.ListItems.Count
      If LvListado.ListItems(i).Checked Then
        ReDim Preserve VecAutorizacionDePago(UBound(VecAutorizacionDePago) + 1)
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_NumeroOrdenDeContratacion = Val(LvListado.ListItems(i).Text)
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_Fecha = LvListado.ListItems(i).SubItems(2)
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_PrecioPactado = ValN(LvListado.ListItems(i).SubItems(6))
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_CuentaContable = LvListado.ListItems(i).SubItems(7)
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_CentroDeCosto = LvListado.ListItems(i).SubItems(8)
        VecAutorizacionDePago(UBound(VecAutorizacionDePago)).O_CentroDeCostoEmisor = LvListado.ListItems(i).SubItems(9)
      End If
    Next
 'cierra el formulario actual luego de seleccionar los servicio
    Unload Me
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CreaEncabezado
    Call TraerContratacionesSinPagar(P_Codigo)
    
    ReDim VecAutorizacionDePago(0)

End Sub

Private Sub TraerContratacionesSinPagar(Proveedor As Long)
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    
    Sql = "SpOcOrdenesDeContratacionSinPagarTraer @P_Codigo=" & Proveedor '& _
                                             ", @Periodo =" & FechaSQL(Periodo, "SQL")
          
    LvListado.ListItems.Clear
    
    With RsCargar
        .Open Sql, Conec
        i = 1
       While Not .EOF
          LvListado.ListItems.Add
             
          LvListado.ListItems(i).Text = Format(!O_NumeroOrdenDeContratacion, "0000000000")
          LvListado.ListItems(i).SubItems(1) = BuscarDescCentroEmisor(!O_CentroDeCostoEmisor)
          LvListado.ListItems(i).SubItems(2) = VerificarNulo(!O_Fecha)
          LvListado.ListItems(i).SubItems(3) = VerificarNulo(!O_Observaciones)
          LvListado.ListItems(i).SubItems(4) = BuscarDescCta(!O_CuentaContable)
          LvListado.ListItems(i).SubItems(5) = BuscarDescCentro(!O_CentroDeCosto)
          LvListado.ListItems(i).SubItems(6) = Format(!O_PrecioPactado, "0.00##")
          LvListado.ListItems(i).SubItems(7) = !O_CuentaContable
          LvListado.ListItems(i).SubItems(8) = !O_CentroDeCosto
          LvListado.ListItems(i).SubItems(9) = !O_CentroDeCostoEmisor
          If VerificarNulo(!O_Autorizado, "B") = False Then
             LvListado.ListItems(i).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(4).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(5).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(6).ForeColor = vbRed
          End If
          i = i + 1
          .MoveNext
       Wend
    End With

End Sub

Private Sub CreaEncabezado()
    LvListado.ColumnHeaders.Add , , "Nº de Orden de Contratacion", 1300
    LvListado.ColumnHeaders.Add , , "Centro de Costo", 1500
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Observaciones", 2000
    LvListado.ColumnHeaders.Add , , "Cuenta Contable", 1800
    LvListado.ColumnHeaders.Add , , "Sub-Centro de Costo", 1800
    LvListado.ColumnHeaders.Add , , "Precio", 1000, 1
    LvListado.ColumnHeaders.Add , , "CodCta", 0
    LvListado.ColumnHeaders.Add , , "CodCentro", 0
    LvListado.ColumnHeaders.Add , , "CodCentroEmisor", 0
End Sub

Private Sub LvListado_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.ForeColor = vbRed Then
        MsgBox "La Orden de Contratación No está Autorizada", vbInformation
        Item.Checked = False
        Exit Sub
    End If
    
      If Item.Checked Then
         ItemsChequeados = ItemsChequeados + 1
      Else
         ItemsChequeados = ItemsChequeados - 1
      End If
      
      If ItemsChequeados > 0 Then
         CmdCrearRecepcion.Enabled = True
      Else
         CmdCrearRecepcion.Enabled = False
      End If
End Sub
