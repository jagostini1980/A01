VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_3100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Órdenes por Proveedor"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCrearRecepcion 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1912
      TabIndex        =   2
      Top             =   6795
      Width           =   1320
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3352
      TabIndex        =   1
      Top             =   6795
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6510
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   11483
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
Attribute VB_Name = "A01_3100"
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
    
    ReDim Articulos(0)
    
    For i = 1 To LvListado.ListItems.Count
      If LvListado.ListItems(i).Checked Then
        ReDim Preserve Articulos(UBound(Articulos) + 1)
        Articulos(UBound(Articulos)) = Val(LvListado.ListItems(i).SubItems(3))
      End If
    Next
 'cierra el formulario actual antes de abrir el de creacion de recepciones
    Unload Me
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CreaEncabezado
    Call TraerArticulosPendientes(P_Codigo)
    
    ReDim Articulos(0)

End Sub

Private Sub TraerArticulosPendientes(Proveedor As Long)
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    
    Sql = "SpOcOrdenesDeCompraRecepcionTraer @P_Codigo=" & Proveedor
          
    LvListado.ListItems.Clear
    
    With RsCargar
        .Open Sql, Conec
        i = 1
       While Not .EOF
          LvListado.ListItems.Add
          
          LvListado.ListItems(i).Text = BuscarDescArt(!O_CodigoArticulo, Trim(!C_TablaArticulos))
          LvListado.ListItems(i).SubItems(1) = !O_CantidadPendiente
          LvListado.ListItems(i).SubItems(2) = !O_CantidadPedida - !O_CantidadPendiente
          LvListado.ListItems(i).SubItems(3) = !O_CodigoArticulo
          If VerificarNulo(!O_Autorizado, "B") = False Then
             LvListado.ListItems(i).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
             LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
          End If
          
          i = i + 1
          .MoveNext
       Wend
    End With

End Sub

Private Sub CreaEncabezado()
    LvListado.ColumnHeaders.Add , , "Artículo", LvListado.Width - 3000
    LvListado.ColumnHeaders.Add , , "Cant. Pendiente", 1500, 1
    LvListado.ColumnHeaders.Add , , "Cant.Recibida", 1300, 1
    LvListado.ColumnHeaders.Add , , "A_Codigo", 0
End Sub

Private Sub LvListado_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.ForeColor = vbRed Then
        MsgBox "La Orden de Compra No está Autorizada", vbInformation
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
