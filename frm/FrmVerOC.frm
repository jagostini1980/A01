VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form FrmVerOC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Orden de Compra"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Orden de Compra"
      Enabled         =   0   'False
      Height          =   1725
      Left            =   180
      TabIndex        =   6
      Top             =   45
      Width           =   10095
      Begin VB.TextBox TxtNroOrden 
         Height          =   315
         Left            =   1350
         TabIndex        =   9
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox TxtResp 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   3480
      End
      Begin VB.TextBox TxtLugar 
         Height          =   315
         Left            =   6390
         MaxLength       =   50
         TabIndex        =   7
         Top             =   945
         Width           =   3570
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   8505
         TabIndex        =   10
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbFormaPago 
         Height          =   330
         Left            =   1530
         TabIndex        =   11
         Top             =   945
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1125
         TabIndex        =   12
         Top             =   1305
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   582
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin Controles.ComboEsp CmbEmp 
         Height          =   330
         Left            =   6390
         TabIndex        =   13
         Top             =   1305
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   7110
         TabIndex        =   14
         Top             =   585
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   582
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin VB.Label LBAnulada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anulada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   4545
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7830
         TabIndex        =   22
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Factura a Nombre de:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4455
         TabIndex        =   21
         Top             =   1350
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Responsable:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Órden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lugar de entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4770
         TabIndex        =   16
         Top             =   990
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo Emisor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4995
         TabIndex        =   15
         Top             =   630
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Re-Imprimir"
      Height          =   350
      Left            =   7695
      TabIndex        =   2
      Top             =   7920
      Width           =   1230
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9045
      TabIndex        =   5
      Text            =   "0"
      Top             =   4620
      Width           =   1275
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   2760
      Left            =   90
      TabIndex        =   1
      Top             =   5040
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   4868
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   9045
      TabIndex        =   3
      Top             =   7920
      Width           =   1230
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2715
      Left            =   135
      TabIndex        =   0
      Top             =   1845
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8415
      TabIndex        =   4
      Top             =   4680
      Width           =   510
   End
End
Attribute VB_Name = "FrmVerOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private VecOrdenDeCompra() As TipoOrdenDeCompra
Private VecCentroCta() As TipoCentroCta
Private Modificado As Boolean
Private A_Codigo As Integer
Private CantNoAsignada As Integer
Public NroOrden As Integer
Public CentroEmisor As String

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeOrden
    RepOrdenDeCompra.Show vbModal
End Sub

Private Sub Form_Load()
     Call CrearEncabezados
    
    'Call CargarComboCuentasContables(CmbCuentas)
    Call CargarComboProveedores(CmbProv)
    Call CargarComboEmpresas(CmbEmp)
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    'Call CargarCmbArtCompra(CmbArtCompra)
      
    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
    
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Selected = True
    
    
    Call CargarOrden(NroOrden)
End Sub

Private Sub ConfImpresionDeOrden()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "Articulo", adVarChar, 50
    RsListado.Fields.Append "Cantidad", adInteger
    RsListado.Fields.Append "Precio", adDouble
    RsListado.Fields.Append "Importe", adDouble
    RsListado.Open
    i = 1
    While i <= LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
        RsListado!Articulo = .Text
        RsListado!Cantidad = CInt(.SubItems(1))
        RsListado!Precio = Val(Replace(.SubItems(2), ",", "."))
        RsListado!Importe = Val(Replace(.SubItems(3), ",", "."))
      End With
        i = i + 1
    Wend
    RsListado.MoveFirst
    TxtNroOrden.Text = Format(NroOrden, "0000000000")
    
    RepOrdenDeCompra.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    RepOrdenDeCompra.TxtFormaDePago.Text = CmbFormaPago.Text
    RepOrdenDeCompra.TxtFecha = CStr(CalFecha.Value)
    RepOrdenDeCompra.TxtFactNombre.Text = CmbEmp.Text
    RepOrdenDeCompra.TxtNroOrden.Text = TxtNroOrden.Text
    RepOrdenDeCompra.TxtLugarDeEntrega = TxtLugar.Text
    RepOrdenDeCompra.TxtProv.Text = CmbProv.Text
    RepOrdenDeCompra.TxtResp.Text = TxtResp.Text
    'RepOrdenDeCompra.TxtAnulada.Visible = LBAnulada.Visible
    'RepOrdenDeCompra.TxtAnulada.Text = LBAnulada.Caption
    RepOrdenDeCompra.DataControl1.Recordset = RsListado

End Sub

Private Sub CargarOrden(NroOrden As Integer)
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Dim PeriodoCerrado As Boolean
    Dim RsValidarPeriodo As ADODB.Recordset
    Set RsValidarPeriodo = New ADODB.Recordset
    

  With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    Sql = "SpOCOrdenesDeCompraCabeceraTraerNro @NroOrden= " & NroOrden & _
                                            ", @Usuario='" & Usuario & _
                                           "', @O_CentroDeCostoEmisor=" & CentroEmisor
    .Open Sql, Conec
    
  If .EOF Then
      MsgBox "No existe una orden de compra con esa numeración", vbInformation
  Else
    '  Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(!O_Fecha, "MM/yyyy")) & "'"
    '  RsValidarPeriodo.Open Sql, Conec
    '  PeriodoCerrado = RsValidarPeriodo!Cerrado > 0

    If Not IsNull(!O_FechaAnulacion) Then
        If Not IsNull(!O_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!O_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
        
       ' LBPerCerrado.Visible = PeriodoCerrado
        'FraArt.Enabled = False
        'FrameAsig.Enabled = False
        'CmdCambiar.Enabled = False
        TxtLugar.Enabled = False
        TxtResp.Enabled = False
        CmbCentroDeCostoEmisor.Enabled = False
        CmbEmp.Enabled = False
        CmbFormaPago.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
    Else
        LBAnulada.Visible = False
        TxtLugar.Enabled = True
        TxtResp.Enabled = True
        CmbCentroDeCostoEmisor.Enabled = True
        CmbEmp.Enabled = True
        CmbFormaPago.Enabled = True
        CmbProv.Enabled = True
        CalFecha.Enabled = True
    End If
    
    TxtNroOrden.Text = Format(!O_NumeroOrdenDeCompra, "0000000000")
    Me.NroOrden = !O_NumeroOrdenDeCompra
    
    CmdImprimir.Enabled = True
        
    CalFecha.Value = !O_Fecha
    TxtResp = RsCargar!O_Responsable
    Call BuscarProveedor(!O_CodigoProveedor, CmbProv)
    TxtLugar.Text = !O_LugarDeEntrega
    CmbFormaPago.Text = !O_FormaDePagoPactada
    Call UbicarEmpresa(!O_EmpresaFacturaANombreDe, CmbEmp)
    Call BuscarCentroEmisor(!O_CentroDeCostoEmisor, CmbCentroDeCostoEmisor)
    
    .Close
    Sql = "SpOCOrdenesDeCompraRenglonesArticulosTraer @NroOrden=" & NroOrden & _
                                                   ", @O_CentroDeCostoEmisor=" & CentroEmisor
    .Open Sql, Conec
    ReDim VecOrdenDeCompra(.RecordCount)
    i = 1
    LvListado.ListItems.Clear
    
    While Not .EOF
        LvListado.ListItems.Add
        VecOrdenDeCompra(i).A_Codigo = !O_CodigoArticulo
        VecOrdenDeCompra(i).A_Descripcion = BuscarDescArt(!O_CodigoArticulo, BuscarTablaCentroEmisor(CentroEmisor))
        VecOrdenDeCompra(i).Cantidad = !O_CantidadPedida
        VecOrdenDeCompra(i).PrecioUnit = !O_PrecioPactado
        VecOrdenDeCompra(i).CantPendiente = !O_CantidadPendiente
        
        LvListado.ListItems(i).Text = VecOrdenDeCompra(i).A_Descripcion
        LvListado.ListItems(i).SubItems(1) = VecOrdenDeCompra(i).Cantidad
        LvListado.ListItems(i).SubItems(2) = Format(VecOrdenDeCompra(i).PrecioUnit, "0.00##")
        LvListado.ListItems(i).SubItems(3) = Format(VecOrdenDeCompra(i).Cantidad * VecOrdenDeCompra(i).PrecioUnit, "0.00##")
        .MoveNext
        i = i + 1
    Wend
    
    Call CalcularTotal
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
    .Close
    
    Sql = "SpOCOrdenesDeCompraRenglonesTraer @NroOrden=" & NroOrden & _
                                          ", @O_CentroDeCostoEmisor=" & CentroEmisor
    .Open Sql, Conec
        ReDim VecCentroCta(.RecordCount)
    i = 1
    While Not .EOF
        VecCentroCta(i).Centro_Descripcion = !CentroDeCosto
        VecCentroCta(i).Cta_Descripcion = Trim(!CtaContable)
        VecCentroCta(i).O_CantidadPedida = !O_CantidadPedida
        VecCentroCta(i).O_CodigoArticulo = !O_CodigoArticulo
        VecCentroCta(i).O_CentroDeCosto = !O_CentroDeCosto
        VecCentroCta(i).O_CuentaContable = !O_CuentaContable
        VecCentroCta(i).O_CantidadPendiente = !O_CantidadPendiente
        i = i + 1
        .MoveNext
    Wend
  End If
  End With
  LvListado.ListItems(1).Selected = True
  Call LvListado_ItemClick(LvListado.ListItems(1))
  
End Sub

Private Sub CrearEncabezados()
    LvListado.ColumnHeaders.Add , , "Descripción Artículo", LvListado.Width - 3850
    LvListado.ColumnHeaders.Add , , "Cantidad", 1000
    LvListado.ColumnHeaders.Add , , "Precio Unitario", 1500
    LvListado.ColumnHeaders.Add , , "Importe", 1000
    
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", (LvCenCostoCtas.Width / 2) - 660
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos", (LvCenCostoCtas.Width / 2) - 660
    LvCenCostoCtas.ColumnHeaders.Add , , "Cantidad", 1000
    LvCenCostoCtas.ColumnHeaders.Add , , "Index vector", 0

End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To LvListado.ListItems.Count
        Total = Total + Val(LvListado.ListItems(i).SubItems(3))
    Next
        txtTotal.Text = Format(Total, "0.00##")

End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
'NO SE TOCA
'   If Item.Index < LvListado.ListItems.Count Then
    A_Codigo = VecOrdenDeCompra(Item.Index).A_Codigo
       Call CargarLvCenCostoCtas(A_Codigo)
       
  '  Else
  '      LvCenCostoCtas.ListItems.Clear
  ' End If
'errores:
 '   ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub CargarLvCenCostoCtas(A_Codigo As Integer)
  Dim i As Integer
    LvCenCostoCtas.ListItems.Clear
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        If .O_CodigoArticulo = A_Codigo Then
            LvCenCostoCtas.ListItems.Add
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Text = .Cta_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(1) = .Centro_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(2) = .O_CantidadPedida
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(3) = i
        End If
      End With
    Next
 '   LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Selected = True
End Sub


Public Sub BuscarArt(Codigo As Integer, CmbArt As ComboEsp)
 Dim i As Integer
    i = 1
   While VecArtCompra(i).A_Codigo <> Codigo
    i = i + 1
   Wend
   CmbArt.ListIndex = i
End Sub


