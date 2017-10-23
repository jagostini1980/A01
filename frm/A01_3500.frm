VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_3500 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobación de Órdenes"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LvRenglones 
      Height          =   2400
      Left            =   60
      TabIndex        =   2
      Top             =   5445
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
      Left            =   5234
      TabIndex        =   4
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   10230
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   330
         Left            =   9360
         TabIndex        =   0
         Top             =   180
         Width           =   765
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49676289
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49676289
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   6420
         TabIndex        =   12
         Top             =   195
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
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
      Begin VB.Label LbCentroEmisor 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro Emisor:"
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
         Left            =   5130
         TabIndex        =   13
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Desde:"
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
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label LBFechaHasta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   255
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2400
      Left            =   45
      TabIndex        =   1
      Top             =   615
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   4233
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confir&mar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      Top             =   7935
      Width           =   1300
   End
   Begin MSComctlLib.ListView LvOrdenesDeContratacion 
      Height          =   2400
      Left            =   45
      TabIndex        =   11
      Top             =   3030
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   4233
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label LbTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Aprobado: $"
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
      Left            =   8685
      TabIndex        =   6
      Top             =   8010
      Width           =   1545
   End
End
Attribute VB_Name = "A01_3500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoPresupuesto
    Numero As Integer
    CodigoArticulo As Integer
    CuentaContable As String
    CentroDeCosto As Integer
    'P_Cantidad As Long
    'P_ImporteUnitario As Double
    Importe As Double
    FechaAprobacion As String
    'P_CentroDeCostoEmisor As String
    'P_Observaciones As String
    'P_ObservacionesPresupuesto As String
    Estado As String
End Type

Private VecCabecera() As TipoOrdenDeCompra

Private VecRenglones() As TipoPresupuesto
Private Modificado As Boolean
Private LvIndex As Integer

Private Sub CalFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call CmdTraer_Click
    End If
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
Dim Rta As Integer

    Rta = MsgBox("¿Desea confirmar la Aprobación del Presupuesto?", vbYesNo)
    If Rta = vbYes Then
        Call AprobarPresupuesto
    End If
End Sub

Private Sub AprobarPresupuesto()
  Dim Sql As String
  Dim i As Integer
On Error GoTo ErrorInsert
  
    Conec.BeginTrans
    For i = 1 To LvListado.ListItems.Count
        With LvListado.ListItems(i)
             If .Checked Then
                  Sql = "SpOCOrdenesDeCompraCabeceraAutorizar @O_NumeroOrdenDeCompra = " & .SubItems(1) & _
                                                           ", @O_CentroDeCostoEmisor = '" & .SubItems(4) & _
                                                          "', @O_UsuarioAutorizo ='" & Usuario & "'"
            
                  Conec.Execute Sql
             End If
        End With
    Next
    
    For i = 1 To LvOrdenesDeContratacion.ListItems.Count
        With LvOrdenesDeContratacion.ListItems(i)
             If .Checked Then
                  Sql = "SpOcOrdenesDeContratacionCabeceraAutorizar @O_NumeroOrdenDeContratacion = " & .SubItems(1) & _
                                                                 ", @O_CentroDeCostoEmisor = '" & .SubItems(4) & _
                                                                "', @O_UsuarioAutorizo ='" & Usuario & "'"
            
                  Conec.Execute Sql
             End If
        End With
    Next
    
    Modificado = False
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       MsgBox "La Aprobación  se realizó correctamente"
    Else
       Conec.RollbackTrans
       Call TratarError(Err.Number, Err.Description)
    End If
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV
End Sub

Private Sub CargarLV()
    Dim Sql As String
    Dim i As Integer
    LvIndex = 0
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Sql = "SpOCOrdenesDeCompraCabeceraAutorizarTraer @FechaDesde =" & FechaSQL(CalFechaDesde, "SQL") & _
                                                  ", @FechaHasta=" & FechaSQL(CalFechaHasta, "SQL") & _
                                                  ", @CentroDeCostoEmisor='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
    With RsCargar
        .Open Sql, Conec
        i = 1
        LvListado.Sorted = False
        LvListado.ListItems.Clear
        While Not .EOF
            LvListado.ListItems.Add
            LvListado.ListItems(i).SubItems(1) = Format(!O_NumeroOrdenDeCompra, "00000000")
            LvListado.ListItems(i).SubItems(2) = !O_Fecha
            LvListado.ListItems(i).SubItems(3) = BuscarDescCentroEmisor(!O_CentroDeCostoEmisor)
            LvListado.ListItems(i).SubItems(4) = !O_CentroDeCostoEmisor
            LvListado.ListItems(i).SubItems(5) = !O_Responsable
            LvListado.ListItems(i).SubItems(6) = BuscarDescProv(!O_CodigoProveedor)
            'LvListado.ListItems(i).Checked = !RengDesaprobados = 0
            i = i + 1
            .MoveNext
        Wend
        .Close
         Sql = "SpOcOrdenesDeContratacionCabeceraAutorizarTraer @FechaDesde =" & FechaSQL(CalFechaDesde, "SQL") & _
                                                             ", @FechaHasta=" & FechaSQL(CalFechaHasta, "SQL") & _
                                                             ", @CentroDeCostoEmisor='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
        .Open Sql, Conec
        i = 1
        LvOrdenesDeContratacion.Sorted = False
        LvOrdenesDeContratacion.ListItems.Clear
        While Not .EOF
            LvOrdenesDeContratacion.ListItems.Add
            LvOrdenesDeContratacion.ListItems(i).SubItems(1) = Format(!O_NumeroOrdenDeContratacion, "00000000")
            LvOrdenesDeContratacion.ListItems(i).SubItems(2) = !O_Fecha
            LvOrdenesDeContratacion.ListItems(i).SubItems(3) = BuscarDescCentroEmisor(!O_CentroDeCostoEmisor)
            LvOrdenesDeContratacion.ListItems(i).SubItems(4) = !O_CentroDeCostoEmisor
            LvOrdenesDeContratacion.ListItems(i).SubItems(5) = !O_Responsable
            LvOrdenesDeContratacion.ListItems(i).SubItems(6) = BuscarDescProv(!O_CodigoProveedor)
            'LvListado.ListItems(i).Checked = !RengDesaprobados = 0
            i = i + 1
            .MoveNext
        Wend
    End With
    LvListado.Sorted = True
    
    LvRenglones.ListItems.Clear
    If LvListado.ListItems.Count > 0 Then
        CmdConfirmar.Enabled = True
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    Else
        CmdConfirmar.Enabled = LvListado.ListItems.Count + LvOrdenesDeContratacion.ListItems.Count
    End If
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Modificado = False
    
    CalFechaDesde.Value = Date - 30
    CalFechaHasta.Value = Date
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.ListIndex = 0
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A013500") = 2

    LvIndex = 0
    ReDim VecCabecera(0)
    ReDim VecRenglones(0)
    ReDim VecDistribucionPresupuesto(0)
End Sub

Private Sub CrearEncabezados()
    
    LvListado.ColumnHeaders.Add , , "Aprobado", 900
    LvListado.ColumnHeaders.Add , , "O. Compra Nº", 1300, 1
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Centro de Costo Emisor", (LvListado.Width - 3550) / 3
    LvListado.ColumnHeaders.Add , , "CentroEmisor", 0
    LvListado.ColumnHeaders.Add , , "Responsable", (LvListado.Width - 3750) / 3
    LvListado.ColumnHeaders.Add , , "Proveedor", (LvListado.Width - 3600) / 3
    
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "Aprobado", 900
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "O. Contratación Nº", 1300, 1
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "Fecha", 1100
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "Centro de Costo Emisor", (LvOrdenesDeContratacion.Width - 3550) / 3
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "CentroEmisor", 0
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "Responsable", (LvOrdenesDeContratacion.Width - 3750) / 3
    LvOrdenesDeContratacion.ColumnHeaders.Add , , "Proveedor", (LvOrdenesDeContratacion.Width - 3600) / 3
    
    LvRenglones.ColumnHeaders.Add , , "Artículo", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Cuenta Contable", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Sub-Centro de Costo", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Importe Total", 1200, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Rta As Integer
    If Modificado Then
       Rta = MsgBox("Ha efectuedo cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            Call AprobarPresupuesto
         End If
       End If
    End If
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
    Call CargarLvRenglones(Val(Item.SubItems(1)))
    LvIndex = Item.Index
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarLvRenglones(Nro As Integer)
  Dim i As Integer
  Dim Sql As String
  Dim TablaArticulos As String
  Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LvRenglones.ListItems.Clear
    Sql = "SpOCOrdenesDeCompraRenglonesTraer @NroOrden = " & Nro & _
                                          ", @O_CentroDeCostoEmisor = '" & LvListado.SelectedItem.SubItems(4) & "'"
    
      With RsCargar
        .Open Sql, Conec
        i = 1
        'ReDim VecRenglones(.RecordCount)
        TablaArticulos = VecCentroDeCostoEmisor(BuscarIndexCentroEmisor(LvListado.SelectedItem.SubItems(4))).C_TablaArticulos
  
        While Not .EOF
            LvRenglones.ListItems.Add
            LvRenglones.ListItems(i).Text = BuscarDescArt(!O_CodigoArticulo, TablaArticulos)
            LvRenglones.ListItems(i).SubItems(1) = BuscarDescCta(!O_CuentaContable)
            LvRenglones.ListItems(i).SubItems(2) = BuscarDescCentro(!O_CentroDeCosto)
            LvRenglones.ListItems(i).SubItems(3) = Format(!O_CantidadPedida * !O_PrecioPactado, "#.00")
            i = i + 1
            .MoveNext
        Wend
     End With
     
     Call CalcularTotal
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CalcularTotal()
   Dim Total As Double
   Dim i As Integer
   
    For i = 1 To LvRenglones.ListItems.Count
       If LvRenglones.ListItems(i).Checked Then
          Total = Total + Val(Replace(LvRenglones.ListItems(i).SubItems(3), ",", "."))
       End If
    Next
    
    LbTotal.Caption = "Total Aprobado: $" & Format(Total, "0.00")
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LvOrdenesDeContratacion_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
    Call CargarLvRenglonesContratacion(Val(Item.SubItems(1)))
    LvIndex = Item.Index
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub CargarLvRenglonesContratacion(NroOrden As Integer)
  Dim i As Integer
  Dim Sql As String
  Dim TablaArticulos As String
  Dim RsCargar As New ADODB.Recordset

'On Error GoTo Error
    LvRenglones.ListItems.Clear
    
    Sql = "SpOCOrdenesDeContratacionRenglonesTraer @NroOrden= " & NroOrden & _
                                                ", @O_CentroDeCostoEmisor = '" & LvOrdenesDeContratacion.SelectedItem.SubItems(4) & "'"

      With RsCargar
        .Open Sql, Conec
        i = 1
        'ReDim VecRenglones(.RecordCount)
        While Not .EOF
            LvRenglones.ListItems.Add
            LvRenglones.ListItems(i).Text = "Servicio"
            LvRenglones.ListItems(i).SubItems(1) = BuscarDescCta(!O_CuentaContable)
            LvRenglones.ListItems(i).SubItems(2) = BuscarDescCentro(!O_CentroDeCosto)
            LvRenglones.ListItems(i).SubItems(3) = Format(!O_PrecioPactado, "#.00")
            i = i + 1
            .MoveNext
        Wend
     End With
     
     Call CalcularTotal
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub


