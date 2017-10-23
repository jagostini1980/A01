VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_3200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Mercaderías"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCantTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cantidad Total Recibida"
      Height          =   1050
      Left            =   9675
      TabIndex        =   34
      Top             =   4500
      Width           =   2310
      Begin VB.CommandButton CmdAsignarCant 
         Caption         =   "Recibir Cant."
         Enabled         =   0   'False
         Height          =   350
         Left            =   540
         TabIndex        =   36
         Top             =   585
         Width           =   1300
      End
      Begin VB.TextBox TxtCantTotal 
         Height          =   315
         Left            =   1395
         TabIndex        =   35
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cant. Recibida:"
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
         Left            =   45
         TabIndex        =   37
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4680
      TabIndex        =   33
      Top             =   7375
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Autorización"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7395
      TabIndex        =   16
      Top             =   7365
      Width           =   1815
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   3375
      TabIndex        =   14
      Top             =   7365
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   6030
      TabIndex        =   15
      Top             =   7365
      Width           =   1230
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   9315
      TabIndex        =   17
      Top             =   7380
      Width           =   1320
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10755
      TabIndex        =   19
      Top             =   7365
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   135
      Top             =   7335
   End
   Begin VB.Frame FrameCant 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cant. por Centro y Cta."
      Height          =   1005
      Left            =   9675
      TabIndex        =   25
      Top             =   5760
      Width           =   2310
      Begin VB.TextBox TxtCantRecibida 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   495
         Width           =   915
      End
      Begin VB.CommandButton CmdModifCant 
         Caption         =   "Recibir Cant."
         Enabled         =   0   'False
         Height          =   350
         Left            =   1080
         TabIndex        =   13
         Top             =   540
         Width           =   1125
      End
      Begin VB.Label LbCant 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cant. por Centro / Cta.:"
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
         TabIndex        =   26
         Top             =   270
         Width           =   2025
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   10620
      TabIndex        =   24
      Text            =   "0"
      Top             =   6960
      Width           =   1365
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   2850
      Left            =   45
      TabIndex        =   11
      Top             =   4410
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   5027
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Recepción"
      Height          =   1770
      Left            =   45
      TabIndex        =   20
      Top             =   0
      Width           =   11985
      Begin VB.TextBox TxtFactura 
         Height          =   315
         Left            =   9180
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1035
         Width           =   2685
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1485
         TabIndex        =   9
         Top             =   1395
         Width           =   10410
      End
      Begin VB.TextBox TxtResp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2745
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1035
         Width           =   5595
      End
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         TabIndex        =   5
         Top             =   660
         Width           =   1050
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   3060
         TabIndex        =   1
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4185
         TabIndex        =   2
         Top             =   225
         Width           =   1000
      End
      Begin VB.TextBox TxtLugar 
         Height          =   315
         Left            =   7290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   660
         Width           =   4605
      End
      Begin VB.TextBox TxtNroRecepcion 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   10485
         TabIndex        =   3
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   103219201
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   315
         Left            =   1125
         TabIndex        =   4
         Top             =   660
         Width           =   2940
         _ExtentX        =   5186
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Factura:"
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
         TabIndex        =   38
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Observaciones:"
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
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Responsable de la Recepcion:"
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
         TabIndex        =   31
         Top             =   1095
         Width           =   2625
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
         Left            =   6075
         TabIndex        =   30
         Top             =   225
         Visible         =   0   'False
         Width           =   2400
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
         TabIndex        =   29
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lugar de Recepción:"
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
         Left            =   5445
         TabIndex        =   27
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Recepción:"
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
         TabIndex        =   22
         Top             =   270
         Width           =   1530
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
         Left            =   9810
         TabIndex        =   21
         Top             =   225
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2490
      Left            =   45
      TabIndex        =   10
      Top             =   1845
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
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
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   9315
      TabIndex        =   18
      Top             =   7365
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label LbDetalle 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   28
      Top             =   3960
      Width           =   7890
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
      Left            =   10080
      TabIndex        =   23
      Top             =   7020
      Width           =   510
   End
End
Attribute VB_Name = "A01_3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VecRecepcion() As TipoRecepcion
Public Proveedor As Integer
'este vector no se aparea con el Lv
Private VecCentroCta() As TipoCentroCtaRecepcion
Private Modificado As Boolean
Private A_Codigo As Long
Public TablaArticulos As String
Public NroRecepcion As Long
Private FechaMin As Date
Private UsuarioResponsable As String
'Private RsBloquearOc As ADODB.Recordset

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalFecha_Validate(Cancel As Boolean)
    If CalFecha.Value < FechaMin Then
        MsgBox "La fecha debe ser posterior al " & FechaMin, vbInformation, "Fecha inválida"
        CalFecha.Value = FechaMin
    End If
End Sub

Private Sub CmbProv_Click()
    CmdTraer.Enabled = CmbProv.ListIndex <> 0
End Sub

Private Sub CmdAnular_Click()
 Dim Sql As String
 Dim Rta As Integer
 On Error GoTo Error
    Rta = MsgBox("¿Está seguro de que desea Anular la Recepción?", vbYesNo)
    If Rta = vbYes Then
        Sql = "SpOCRecepcionOrdenDeCompraCabeceraAnular @R_NumeroRecepcion =" + CStr(NroRecepcion)
        Conec.Execute Sql
        MsgBox "La recepción se Anuló correctamente", vbInformation
    End If
Error:

  If Err.Number <> 0 Then
     MsgBox "Error", vbCritical
  Else
     Rta = MsgBox("¿Desea realizar otra acción?", vbYesNo)
     If Rta = vbYes Then
        Call LimpiarRecepcion
     Else
        Unload Me
     End If
  End If
End Sub

Private Sub CmdAsignarCant_Click()
'On Error GoTo Errores
  Dim i As Integer
  Dim j As Integer
  Dim PosVec As Integer
  Dim cant As Double
  Dim CantAsig As Double
  
  'i = LvListado.SelectedItem.Index
  i = Val(LvListado.SelectedItem.ListSubItems(6))
'agrega al vector
   
   If VecRecepcion(i).O_CantidadPendiente >= Val(TxtCantTotal.Text) Then
      cant = Val(TxtCantTotal.Text)
      Modificado = True
      j = 1
      While j <= LvCenCostoCtas.ListItems.Count
         PosVec = Val(LvCenCostoCtas.ListItems(j).SubItems(5))
         If cant > 0 Then
            CantAsig = IIf(cant >= VecCentroCta(PosVec).O_CantidadPendiente, VecCentroCta(PosVec).O_CantidadPendiente, cant)
            cant = cant - VecCentroCta(PosVec).O_CantidadPendiente
         Else
            CantAsig = 0
         End If
         
          VecCentroCta(PosVec).CantRecibida = CantAsig
          'lo pone en el LV
          LvCenCostoCtas.ListItems(j).SubItems(4) = VecCentroCta(PosVec).CantRecibida
          j = j + 1
      Wend
        'If LvListado.SelectedItem.Index < LvListado.ListItems.Count Then
        '    LvListado.ListItems(LvListado.SelectedItem.Index + 1).Selected = True
        '    Call LvListado_ItemClick(LvListado.SelectedItem)
        'End If

    Else
        MsgBox "La cantidad Recibida no puede ser mayor a la pendiente"
    End If
    Call CalcularTotal

End Sub

Private Sub CMDBuscar_Click()
   ' NroOrden = 0
    Unload BuscarRecepcion
   BuscarRecepcion.Show vbModal

    Timer1.Enabled = True
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar la Recepción?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarRecepcion
    End If
End Sub

Private Sub CmdCargar_Click()
    Call CargarRecepcion(Val(TxtNroRecepcion))
    Modificado = False
End Sub

Private Sub CargarRecepcion(NroRecepcion As Long)
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim PeriodoCerrado As Boolean
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    
    LBAnulada.Visible = False
    
    j = 1
 With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    ReDim VecRecepcion(0)
    
    LvListado.ListItems.Clear
    LvCenCostoCtas.ListItems.Clear
    
    Sql = "SpOCRecepcionOrdenDeCompraCabeceraTraerNro " + _
          " @NroRecepcion = " + CStr(NroRecepcion) + _
          ", @Usuario = '" + Usuario + "'"
    
    .Open Sql, Conec
     j = 1
    If .EOF Then
        MsgBox "No existe una recepcion con ese número"
        Exit Sub
    End If
    'trae el usuario responsable de la recepción
    UsuarioResponsable = !U_Usuario
    LvListado.Sorted = False
    
    If Not IsNull(!R_FechaAnulacion) Then
        If Not IsNull(!R_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!R_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
    
        FrameCant.Enabled = False
        CmdCambiar.Enabled = False
        TxtLugar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
        TxtFactura.Enabled = False
    Else
        LBAnulada.Visible = False
        FrameCant.Enabled = True
        CmdCambiar.Enabled = True
        TxtLugar.Enabled = True
        CmbProv.Enabled = True
        CalFecha.Enabled = True
        CmdAnular.Visible = True
        TxtObs.Enabled = True
        TxtFactura.Enabled = True
    End If
  
     CalFecha.Value = !R_Fecha
     CalFecha.Enabled = False
     Proveedor = !R_CodigoProveedor
     TxtLugar.Text = !R_LugarDeRecepcion
     'CmbProv.Text = BuscarDescProv(!R_CodigoProveedor)
     Call UbicarProveedor(!R_CodigoProveedor, CmbProv)
     CmbProv.Enabled = False
     CmdTraer.Enabled = False
     TxtNroRecepcion.Text = Format(NroRecepcion, "0000000000")
     TxtResp.Text = VerificarNulo(!R_ResponsableRecepcion)
     TxtObs.Text = VerificarNulo(!R_Observaciones)
     TxtFactura.Text = VerificarNulo(!R_Factura)
     Me.NroRecepcion = NroRecepcion
     .Close
     Sql = "SpOCRecepcionOrdenesDeCompraRenglonesArticulosTraer " + _
           " @NroRecepcion = " + CStr(NroRecepcion)
     .Open Sql, Conec
     ReDim Preserve VecRecepcion(.RecordCount)
    While Not .EOF
        VecRecepcion(j).A_Codigo = !R_CodigoArticulo
        VecRecepcion(j).A_Descripcion = BuscarDescArt(!R_CodigoArticulo, BuscarTablaCentroEmisor(!O_CentroDeCostoEmisor))
        VecRecepcion(j).O_CantidadPedida = !O_CantidadPedida
        VecRecepcion(j).O_CantidadPendiente = !O_CantidadPendiente
        VecRecepcion(j).O_NumeroOrdenDeCompra = !O_NumeroOrdenDeCompra
        VecRecepcion(j).O_PrecioPactado = VerificarNulo(!R_PrecioOrdenDeCompra, "N")
        VecRecepcion(j).R_Precio = VerificarNulo(!R_PrecioRecepcionado, "N")
        VecRecepcion(j).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        j = j + 1
        .MoveNext
    Wend

    .Close
   For i = 1 To UBound(VecRecepcion)
     LvListado.ListItems.Add
     LvListado.ListItems(i).Text = Format(VecRecepcion(i).O_NumeroOrdenDeCompra, "000000000")
     LvListado.ListItems(i).SubItems(1) = BuscarDescCentroEmisor(VecRecepcion(i).O_CentroDeCostoEmisor)
     LvListado.ListItems(i).SubItems(2) = VecRecepcion(i).A_Descripcion
     LvListado.ListItems(i).SubItems(3) = VecRecepcion(i).O_CantidadPedida
     LvListado.ListItems(i).SubItems(4) = VecRecepcion(i).O_CantidadPendiente
     LvListado.ListItems(i).SubItems(5) = Format(VecRecepcion(i).R_Precio, "0.00##")
     LvListado.ListItems(i).SubItems(6) = i
   Next
     Sql = "SpOCRecepcionOrdenesDeCompraRenglonesTraer @NroRecepcion=" & NroRecepcion
     .Open Sql, Conec
     ReDim VecCentroCta(.RecordCount)
     i = 1
     
    While Not .EOF
        VecCentroCta(i).NroOrden = !O_NumeroOrdenDeCompra
        VecCentroCta(i).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        VecCentroCta(i).O_FormaDePagoPactada = !O_FormaDePagoPactada
        VecCentroCta(i).O_CodigoArticulo = !R_CodigoArticulo
        VecCentroCta(i).Centro_Descripcion = BuscarDescCentro(!R_CentroDeCosto)
        VecCentroCta(i).Cta_Descripcion = BuscarDescCta(!R_CuentaContable)
        VecCentroCta(i).O_CantidadPedida = !O_CantidadPedida
        VecCentroCta(i).O_CentroDeCosto = !R_CentroDeCosto
        VecCentroCta(i).O_CuentaContable = !R_CuentaContable
        VecCentroCta(i).O_CantidadPendiente = !O_CantidadPendiente + !R_CantidadRecibida
        VecCentroCta(i).O_Precio = VerificarNulo(!R_PrecioOrdenDeCompra, "N")
        VecCentroCta(i).R_Precio = !R_PrecioRecepcionado
        VecCentroCta(i).CantRecibida = !R_CantidadRecibida
        VecCentroCta(i).FechaOrden = !FechaOrden
        i = i + 1
        .MoveNext
    Wend
        .Close
  
End With
    If LvListado.ListItems.Count > 0 Then
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    End If

    Call CalcularTotal
    'CmdCambiar.Visible = True
    FrameCant.Enabled = False
    TxtResp.Enabled = False
    TxtObs.Enabled = False
    TxtFactura.Enabled = False
    TxtLugar.Enabled = False
    CmdConfirnar.Visible = False
    CmdImprimir.Enabled = True
    CmdExpPdf.Enabled = True
    
    LvListado.Sorted = True

End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirnar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Crear la Recepcion de Productos?", vbYesNo)
    If Rta = vbYes Then
        Call GrabarRecepcion
    End If
End Sub

Private Function GrabarRecepcion() As Boolean
  Dim Sql As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  'Dim NroRecepcion As Integer
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

'On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    If Not ValidadIntegridad Then Exit Function
    
    GrabarRecepcion = True
    Conec.BeginTrans
       Sql = "SpOCRecepcionOrdenesDeCompraCabeceraAgregar " & _
             " @R_Fecha=" & FechaSQL(CStr(CalFecha.Value), "SQL") & _
             ", @R_CodigoProveedor = " & CStr(Proveedor) & _
             ", @R_LugarDeRecepcion = '" & TxtLugar.Text & _
             "', @U_Usuario ='" & Usuario & _
             "', @R_ResponsableRecepcion = '" & TxtResp.Text & _
             "', @R_Observaciones = '" & TxtObs.Text & _
             "', @R_Factura ='" & TxtFactura & "'"
     'graba el encabezado i retorna el Nro de Recepcion
        RsGrabar.Open Sql, Conec
        NroRecepcion = RsGrabar!R_NumeroRecepcion
        
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        If .CantRecibida > 0 Then
         Precio = Replace(CStr(.R_Precio), ",", ".")
         Sql = "SpOCRecepcionOrdenesDeCompraRenglonesAgregar " & _
               "  @R_NumeroRecepcion =" & NroRecepcion & _
               ", @R_CodigoArticulo =" & .O_CodigoArticulo & _
               ", @R_CuentaContable ='" & .O_CuentaContable & _
               "', @R_CentroDeCosto ='" & .O_CentroDeCosto & _
               "', @R_CantidadRecibida =" & Replace(.CantRecibida, ",", ".") & _
               ", @R_PrecioRecepcionado = " & Precio & _
               ", @O_NumeroOrdenDeCompra = " & .NroOrden & _
               ", @O_CentroDeCostoEmisor = '" & .O_CentroDeCostoEmisor & _
               "', @R_PrecioOrdenDeCompra = " & Replace(.O_Precio, ",", ".")
            
            Conec.Execute Sql
        End If
      End With
    Next
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       CmdConfirnar.Visible = False
       CmdCambiar.Visible = True
       CmdImprimir.Enabled = True
       
       Modificado = False

       FrmMensaje.LbMensaje = "La Recepcion se Grabó correctamente con el Nº: " + CStr(NroRecepcion) & _
                               Chr(13) & " ¿Que desea hacer?"
       FrmMensaje.Show vbModal
       
       Modificado = False
       If FrmMensaje.Retorno = vbimprimir Then
         Call ConfImpresionDeAutorizacion
         RepAutorizacionDePago.Show vbModal
       End If
         
       If FrmMensaje.Retorno = vbNuevo Then
         Call LimpiarRecepcion
       End If
       
       If FrmMensaje.Retorno = vbExportesPDF Then
          CmdExpPdf_Click
       End If
       
       If FrmMensaje.Retorno = vbCerrar Then
          Unload Me
       End If
   Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
   End If
  End If
End Function

Private Sub ModificarRecepcion()
  Dim Sql As String
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
       Sql = "SpOCRecepcionOrdenesDeCompraCabeceraModificar @R_NumeroRecepcion=" & CStr(NroRecepcion) & _
             ", @R_Fecha=" & FechaSQL(CStr(CalFecha.Value), "SQL") & _
             ", @R_CodigoProveedor = " & CStr(Proveedor) & _
             ", @R_LugarDeRecepcion = '" & TxtLugar.Text & _
             "', @U_Usuario ='" & Usuario & _
             "', @R_ResponsableRecepcion = '" & TxtResp.Text & _
             "', @R_Observaciones = '" & TxtObs.Text & _
             "', @R_Factura ='" & TxtFactura & "'"
             
       Conec.Execute Sql
        
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        
         Precio = Replace(CStr(.R_Precio), ",", ".")
         Sql = "SpOCRecepcionOrdenesDeCompraRenglonesModificar " & _
               " @R_NumeroRecepcion =" & CStr(NroRecepcion) & _
               ", @R_CodigoArticulo =" & CStr(.O_CodigoArticulo) & _
               ", @R_CuentaContable ='" & .O_CuentaContable & _
               "', @R_CentroDeCosto =" & CStr(.O_CentroDeCosto) & _
               ", @R_CantidadRecibida =" & CStr(.CantRecibida) & _
               ", @R_PrecioRecepcionado = " & Precio & _
               ", @O_NumeroOrdenDeCompra = " & CStr(.NroOrden)
            
            Conec.Execute Sql
        
      End With
    Next
    Conec.CommitTrans
    
ErrorInsert:
   If Err.Number = 0 Then
       CmdConfirnar.Visible = False
       CmdCambiar.Visible = False
       
       MsgBox "La Recepcion se Modificó correctamente"
       Modificado = False
    
      Rta = MsgBox("¿Desea realizar otra acción?", vbYesNo)
      
      If Rta = vbYes Then
         Call LimpiarRecepcion
      Else
         Unload Me
      End If

   Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
   End If
  End If
End Sub

Private Sub LimpiarRecepcion()
    'Call limpiarTXT(Me)
    TxtCantRecibida = ""
    TxtCantTotal = ""
    TxtLugar = ""
    TxtObs = ""
    TxtResp = ""
    txtTotal = "0"
    TxtFactura = ""
    TxtLugar.Enabled = True
    TxtObs.Enabled = True
    ReDim VecRecepcion(0)
    ReDim VecCentroCta(0)
    Proveedor = 0
    NroRecepcion = 0
    UsuarioResponsable = Usuario
    LvListado.ListItems.Clear
    LvCenCostoCtas.ListItems.Clear
    CmbProv.Enabled = True
    TxtFactura.Enabled = True
    CmbProv.ListIndex = 0
    
    CalFecha.Value = ValidarPeriodo(Date, False)
    CalFecha.Enabled = True
    CmdConfirnar.Visible = True
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    CmdCambiar.Visible = False
    CmdAnular.Visible = False
    Modificado = False
    TxtResp.Text = NombreUsuario
End Sub

Private Function ValidarEncabezado() As Boolean
Dim i As Integer
Dim Asignado As Boolean
    
    Asignado = False
    ValidarEncabezado = True
    
    If TxtLugar.Text = "" Then
        MsgBox "Debe ingresar el lugar de recepcion"
        TxtLugar.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
    
    'If TxtResp.Text = "" Then
    '    MsgBox "Debe ingresar un Responsable"
    '    TxtResp.SetFocus
    '    ValidarEncabezado = False
    '    Exit Function
    'End If
    
    For i = 1 To UBound(VecCentroCta)
        If VecCentroCta(i).CantRecibida > 0 Then
            Asignado = True
            Exit For
        End If
    Next
    If Not Asignado Then
        MsgBox "Debe recibir algún Artículo"
        LvCenCostoCtas.SetFocus
        ValidarEncabezado = False
    End If
    
End Function

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeAutorizacion
         RepAutorizacionDePago.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepAutorizacionDePago.Pages
         Unload RepAutorizacionDePago
  End If
Error:
    If Err.Number = 0 Then
        MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub


Private Sub CmdImprimir_Click()
    Call ConfImpresionDeAutorizacion
    RepAutorizacionDePago.Show
End Sub

Private Sub ConfImpresionDeAutorizacion()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "Articulo", adVarChar, 100
    RsListado.Fields.Append "Centro", adVarChar, 100
    RsListado.Fields.Append "CentroPadre", adVarChar, 50
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "Cantidad", adDouble
    RsListado.Fields.Append "Precio", adDouble
    RsListado.Fields.Append "Importe", adDouble
    RsListado.Fields.Append "NroOrden", adVarChar, 8
    RsListado.Fields.Append "FechaOrden", adVarChar, 10
    RsListado.Fields.Append "FormaDePago", adVarChar, 50
    RsListado.Open
    i = 1
    While i <= UBound(VecCentroCta)
    If VecCentroCta(i).CantRecibida > 0 Then
        RsListado.AddNew
        
        With VecCentroCta(i)
             RsListado!Articulo = BuscarDescArt(.O_CodigoArticulo, BuscarTablaCentroEmisor(.O_CentroDeCostoEmisor))
             RsListado!Centro = .Centro_Descripcion & " - Cód. " & BuscarCodigoCentro(.O_CentroDeCosto)
             RsListado!CentroPadre = BuscarDescCentroEmisor(.O_CentroDeCostoEmisor) ' BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
             RsListado!Cuenta = .Cta_Descripcion & " - Cód. " & .O_CuentaContable
             RsListado!Cantidad = ValN(.CantRecibida)
             RsListado!Precio = .R_Precio
             RsListado!Importe = .CantRecibida * .R_Precio
             RsListado!NroOrden = Format(.NroOrden, "00000000")
             RsListado!FechaOrden = .FechaOrden
             RsListado!FormaDePago = .O_FormaDePagoPactada
        End With
    End If
        i = i + 1
    Wend
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If

    RepAutorizacionDePago.TxtFecha = CStr(CalFecha.Value)
    RepAutorizacionDePago.TxtNroOrden.Text = Format(NroRecepcion, "000000000")
    RepAutorizacionDePago.BarNroAutorizacion.Caption = Format(NroRecepcion, "000000000")
    RepAutorizacionDePago.TxtLugarDeEntrega = TxtLugar.Text
    RepAutorizacionDePago.TxtProv.Text = CmbProv.Text & " (Cod. " & VecProveedores(CmbProv.ListIndex).Codigo & ")"
    RepAutorizacionDePago.TxtUsuario.Text = UsuarioResponsable
    RepAutorizacionDePago.TxtResp.Text = TxtResp.Text
    RepAutorizacionDePago.TxtAnulada.Visible = LBAnulada.Visible
    RepAutorizacionDePago.TxtAnulada.Text = LBAnulada.Caption
    RepAutorizacionDePago.TxtObservaciones = TxtObs
    RepAutorizacionDePago.TxtFactura = TxtFactura
    RepAutorizacionDePago.DataControl1.Recordset = RsListado
    RepAutorizacionDePago.Zoom = -1
End Sub


Private Sub CmdModifCant_Click()
'On Error GoTo Errores
  Dim i As Integer
  Dim PosVec As Integer
  'i = LvCenCostoCtas.SelectedItem.Index
   i = Val(LvListado.SelectedItem.ListSubItems(6))

'agrega al vector
   PosVec = Val(LvCenCostoCtas.ListItems(i).SubItems(5))
   If VecCentroCta(PosVec).O_CantidadPendiente >= Val(TxtCantRecibida.Text) Then
      Modificado = True
      VecCentroCta(PosVec).CantRecibida = Val(TxtCantRecibida.Text)
        'lo pone en el LV
      LvCenCostoCtas.ListItems(i).SubItems(4) = VecCentroCta(PosVec).CantRecibida
       If i < LvCenCostoCtas.ListItems.Count Then
          LvCenCostoCtas.ListItems(i + 1).Selected = True
          Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
       Else
           If LvListado.SelectedItem.Index < LvListado.ListItems.Count Then
              LvListado.ListItems(LvListado.SelectedItem.Index + 1).Selected = True
              Call LvListado_ItemClick(LvListado.SelectedItem)
           End If
       End If
    Else
        MsgBox "La cantidad Recibida no puede ser mayor a la pendiente"
    End If
    Call CalcularTotal
End Sub


Private Sub CmdNuevo_Click()
    Call LimpiarRecepcion
    TxtNroRecepcion = ""
    LBAnulada.Visible = False
    FrameCant.Enabled = True
    CmdModifCant.Enabled = False
    FechaMin = Date
End Sub

Private Sub CmdTraer_Click()
    A01_3100.P_Codigo = VecProveedores(CmbProv.ListIndex).Codigo
    'A01_3100.Periodo = CalFecha.Value
    
    A01_3100.Show vbModal
    Proveedor = VecProveedores(CmbProv.ListIndex).Codigo
  
  If UBound(Articulos) > 0 Then
     Call CrearRecepcion
     CmbProv.Enabled = False
     CmdTraer.Enabled = False
  End If
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    CalFecha.Value = Date
    UsuarioResponsable = Usuario
    TxtResp.Text = NombreUsuario
    Modificado = False
End Sub

Private Sub CmbProv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call CmdTraer_Click
    End If
End Sub

Private Sub TxtCantTotal_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtCantTotal, KeyAscii)
End Sub

Private Sub TxtNroRecepcion_LostFocus()
  If Val(TxtNroRecepcion.Text) <> NroRecepcion Then
    CmdConfirnar.Visible = TxtNroRecepcion.Text = ""
    CmdCambiar.Visible = TxtNroRecepcion.Text <> ""
    Call LimpiarRecepcion
  End If

End Sub

Private Sub CrearEncabezados()
    LvListado.ColumnHeaders.Add , , "Orden de Compra", 1600
    LvListado.ColumnHeaders.Add , , "Centro de Costo Emisor", (LvListado.Width - 5500) / 2
    LvListado.ColumnHeaders.Add , , "Descripción Artículo", (LvListado.Width - 5500) / 2
    LvListado.ColumnHeaders.Add , , "Cant. Pedida", 1250, 1
    LvListado.ColumnHeaders.Add , , "Cant. Pendiente", 1300, 1
    LvListado.ColumnHeaders.Add , , "Precio Unit. sin IVA", 1000, 1
    LvListado.ColumnHeaders.Add , , "Index", 0, 1
   
    
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos", (LvCenCostoCtas.Width - 4000) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Sub-Centros de Costos", (LvCenCostoCtas.Width - 2150) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", (LvCenCostoCtas.Width - 2500) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Cant. Pendiente", 1300, 1
    LvCenCostoCtas.ColumnHeaders.Add , , "Cant. Recibida", 1250, 1
    LvCenCostoCtas.ColumnHeaders.Add , , "Index vector", 0

End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
Dim Importe As String
    For i = 1 To UBound(VecCentroCta)
       Importe = CStr(VecCentroCta(i).CantRecibida * VecCentroCta(i).R_Precio)
        Total = Total + Val(Replace(Importe, ",", "."))
    Next
        txtTotal.Text = Format(Total, "$ 0.00##")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Rta As Integer
    If Modificado And NroRecepcion = 0 Then
       Rta = MsgBox("Ha efectuado cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            'If NroRecepcion = 0 Then
             If Not GrabarRecepcion Then
                Cancel = 1
             End If
            'Else
               'Call ModificarRecepcion
            'End If
         End If
       End If
    End If
End Sub

Private Sub LvCenCostoCtas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CmdModifCant.Enabled = VecCentroCta(Item.SubItems(5)).R_Precio <> 0
    If VecCentroCta(Item.SubItems(5)).CantRecibida <> 0 Then
        TxtCantRecibida.Text = Replace(VecCentroCta(Item.SubItems(5)).CantRecibida, ",", ".")
    Else
        TxtCantRecibida.Text = ""
    End If
End Sub

Private Sub Timer1_Timer()
   If NroRecepcion <> 0 Then
      TxtNroRecepcion.Text = CStr(NroRecepcion)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Errores
  Call CargarLvCenCostoCtas(VecRecepcion(Val(LvListado.SelectedItem.SubItems(6))).A_Codigo, VecRecepcion(Val(LvListado.SelectedItem.SubItems(6))).O_NumeroOrdenDeCompra, VecRecepcion(Val(LvListado.SelectedItem.SubItems(6))).O_CentroDeCostoEmisor)
    LbDetalle.Caption = "Orden de Compra Nº: " + LvListado.SelectedItem.Text + _
    " Artículo: " + LvListado.SelectedItem.SubItems(1)
    CmdAsignarCant.Enabled = True
    TxtCantTotal.Text = ""
Errores:
    Call ManipularError(Err.Number, Err.Description, Timer1)
End Sub

Private Sub CargarLvCenCostoCtas(A_Codigo As Long, NroOrden As Integer, CentroEmisor As String)
  Dim i As Integer
    LvCenCostoCtas.ListItems.Clear
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        If .O_CodigoArticulo = A_Codigo And .NroOrden = NroOrden And .O_CentroDeCostoEmisor = CentroEmisor Then
            LvCenCostoCtas.ListItems.Add
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Text = BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(1) = .Centro_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(2) = .Cta_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(3) = Format(.O_CantidadPendiente, "0.00")
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(4) = Format(.CantRecibida, "0.00")
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(5) = i
        End If
      End With
    Next
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Public Sub BuscarArt(Codigo As Integer, CmbArt As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = Codigo Then
           Exit For
        End If
   Next
   If i < UBound(VecArtCompra) Then
      CmbArt.ListIndex = i
   End If
End Sub

Private Sub TxtCantRecibida_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtCantRecibida, KeyAscii)
End Sub

Private Sub TxtLugar_Change()
    Modificado = True
End Sub

Private Sub TxtNroRecepcion_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii = 13 Then
       Call CmdCargar_Click
    Else
       If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub CrearRecepcion()
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    j = 1
 With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    ReDim VecRecepcion(0)
    ReDim VecCentroCta(0)
    
    LvListado.ListItems.Clear
    LvCenCostoCtas.ListItems.Clear
    LvListado.Sorted = False

    
  For i = 1 To UBound(Articulos)
    Sql = "SpOcOrdenesPendientesDeArticuloTraer @P_Codigo=" & Proveedor & _
               ", @A_Codigo=" & Articulos(i)
    
    .Open Sql, Conec
    ReDim Preserve VecRecepcion(j + .RecordCount)
    While Not .EOF
        VecRecepcion(j).A_Codigo = Articulos(i)
        VecRecepcion(j).A_Descripcion = BuscarDescArt(Articulos(i), BuscarTablaCentroEmisor(!O_CentroDeCostoEmisor))
        VecRecepcion(j).O_CantidadPedida = !O_CantidadPedida
        VecRecepcion(j).O_CantidadPendiente = !O_CantidadPendiente
        VecRecepcion(j).O_NumeroOrdenDeCompra = !O_NumeroOrdenDeCompra
        VecRecepcion(j).O_PrecioPactado = !O_PrecioPactado
        VecRecepcion(j).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        VecRecepcion(j).R_Precio = VecRecepcion(j).O_PrecioPactado
        VecRecepcion(j).O_FormaDePagoPactada = !O_FormaDePagoPactada
        j = j + 1
        FechaMin = IIf(FechaMin < !O_Fecha, !O_Fecha, FechaMin)
        .MoveNext
    Wend
    
    .Close
 Next
   For i = 1 To UBound(VecRecepcion) - 1
        LvListado.ListItems.Add
        LvListado.ListItems(i).Text = Format(VecRecepcion(i).O_NumeroOrdenDeCompra, "000000000")
        LvListado.ListItems(i).SubItems(1) = BuscarDescCentroEmisor(VecRecepcion(i).O_CentroDeCostoEmisor)
        LvListado.ListItems(i).SubItems(2) = VecRecepcion(i).A_Descripcion
        LvListado.ListItems(i).SubItems(3) = Format(VecRecepcion(i).O_CantidadPedida, "0.00")
        LvListado.ListItems(i).SubItems(4) = Format(VecRecepcion(i).O_CantidadPendiente, "0.00")
        LvListado.ListItems(i).SubItems(5) = Format(VecRecepcion(i).O_PrecioPactado, "0.00##")
        LvListado.ListItems(i).SubItems(6) = i
   Next
      i = 1
   For j = 1 To UBound(VecRecepcion) - 1
        Sql = "SpOcOrdenesPendientesDeCentrosCtas @O_NumeroOrdenDeCompra=" & VecRecepcion(j).O_NumeroOrdenDeCompra & _
                                              " , @O_CentroDeCostoEmisor='" & VecRecepcion(j).O_CentroDeCostoEmisor & _
                                              "', @A_Codigo=" & VecRecepcion(j).A_Codigo
       .Open Sql, Conec
        ReDim Preserve VecCentroCta(i + .RecordCount - 1)
    While Not .EOF
        VecCentroCta(i).NroOrden = VecRecepcion(j).O_NumeroOrdenDeCompra
        VecCentroCta(i).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        VecCentroCta(i).O_CodigoArticulo = VecRecepcion(j).A_Codigo
        VecCentroCta(i).O_FormaDePagoPactada = VecRecepcion(j).O_FormaDePagoPactada
        VecCentroCta(i).Centro_Descripcion = !Centro
        VecCentroCta(i).Cta_Descripcion = BuscarDescCta(!CodCta)
        VecCentroCta(i).O_CantidadPedida = !O_CantidadPedida
        VecCentroCta(i).O_CentroDeCosto = !CodCentro
        VecCentroCta(i).O_CuentaContable = !CodCta
        VecCentroCta(i).O_CantidadPendiente = !O_CantidadPendiente
        VecCentroCta(i).O_Precio = !O_PrecioPactado
        VecCentroCta(i).R_Precio = !O_PrecioPactado
        VecCentroCta(i).FechaOrden = !O_Fecha
        i = i + 1
        .MoveNext
    Wend
        .Close
   Next
  End With
    If LvListado.ListItems.Count > 0 Then
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    End If
    
    LvListado.Sorted = True

End Sub

Private Function ValidadIntegridad() As Boolean
    Dim i As Integer
    Dim Sql As String
    Dim RsValidar As New ADODB.Recordset

On Error GoTo ErrorIntegridad
    ValidadIntegridad = True
    For i = 1 To UBound(VecCentroCta)
        With VecCentroCta(i)
            If .CantRecibida > 0 Then
                Sql = "SpOcValidarRecepcion @O_NumeroOrdenDeCompra =" & .NroOrden & _
                                         ", @O_CodigoArticulo=" & .O_CodigoArticulo & _
                                         ", @O_CuentaContable ='" & .O_CuentaContable & _
                                        "', @O_CentroDeCosto ='" & .O_CentroDeCosto & _
                                        "', @O_CentroDeCostoEmisor ='" & .O_CentroDeCostoEmisor & "'"
                
                If RsValidar.State = adStateOpen Then RsValidar.Close
                
                RsValidar.Open Sql, Conec
                If RsValidar.EOF Then
                    MsgBox "El Registro a sido Borrado de la Orden de Compra Por Otro Usuario", vbCritical
                    ValidadIntegridad = False
                    Exit Function
                End If
                
                If Not IsNull(RsValidar!FechaAnulacion) Then
                    MsgBox "La Orden de Compras Nº " & .NroOrden & " a sido anulada por otro Usuario", vbCritical
                    ValidadIntegridad = False
                    Exit Function
                End If
                
                If RsValidar!O_CantidadPendiente <> .O_CantidadPendiente Then
                    MsgBox "Alguna cantidad a sido Modificada por otro Usuario", vbCritical
                    ValidadIntegridad = False
                    Exit Function
              End If
            End If
        End With
    Next
ErrorIntegridad:
        If Err.Number <> 0 Then
            Call ManipularError(Err.Number, Err.Description)
            ValidadIntegridad = False
        End If
End Function
