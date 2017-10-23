VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_6200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificación de Servicios"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPrecioOrden 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10620
      TabIndex        =   27
      Top             =   1800
      Width           =   1230
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4680
      TabIndex        =   25
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Autorización"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7380
      TabIndex        =   11
      Top             =   4815
      Width           =   1815
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   3375
      TabIndex        =   9
      Top             =   4815
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   6030
      TabIndex        =   10
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   9315
      TabIndex        =   12
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10755
      TabIndex        =   14
      Top             =   4800
      Width           =   1230
   End
   Begin VB.Frame FramePrecio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingreso de Precio"
      Height          =   1365
      Left            =   10440
      TabIndex        =   19
      Top             =   2295
      Width           =   1545
      Begin VB.CommandButton CmdCambiarPrecio 
         Caption         =   "Ingresar Precio"
         Enabled         =   0   'False
         Height          =   350
         Left            =   135
         TabIndex        =   8
         Top             =   900
         Width           =   1300
      End
      Begin VB.TextBox TxtPrecioU 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio sin IVA:"
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
         Left            =   150
         TabIndex        =   20
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   4680
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
      Left            =   11025
      TabIndex        =   18
      Text            =   "0"
      Top             =   4350
      Width           =   960
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   3210
      Left            =   45
      TabIndex        =   6
      Top             =   1530
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   5662
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
      Caption         =   "Datos de la Certificación de Servicios"
      Height          =   1455
      Left            =   45
      TabIndex        =   15
      Top             =   45
      Width           =   11940
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   23
         Top             =   1035
         Width           =   10275
      End
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5355
         TabIndex        =   5
         Top             =   630
         Width           =   1050
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   3150
         TabIndex        =   1
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4275
         TabIndex        =   2
         Top             =   225
         Width           =   1000
      End
      Begin VB.TextBox TxtNroAutorizacion 
         Height          =   315
         Left            =   1845
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   10530
         TabIndex        =   3
         Top             =   225
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   108199939
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1125
         TabIndex        =   4
         Top             =   630
         Width           =   4155
         _ExtentX        =   7329
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "F. Imputación:"
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
         Left            =   9270
         TabIndex        =   28
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label16 
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
         Left            =   135
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
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
         Left            =   5445
         TabIndex        =   22
         Top             =   270
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
         Left            =   135
         TabIndex        =   21
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Autorizacion:"
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
         Left            =   135
         TabIndex        =   16
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   9315
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio Orden:"
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
      Left            =   10620
      TabIndex        =   26
      Top             =   1575
      Width           =   1185
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
      Left            =   10485
      TabIndex        =   17
      Top             =   4410
      Width           =   510
   End
End
Attribute VB_Name = "A01_6200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Proveedor As Integer
'este vector se aparea con el Lv
Private VecCentroCta() As TipoAutorizacionDePago
Private Modificado As Boolean
Public NroAutorizacion As Long
Private FechaMin As Date
Private MontoSinPres As Double
Private TotalMontoSinPres As Double

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalFecha_GotFocus()
    CalFecha.Format = dtpShortDate
End Sub

Private Sub CalFecha_Validate(Cancel As Boolean)
    If CalFecha.Value < FechaMin Then
        MsgBox "La fecha debe posterior al " & FechaMin, vbInformation, "Fecha inválida"
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
    Rta = MsgBox("¿Está seguro de que desea Anular la Autorización de pago?", vbYesNo)
    If Rta = vbYes Then
        Sql = "SpOCAutorizacionesDePagoCabeceraAnular @A_NumeroDeAutorizacionDePago =" & NroAutorizacion
        Conec.Execute Sql
        MsgBox "La Autorización de Pago se Anuló correctamente", vbInformation
    Else
        Exit Sub
    End If
Error:

  If Err.Number <> 0 Then
     Call ManipularError(Err.Number, Err.Description)
  Else
     Rta = MsgBox("¿Desea realizar otra acción?", vbYesNo)
     If Rta = vbYes Then
        Call LimpiarAutorizacion
     Else
        Unload Me
     End If
  End If
End Sub

Private Sub CMDBuscar_Click()
    Unload BuscarAutorizacionDePago
    BuscarAutorizacionDePago.Show vbModal
    Timer1.Enabled = True
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar la Recepción?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarAutorizacion
    End If
End Sub

Private Sub CmdCambiarPrecio_Click()
  Dim i As Integer

    i = LvCenCostoCtas.SelectedItem.Index
    TotalMontoSinPres = TotalMontoSinPres - VecCentroCta(i).MontoSinPresupuestar
    If ValidarImporte Then
        VecCentroCta(i).PrecioReal = Val(TxtPrecioU.Text)
        MontoSinPres = Val(TxtPrecioU.Text) - VecCentroCta(i).O_PrecioPactado
        VecCentroCta(i).MontoSinPresupuestar = IIf(MontoSinPres > 0, MontoSinPres, 0)
        LvCenCostoCtas.SelectedItem.SubItems(4) = Format(Val(TxtPrecioU.Text), "0.00##")
        Modificado = True
        Call CalcularTotal
    End If
    TotalMontoSinPres = TotalMontoSinPres + VecCentroCta(i).MontoSinPresupuestar
End Sub

Private Function ValidarImporte() As Boolean
    Dim i As Integer
    Dim Rta As Integer
    Dim Sql As String
    Dim MontoDisponible As Double
    Dim RsValidar As New ADODB.Recordset
    ValidarImporte = True
    i = LvCenCostoCtas.SelectedItem.Index
    MontoSinPres = 0
    
    If VecCentroCta(i).O_PrecioPactado <> Val(TxtPrecioU.Text) Then
        Rta = MsgBox("El precio no coincide con el de la Orden de Contratación Emitida, Confirma el ingreso a ese precio?", vbYesNo, "Precio $" & VecCentroCta(LvCenCostoCtas.SelectedItem.Index).O_PrecioPactado)
        If Rta = vbNo Then
            ValidarImporte = False
            Exit Function
        End If
    
        Sql = " SpOCImporteSinPresupuestarAutorizacionDeCargaContable @CentroDeCosto ='" & VecCentroCta(i).O_CentroDeCostoEmisor & _
                         "', @NroAutorizacion =" & NroAutorizacion & _
                         " , @Periodo=" & FechaSQL(CalFecha, "SQL")
                         
        RsValidar.Open Sql, Conec
        MontoSinPres = Val(TxtPrecioU.Text) - VecCentroCta(i).O_PrecioPactado
        MontoDisponible = RsValidar!MontoSinPresupuestarMensual - RsValidar!MontoSinPres
        If MontoDisponible < MontoSinPres Then
            MsgBox "El monto disponible para este período es de $" & Format(MontoDisponible, "0.00"), vbInformation
            ValidarImporte = False
            Exit Function
        End If
            
    End If
End Function

Private Sub CmdCargar_Click()
    Call CargarAutorizacionPago(Val(TxtNroAutorizacion))
    Modificado = False

End Sub

Private Sub CargarAutorizacionPago(NroAutorizacion As Long)
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
On Error GoTo Error
    LBAnulada.Visible = False

    j = 1
 With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    LvCenCostoCtas.ListItems.Clear
    
    Sql = "SpOCAutorizacionesDePagoCabeceraTraerNro " & _
                  "@NroAutorizacion = " & NroAutorizacion & _
                ", @Usuario = '" & Usuario & "'"
    
    .Open Sql, Conec
     j = 1
    If .EOF Then
        MsgBox "No existe una Autorización de Pago con ese número"
        Call CmdNuevo_Click
        Exit Sub
    End If

    If Not IsNull(!A_FechaAnulacion) Then
        If Not IsNull(!A_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!A_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
        
        FramePrecio.Enabled = False
        CmdCambiar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
    Else
        LBAnulada.Visible = False
        FramePrecio.Enabled = True
        CmdCambiar.Enabled = True
        CmbProv.Enabled = True
        CalFecha.Enabled = True
        CmdAnular.Visible = True
        TxtObs.Enabled = True
    End If
    
    If VerificarNulo(!A_Seguro, "B") = True Then
        CmdCambiar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
        FramePrecio.Enabled = False
        MsgBox "La Certificación de servicio no puede ser Modificada por se de Seguros", vbInformation
    End If

     CalFecha.Value = !A_Fecha
     CalFecha.Format = dtpShortDate
     
     Proveedor = !A_CodigoProveedor
     CmbProv.Text = BuscarDescProv(!A_CodigoProveedor)
     CmbProv.Enabled = False
     CmdTraer.Enabled = False
     TxtObs.Text = VerificarNulo(!A_Observaciones)
     TxtNroAutorizacion.Text = Format(NroAutorizacion, "0000000000")
     
     Me.NroAutorizacion = NroAutorizacion
     .Close
     
     Sql = "SpOCAutorizacionesDePagoRenglonesTraer @A_NumeroDeAutorizacionDePago =" & NroAutorizacion
     .Open Sql, Conec
     ReDim VecCentroCta(.RecordCount)
    i = 1
    While Not .EOF
        VecCentroCta(i).O_NumeroOrdenDeContratacion = !O_NumeroDeOrdenDeContratacionDeServicios
        VecCentroCta(i).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        VecCentroCta(i).O_CentroDeCosto = !A_CentroDeCosto
        VecCentroCta(i).O_CuentaContable = !A_CuentaContable
        VecCentroCta(i).O_PrecioPactado = !O_ImporteOrdenDeContratacion
        VecCentroCta(i).PrecioReal = !A_Importe
        VecCentroCta(i).MontoSinPresupuestar = !A_ImporteSinPresupuestar
        
        LvCenCostoCtas.ListItems.Add
        LvCenCostoCtas.ListItems(i).Text = Format(!O_NumeroDeOrdenDeContratacionDeServicios, "0000000000")
        LvCenCostoCtas.ListItems(i).SubItems(1) = BuscarDescCentroEmisor(!O_CentroDeCostoEmisor)
        LvCenCostoCtas.ListItems(i).SubItems(2) = "Centro: " & BuscarDescCentroEmisor(BuscarCentroPadre(VecCentroCta(i).O_CentroDeCosto)) & " - " & BuscarDescCentro(!A_CentroDeCosto)
        LvCenCostoCtas.ListItems(i).SubItems(3) = BuscarDescCta(!A_CuentaContable)
        LvCenCostoCtas.ListItems(i).SubItems(4) = Format(!A_Importe, "0.00##")
        i = i + 1
        .MoveNext
    Wend
        .Close
    LvCenCostoCtas.ListItems(1).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
End With
    Call CalcularTotal
    TotalMontoSinPres = CalcularTotalSinPresupuestar
    CmdCambiar.Visible = True
    CmdConfirnar.Visible = False
    CmdImprimir.Enabled = True
    CmdExpPdf.Enabled = True
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirnar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Crear la Autorizacion de Pago?", vbYesNo)
    If Rta = vbYes Then
        Call GrabarAutorizacion
    End If
End Sub

Private Sub GrabarAutorizacion()
  Dim Sql As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
       Sql = "SpOCAutorizacionesDePagoCabeceraAgregar " & _
                    "  @A_Fecha= " & FechaSQL(CStr(CalFecha.Value), "SQL") & _
                    ", @A_CodigoProveedor = " & Proveedor & _
                    ", @U_Usuario ='" & Usuario & _
                    "', @A_Observaciones = '" & TxtObs.Text & _
                    "', @A_FechaEmision = " & FechaSQL(CalFecha, "SQL")
          
     'graba el encabezado y retorna el Nro de Autorizacion
        RsGrabar.Open Sql, Conec
        NroAutorizacion = RsGrabar!A_NumeroDeAutorizacionDePago
        
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        If .PrecioReal <> 0 Then
         Precio = Replace(.PrecioReal, ",", ".")
         Sql = "SpOCAutorizacionesDePagoRenglonesAgregar " & _
                     " @A_NumeroDeAutorizacionDePago = " & NroAutorizacion & _
                    ", @A_CuentaContable ='" & .O_CuentaContable & _
                   "', @A_CentroDeCosto ='" & .O_CentroDeCosto & _
                    "', @A_Importe =" & Replace(.PrecioReal, ",", ".") & _
                    ", @O_NumeroDeOrdenDeContratacionDeServicios =" & .O_NumeroOrdenDeContratacion & _
                    ", @O_CentroDeCostoEmisor ='" & .O_CentroDeCostoEmisor & _
                    "', @O_ImporteOrdenDeContratacion = " & Replace(.O_PrecioPactado, ",", ".") & _
                    ", @A_ImporteSinPresupuestar = " & Replace(.MontoSinPresupuestar, ",", ".")
            
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
       
      'Rta = MsgBox("La Autorización de Pago se Grabó correctamente con el Nº: " & CStr(NroAutorizacion) & " ¿Desea imprimirla la Autorización?", vbYesNo)
       Modificado = False

       FrmMensaje.LbMensaje = "La Autorización de Pago se Grabó correctamente con el Nº: " + CStr(NroAutorizacion) & _
                               Chr(13) & " ¿Que desea hacer?"
       FrmMensaje.Show vbModal
       
       Modificado = False
       If FrmMensaje.Retorno = vbimprimir Then
         Call ConfImpresionDeAutorizacion
         RepAutorizacionDePagoServicio.Show vbModal
       End If
         
       If FrmMensaje.Retorno = vbNuevo Then
         Call LimpiarAutorizacion
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
End Sub

Private Sub ModificarAutorizacion()
  Dim Sql As String
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
       Sql = "SpOCAutorizacionesDePagoCabeceraModificar @A_NumeroDeAutorizacionDePago =" & NroAutorizacion & _
                    ", @A_Fecha= " & FechaSQL(CStr(CalFecha.Value), "SQL") & _
                    ", @A_CodigoProveedor = " & Proveedor & _
                    ", @U_Usuario ='" & Usuario & _
                    "', @A_Observaciones = '" & TxtObs.Text & _
                    "', @A_FechaEmision = " & FechaSQL(CalFecha, "SQL")
      Conec.Execute Sql
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        'If .PrecioReal > 0 Then
         Precio = Replace(.PrecioReal, ",", ".")
         Sql = "SpOCAutorizacionesDePagoRenglonesModificar " & _
                     " @A_NumeroDeAutorizacionDePago = " & NroAutorizacion & _
                    ", @A_CuentaContable ='" & .O_CuentaContable & _
                   "', @A_CentroDeCosto ='" & .O_CentroDeCosto & _
                   "', @A_Importe =" & Replace(.PrecioReal, ",", ".") & _
                    ", @O_NumeroDeOrdenDeContratacionDeServicios =" & .O_NumeroOrdenDeContratacion & _
                    ", @O_CentroDeCostoEmisor ='" & .O_CentroDeCostoEmisor & _
                   "', @O_ImporteOrdenDeContratacion = " & Replace(.O_PrecioPactado, ",", ".") & _
                    ", @A_ImporteSinPresupuestar = " & Replace(.MontoSinPresupuestar, ",", ".")
            
            Conec.Execute Sql
        'End If
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
         Call LimpiarAutorizacion
      Else
         Unload Me
      End If

   Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
   End If
  End If
End Sub

Private Sub LimpiarAutorizacion()
    Call LimpiarTXT(Me)
    ReDim VecRecepcion(0)
    ReDim VecCentroCta(0)
    Proveedor = 0
    NroAutorizacion = 0
    LvCenCostoCtas.ListItems.Clear
    CmbProv.Enabled = True
    CmbProv.ListIndex = 0
    
    CalFecha.Value = ValidarPeriodo(Date, False)
    CalFecha.Format = dtpCustom
    CalFecha.CustomFormat = " "

    CalFecha.Enabled = True
    CmdConfirnar.Visible = True
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    CmdCambiar.Visible = False
    CmdAnular.Visible = False
    Modificado = False
End Sub

Private Function ValidarEncabezado() As Boolean
    
    ValidarEncabezado = True
   ' Dim Sql As String
   ' Dim RsValidarPeriodo As ADODB.Recordset
   ' Set RsValidarPeriodo = New ADODB.Recordset
   ' Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(CalFecha.Value, "MM/yyyy")) & "'"
   ' RsValidarPeriodo.Open Sql, Conec
   ' If RsValidarPeriodo!Cerrado > 0 Then
   '    MsgBox "El período está Cerrado", vbExclamation, "Período Cerrado"
   '    CalFecha.SetFocus
   '    ValidarEncabezado = False
   '    Exit Function
   ' End If
    If CalFecha.Format = dtpCustom Then
        MsgBox "Debe ingresar Fecha de Imputación"
        CalFecha.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If

    If Val(Replace(txtTotal.Text, ",", ".")) = 0 Then
       MsgBox "La autorización de pago debe tener un importe mayor a 0", vbExclamation, "Total"
       LvCenCostoCtas.SetFocus
       ValidarEncabezado = False
       Exit Function
        
    End If
End Function

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeAutorizacion
         RepAutorizacionDePagoServicio.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepAutorizacionDePagoServicio.Pages
         Unload RepAutorizacionDePagoServicio
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
    RepAutorizacionDePagoServicio.Show
End Sub

Private Sub ConfImpresionDeAutorizacion()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    RsListado.Fields.Append "Orden", adVarChar, 10
    RsListado.Fields.Append "CentroEmisor", adVarChar, 100
    RsListado.Fields.Append "CentroPadre", adVarChar, 100
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "Centro", adVarChar, 100
    RsListado.Fields.Append "Importe", adDouble
    RsListado.Open
    i = 1
    For i = 1 To UBound(VecCentroCta)
        RsListado.AddNew
      With VecCentroCta(i)
        RsListado!Orden = Format(.O_NumeroOrdenDeContratacion, "0000000000")
        RsListado!CentroEmisor = BuscarDescCentroEmisor(.O_CentroDeCostoEmisor)
        RsListado!CentroPadre = BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
        RsListado!Centro = BuscarDescCentro(.O_CentroDeCosto) & " - Cód. " & BuscarCodigoCentro(.O_CentroDeCosto)
        RsListado!Cuenta = BuscarDescCta(.O_CuentaContable) & " - Cód. " & .O_CuentaContable
        RsListado!Importe = .PrecioReal
      End With
    Next
    
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    
    TxtNroAutorizacion.Text = Format(NroAutorizacion, "0000000000")
     
    RepAutorizacionDePagoServicio.TxtFecha = CStr(CalFecha.Value)
    RepAutorizacionDePagoServicio.TxtNroOrden.Text = TxtNroAutorizacion.Text
    RepAutorizacionDePagoServicio.BarNroAutorizacion.Caption = Format(TxtNroAutorizacion, "000000000")
    RepAutorizacionDePagoServicio.TxtProv.Text = CmbProv.Text & " (Cod. " & Proveedor & ")"
    RepAutorizacionDePagoServicio.TxtAnulada.Visible = LBAnulada.Visible
    RepAutorizacionDePagoServicio.TxtAnulada.Text = LBAnulada.Caption
    RepAutorizacionDePagoServicio.TxtObservaciones.Text = TxtObs.Text
    RepAutorizacionDePagoServicio.Zoom = -1
    RepAutorizacionDePagoServicio.DataControl1.Recordset = RsListado
End Sub

Private Sub CmdNuevo_Click()
    Call LimpiarAutorizacion
    FramePrecio.Enabled = True
    LBAnulada.Visible = False
    FechaMin = Date
    CmdCambiarPrecio.Enabled = False
End Sub

Private Sub CmdTraer_Click()
    ContratacionesSinPagar.P_Codigo = VecProveedores(CmbProv.ListIndex).Codigo
    'ContratacionesSinPagar.Periodo = CalFecha.Value
    ContratacionesSinPagar.Show vbModal
    Proveedor = VecProveedores(CmbProv.ListIndex).Codigo
  
  If UBound(VecAutorizacionDePago) > 0 Then
    Call CrearAutorizacionDePago
     Call CalcularTotal
    CmbProv.Enabled = False
    CmdTraer.Enabled = False
  End If
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    CalFecha.Value = Date 'ValidarPeriodo(Date, False)
    Modificado = False
End Sub

Private Sub CmbProv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call CmdTraer_Click
    End If
End Sub

Private Sub TxtNroAutorizacion_LostFocus()
  If Val(TxtNroAutorizacion.Text) <> NroAutorizacion Then
    CmdConfirnar.Visible = TxtNroAutorizacion.Text = ""
    CmdCambiar.Visible = TxtNroAutorizacion.Text <> ""
    Call LimpiarAutorizacion
  End If

End Sub

Private Sub TxtPrecioU_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdCambiarPrecio_Click
    End If
    Call TxtNumericoNeg(TxtPrecioU, KeyAscii)
End Sub

Private Sub CrearEncabezados()
    LvCenCostoCtas.ColumnHeaders.Add , , "Nº de Orden de Contratación", 1350
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos Emisor", 2000
    LvCenCostoCtas.ColumnHeaders.Add , , "Sub-Centros de Costos", 3450
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", 2000
    LvCenCostoCtas.ColumnHeaders.Add , , "Precio sin IVA", 1200, 1

End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
Dim Importe As String
    For i = 1 To UBound(VecCentroCta)
       Importe = CStr(VecCentroCta(i).PrecioReal)
        Total = Total + Val(Replace(Importe, ",", "."))
    Next
        txtTotal.Text = Format(Total, "0.00##")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Rta As Integer
    If Modificado Then
       Rta = MsgBox("Ha efectuado cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            If NroAutorizacion = 0 Then
                Call GrabarAutorizacion
            Else
                Call ModificarAutorizacion
            End If
         End If
       End If
    End If
End Sub

Private Sub LvCenCostoCtas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TxtPrecioU.Text = Replace(VecCentroCta(Item.Index).PrecioReal, ",", ".")
    TxtPrecioOrden.Text = Replace(VecCentroCta(Item.Index).O_PrecioPactado, ",", ".")
    CmdCambiarPrecio.Enabled = True
End Sub

Private Sub Timer1_Timer()
   If NroAutorizacion <> 0 Then
      TxtNroAutorizacion.Text = CStr(NroAutorizacion)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub TxtNroAutorizacion_KeyPress(KeyAscii As Integer)
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

Private Sub CrearAutorizacionDePago()
    Dim i As Integer
    Dim j As Integer
    ReDim VecCentroCta(UBound(VecAutorizacionDePago))
    
    LvCenCostoCtas.ListItems.Clear
    
  For i = 1 To UBound(VecAutorizacionDePago)
      VecCentroCta(i) = VecAutorizacionDePago(i)
      VecCentroCta(i).PrecioReal = VecAutorizacionDePago(i).O_PrecioPactado
      FechaMin = IIf(FechaMin < VecCentroCta(i).O_Fecha, VecCentroCta(i).O_Fecha, FechaMin)

      LvCenCostoCtas.ListItems.Add
      LvCenCostoCtas.ListItems(i).Text = Format(VecCentroCta(i).O_NumeroOrdenDeContratacion, "0000000000")
      LvCenCostoCtas.ListItems(i).SubItems(1) = BuscarDescCentroEmisor(VecAutorizacionDePago(i).O_CentroDeCostoEmisor)
      LvCenCostoCtas.ListItems(i).SubItems(2) = "Centro: " & BuscarDescCentroEmisor(BuscarCentroPadre(VecAutorizacionDePago(i).O_CentroDeCosto)) & " - " & BuscarDescCentro(VecAutorizacionDePago(i).O_CentroDeCosto)
      LvCenCostoCtas.ListItems(i).SubItems(3) = BuscarDescCta(VecAutorizacionDePago(i).O_CuentaContable)
      LvCenCostoCtas.ListItems(i).SubItems(4) = VecAutorizacionDePago(i).O_PrecioPactado
  Next
    LvCenCostoCtas.ListItems(1).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)

End Sub

Private Function CalcularTotalSinPresupuestar() As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        Total = Total + VecCentroCta(i).MontoSinPresupuestar
    Next
       
    CalcularTotalSinPresupuestar = Total
End Function
