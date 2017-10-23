VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_4500 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificación de Fondo Fijo"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIVA 
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
      Height          =   315
      Left            =   9360
      TabIndex        =   17
      Top             =   5535
      Width           =   960
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Height          =   1410
      Left            =   60
      TabIndex        =   33
      Top             =   5040
      Width           =   9240
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   5670
         MaxLength       =   4
         TabIndex        =   12
         Top             =   795
         Width           =   645
      End
      Begin VB.CommandButton CmdBuscarSubCentro 
         Caption         =   "Buscar"
         Height          =   300
         Left            =   5580
         TabIndex        =   10
         Top             =   382
         Width           =   735
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7965
         TabIndex        =   14
         Top             =   180
         Width           =   1150
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7965
         TabIndex        =   16
         Top             =   990
         Width           =   1150
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   7965
         TabIndex        =   15
         Top             =   585
         Width           =   1150
      End
      Begin Controles.ComboEsp CmbCentrosDeCostos 
         Height          =   315
         Left            =   1980
         TabIndex        =   9
         Top             =   375
         Width           =   3525
         _ExtentX        =   6218
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
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   315
         Left            =   1980
         TabIndex        =   11
         Top             =   795
         Width           =   3210
         _ExtentX        =   5662
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
      Begin VB.TextBox TxtPrecioU 
         Height          =   315
         Left            =   6615
         TabIndex        =   13
         Top             =   795
         Width           =   1275
      End
      Begin VB.Label LbCodCuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod."
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
         Left            =   5220
         TabIndex        =   37
         Top             =   855
         Width           =   405
      End
      Begin VB.Label LBCC 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sub-Centro de Costo:"
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
         TabIndex        =   36
         Top             =   435
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio sin IVA"
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
         Left            =   6615
         TabIndex        =   35
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label LbCta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta Contable:"
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
         Left            =   450
         TabIndex        =   34
         Top             =   855
         Width           =   1485
      End
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3030
      TabIndex        =   19
      Top             =   6525
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Autorización"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5745
      TabIndex        =   21
      Top             =   6495
      Width           =   1815
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   1725
      TabIndex        =   18
      Top             =   6515
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   4380
      TabIndex        =   20
      Top             =   6515
      Width           =   1230
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   7650
      TabIndex        =   22
      Top             =   6515
      Width           =   1320
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   9105
      TabIndex        =   24
      Top             =   6515
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   6480
   End
   Begin VB.TextBox TxtTotal 
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
      Left            =   9360
      TabIndex        =   29
      Text            =   "0"
      Top             =   6105
      Width           =   960
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   3210
      Left            =   60
      TabIndex        =   8
      Top             =   1845
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
      Height          =   1740
      Left            =   60
      TabIndex        =   25
      Top             =   45
      Width           =   10320
      Begin VB.OptionButton OptRecepcionada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Recepcionar"
         Height          =   240
         Index           =   1
         Left            =   7335
         TabIndex        =   44
         Top             =   1035
         Width           =   1590
      End
      Begin VB.OptionButton OptRecepcionada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recepcionado"
         Height          =   240
         Index           =   0
         Left            =   5850
         TabIndex        =   43
         Top             =   1035
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.TextBox TxtTotalFF 
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
         Left            =   4500
         TabIndex        =   40
         Text            =   "0"
         Top             =   998
         Width           =   1230
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1350
         Width           =   8700
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
         Left            =   8955
         TabIndex        =   3
         Top             =   225
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   52428803
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1530
         TabIndex        =   4
         Top             =   630
         Width           =   4200
         _ExtentX        =   7408
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
      Begin MSComCtl2.DTPicker CalFechaEmitida 
         Height          =   330
         Left            =   8955
         TabIndex        =   5
         Top             =   585
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbFondoFijoNro 
         Height          =   330
         Left            =   1530
         TabIndex        =   6
         Top             =   990
         Width           =   1905
         _ExtentX        =   3360
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total F. F.:"
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
         Left            =   3465
         TabIndex        =   41
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fondo Fijo Nº:"
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
         Left            =   255
         TabIndex        =   39
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "F. Emitida:"
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
         Left            =   8010
         TabIndex        =   38
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período Imputación:"
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
         Left            =   7185
         TabIndex        =   26
         Top             =   270
         Width           =   1740
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
         Left            =   150
         TabIndex        =   32
         Top             =   1395
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
         TabIndex        =   31
         Top             =   270
         Visible         =   0   'False
         Width           =   2220
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
         Left            =   540
         TabIndex        =   30
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
         TabIndex        =   27
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   7665
      TabIndex        =   23
      Top             =   6515
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total I.V.A."
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
      Left            =   9360
      TabIndex        =   42
      Top             =   5310
      Width           =   990
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
      Left            =   9360
      TabIndex        =   28
      Top             =   5895
      Width           =   510
   End
End
Attribute VB_Name = "A01_4500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Proveedor As Integer
'este vector se aparea con el Lv
Private VecCentroCta() As TipoAutorizacionDePago
Private VecFondosFijoPendiente() As TipoFondoFijoPendiente
Private Modificado As Boolean
Public NroAutorizacion As Long
Private FechaMin As Date
Private VecCuentasContables() As CuentasContables
Private SinPres As Boolean
Private TotalMontoSinPres As Double
Private MontoSinPres As Double

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalFecha_Change()
    CalFecha.Value = ValidarPeriodo(CalFecha.Value)
    ReDim VecCentroCta(0)
    LvCenCostoCtas.ListItems.Clear
    LvCenCostoCtas.ListItems.Add
    TxtTotal = ""
End Sub

Private Sub CalFecha_GotFocus()
    CalFecha.CustomFormat = "MM/yyyy"
End Sub

Private Sub CmbFondoFijoNro_Click()
    TxtTotalFF.Text = Replace(Format(VecFondosFijoPendiente(CmbFondoFijoNro.ListIndex).R_TotalARendir, "0.00"), ",", ".")
    If CmbFondoFijoNro.ListIndex > 0 Then
        CalFecha.CustomFormat = "MM/yyyy"
        CalFecha.Value = VecFondosFijoPendiente(CmbFondoFijoNro.ListIndex).R_Fecha
    Else
        CalFecha.CustomFormat = " "
    End If
End Sub

Private Sub CmbProv_Click()
    Proveedor = VecProveedores(CmbProv.ListIndex).Codigo
End Sub

Private Sub CmdAgregar_Click()
'On Error GoTo errores
Dim i As Integer
  SinPres = False
   
 If ValidarCarga Then
    Modificado = True
        
    If LvCenCostoCtas.SelectedItem.Index = LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Index Then
        
       'agrega al vector
        i = LvCenCostoCtas.SelectedItem.Index
        ReDim Preserve VecCentroCta(UBound(VecCentroCta) + 1)
        
        VecCentroCta(UBound(VecCentroCta)).O_CentroDeCostoEmisor = CentroEmisor
        VecCentroCta(UBound(VecCentroCta)).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
        VecCentroCta(UBound(VecCentroCta)).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
        VecCentroCta(UBound(VecCentroCta)).PrecioReal = Val(TxtPrecioU)
        VecCentroCta(UBound(VecCentroCta)).MontoSinPresupuestar = MontoSinPres
        'lo pone en el LV
        With VecCentroCta(UBound(VecCentroCta))
           LvCenCostoCtas.ListItems(i).Text = BuscarDescCentroEmisor(CentroEmisor)
           LvCenCostoCtas.ListItems(i).SubItems(1) = "Centro: " & BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto)) & " - " & BuscarDescCentro(.O_CentroDeCosto)
           LvCenCostoCtas.ListItems(i).SubItems(2) = BuscarDescCta(.O_CuentaContable)
           LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.PrecioReal, "0.00")
        End With

        'es el último registro, por lo tanto quería agregar uno nuevo
         LvCenCostoCtas.ListItems.Add
        'pocisiona en el último
         LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Selected = True
        'limpia los controles
         CmbCentrosDeCostos.ListIndex = 0
         CmbCuentas.ListIndex = 0
         TxtPrecioU.Text = ""
     End If
     
     'le da el foco el combo de centros de costo
     CmbCentrosDeCostos.SetFocus
  End If
  TotalMontoSinPres = CalcularTotalSinPresupuestar
  Call CalcularTotal
Errores:
    If Err.Number <> 0 Then
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Function ValidarCarga() As Boolean
    ValidarCarga = True
    Dim i As Integer
    Dim Rta As Integer
On Error GoTo Errores

    For i = 1 To UBound(VecCentroCta)
       If i <> LvCenCostoCtas.SelectedItem.Index Then
            If VecCentroCta(i).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo _
               And VecCentroCta(i).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
               
                 MsgBox "Ya existe esta Combinación de Centro de Costo - Cuenta"
                 ValidarCarga = False
                 Exit Function
            End If
      End If

    Next
    
    If CmbCentrosDeCostos.ListIndex <= 0 Then
        MsgBox "Debe seleccionat un Centro de Costo", vbInformation
        CmbCentrosDeCostos.SetFocus
        ValidarCarga = False
        Exit Function
    End If

    If CmbCuentas.ListIndex <= 0 Then
        MsgBox "Debe seleccionat una Cuneta Contable", vbInformation
        CmbCuentas.SetFocus
        ValidarCarga = False
        Exit Function
    End If
    
    If Val(TxtPrecioU) = 0 Then
        MsgBox "Debe ingresar un precio", vbInformation
        TxtPrecioU.SetFocus
        ValidarCarga = False
        Exit Function
    End If

    Dim Sql As String
    Dim MontoDisponible As Double
    Dim RsValidar As New ADODB.Recordset
    i = LvCenCostoCtas.SelectedItem.Index
    MontoSinPres = 0
    If VecCuentasContables(CmbCuentas.ListIndex).Codigo = "5121" Then
         Sql = "SpOCPresupuestosDistribucionValidarAutorizacionDeCargaContable" & _
                     "   @CuentaContable ='" & VecCuentasContables(CmbCuentas.ListIndex).Codigo & _
                     "', @CentroEmisor ='" & CentroEmisor & _
                     "', @Periodo ='" & Format(CalFecha, "MM/yyyy") & _
                     "', @NroAutorizacion=" & NroAutorizacion & _
                     " , @SubCentroDeCosto='" & VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo & "'"
                     
         RsValidar.Open Sql, Conec
         'acá está el problema
         MontoDisponible = ValN(RsValidar!MontoDisponible)  'TotalMontoSinPres
        
         If MontoDisponible < Val(TxtPrecioU.Text) Then
            MsgBox "El monto disponible para este período es de $" & Format(MontoDisponible, "0.00"), vbInformation
            ValidarCarga = False
            Exit Function
         End If
    Else
        Sql = "SpOCPresupuestosRenglonesValidarAutorizacionDeCargaContable" & _
                    "   @CuentaContable ='" & VecCuentasContables(CmbCuentas.ListIndex).Codigo & _
                    "', @CentroEmisor ='" & CentroEmisor & _
                    "', @Periodo ='" & Format(CalFecha, "MM/yyyy") & _
                    "', @NroAutorizacion=" & NroAutorizacion
                    
        RsValidar.Open Sql, Conec
        'acá está el problema
        MontoDisponible = ValN(RsValidar!MontoDisponible) - CalcularTotalPorCuenta(VecCuentasContables(CmbCuentas.ListIndex).Codigo)  'TotalMontoSinPres
       
        If MontoDisponible < Val(TxtPrecioU.Text) Then
            Rta = MsgBox("¿Desea imputar la cantidad Sin Presupuestar?", vbYesNo, "Fuera de Presupuesto")
            If Rta = vbNo Then
                ValidarCarga = False
                Exit Function
            End If
            'si no alcanza con lo presupuestado para una cuanta
            'se ve si se tiene disponible para hacerlo sin presupuestar
        
            Sql = " SpOCImporteSinPresupuestarAutorizacionDeCargaContable @CentroDeCosto ='" & CentroEmisor & _
                             "', @NroAutorizacion =" & NroAutorizacion & _
                             " , @Periodo=" & FechaSQL(CalFecha, "SQL")
            
            RsValidar.Close
            RsValidar.Open Sql, Conec
            MontoSinPres = Val(TxtPrecioU.Text) - IIf(MontoDisponible > 0, MontoDisponible, 0)
            MontoDisponible = RsValidar!MontoSinPresupuestarMensual - RsValidar!MontoSinPres - TotalMontoSinPres
            If MontoDisponible < MontoSinPres Then
                MsgBox "El monto disponible para este período es de $" & Format(MontoDisponible, "0.00"), vbInformation
                ValidarCarga = False
                Exit Function
            End If
        End If
    End If
Errores:
    If Err.Number <> 0 Then
        Call ManipularError(Err.Number, Err.Description)
        ValidarCarga = False
    End If
End Function

Private Sub CmdAnular_Click()
 Dim Sql As String
 Dim Rta As Integer
 On Error GoTo Error
    Rta = MsgBox("¿Está seguro de que desea Anular la Autorización de pago?", vbYesNo)
    If Rta = vbYes Then
        Sql = "SpOCAutorizacionesDePagoCabeceraAnular @A_NumeroDeAutorizacionDePago =" + CStr(NroAutorizacion)
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
    Unload BuscarAutorizacionDePagoFF
    BuscarAutorizacionDePagoFF.Show vbModal
    
    Timer1.Enabled = True
End Sub

Private Sub CmdBuscarSubCentro_Click()
    BuscarSubCentro.Show vbModal
    Call BuscarCentro(BuscarSubCentro.CodigoSubCentro, CmbCentrosDeCostos)
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar la Recepción?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarAutorizacion
    End If
End Sub

Private Sub CmdCargar_Click()
    Call CargarAutorizacionPago(Val(TxtNroAutorizacion))
    Modificado = False

End Sub

Private Sub CargarAutorizacionPago(NroAutorizacion As Long)
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim RsCargar As New ADODB.Recordset
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
    
    TotalMontoSinPres = 0
    
    If VerificarNulo(!A_FondoFijoNro, "N") = 0 Then
        MsgBox "La Autorización de Pago No es de F. F."
        Call CmdNuevo_Click
        Exit Sub
    End If
    
    Call CargarCmbFondoFijo(CmbFondoFijoNro)
    OptRecepcionada(0).Value = True
    Call UbicarCmbFondoFijo(!A_FondoFijoNro, CmbFondoFijoNro)
    TxtTotalFF.Enabled = False

    OptRecepcionada(0).Enabled = False
    OptRecepcionada(1).Enabled = False
    CmbFondoFijoNro.Enabled = False
    TxtIVA.Text = Replace(Format(!A_IvaFF, "0.00"), ",", ".")
    If Not IsNull(!A_FechaAnulacion) Then
        If Not IsNull(!A_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!A_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
        
        CmdCambiar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
        FrameAsig.Enabled = False
        CalFechaEmitida.Enabled = False
        TxtIVA.Enabled = False
    Else
        LBAnulada.Visible = False
        CmdCambiar.Enabled = True
        CmbProv.Enabled = True
        CalFecha.Enabled = True
        CmdAnular.Visible = True
        TxtObs.Enabled = True
        FrameAsig.Enabled = True
        CalFechaEmitida.Enabled = True
        TxtIVA.Enabled = True
    End If
  
    If VerificarNulo(!A_Seguro, "B") = True Then
        CmdCambiar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
        FrameAsig.Enabled = False
        MsgBox "La Certificación de servicio no puede ser Modificada por se de Seguros", vbInformation
    End If
    
     
     CalFecha.Value = !A_Fecha
     CalFecha.CustomFormat = "MM/yyyy"
     If IsNull(!A_FechaEmision) Then
        CalFechaEmitida.Value = !A_Fecha
     Else
        CalFechaEmitida.Value = !A_FechaEmision
     End If

     Proveedor = !A_CodigoProveedor
     'CmbProv.Text = Trim(!P_Descripcion)
     Call UbicarProveedor(!A_CodigoProveedor, CmbProv)
     CmbProv.Enabled = False
     TxtObs.Text = VerificarNulo(!A_Observaciones)
     TxtNroAutorizacion.Text = Format(NroAutorizacion, "0000000000")
     
     Me.NroAutorizacion = NroAutorizacion
     .Close
     
     Sql = "SpOCAutorizacionesDePagoRenglonesTraer @A_NumeroDeAutorizacionDePago =" & NroAutorizacion
     .Open Sql, Conec
     ReDim VecCentroCta(.RecordCount)
     
     If !O_NumeroDeOrdenDeContratacionDeServicios <> 0 Then
        MsgBox "La Autorización de Carga Contable no es del tipo especial", vbInformation
        Call CmdNuevo_Click
        Exit Sub
     End If
     
    i = 1
    While Not .EOF
        VecCentroCta(i).O_NumeroOrdenDeContratacion = !O_NumeroDeOrdenDeContratacionDeServicios
        VecCentroCta(i).O_CentroDeCostoEmisor = !O_CentroDeCostoEmisor
        VecCentroCta(i).O_CuentaContable = BuscarDescCentro(!A_CentroDeCosto)
        VecCentroCta(i).O_CentroDeCosto = !A_CentroDeCosto
        VecCentroCta(i).O_CuentaContable = !A_CuentaContable
        VecCentroCta(i).O_PrecioPactado = !O_ImporteOrdenDeContratacion
        VecCentroCta(i).PrecioReal = !A_Importe
        VecCentroCta(i).MontoSinPresupuestar = VerificarNulo(!A_ImporteSinPresupuestar)
        
        LvCenCostoCtas.ListItems.Add
        LvCenCostoCtas.ListItems(i).Text = BuscarDescCentroEmisor(!O_CentroDeCostoEmisor)
        LvCenCostoCtas.ListItems(i).SubItems(1) = "Centro: " & BuscarDescCentroEmisor(BuscarCentroPadre(VecCentroCta(i).O_CentroDeCosto)) & " - " & BuscarDescCentro(!A_CentroDeCosto)
        LvCenCostoCtas.ListItems(i).SubItems(2) = BuscarDescCta(!A_CuentaContable)
        LvCenCostoCtas.ListItems(i).SubItems(3) = Format(!A_Importe, "0.00##")
        i = i + 1
        .MoveNext
    Wend
        .Close
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
End With
    Call CalcularTotal
    TxtTotalFF = Format(ValN(TxtTotal) + ValN(TxtIVA), "0.00")
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
       Sql = "SpOCAutorizacionesDePagoCabeceraAgregar @A_Fecha= " & FechaSQL(CStr(CalFecha.Value), "SQL") & _
                                                  " , @A_CodigoProveedor = " & VecProveedores(CmbProv.ListIndex).Codigo & _
                                                  " , @U_Usuario ='" & Usuario & _
                                                  "', @A_Observaciones = '" & TxtObs.Text & _
                                                  "', @A_FechaEmision = " & FechaSQL(CalFechaEmitida, "SQL") & _
                                                  " , @@A_FondoFijoNro= " & VecFondosFijoPendiente(CmbFondoFijoNro.ListIndex).R_NumeroFondoFijo & _
                                                  " , @@A_IvaFF=" & Replace(Val(TxtIVA), ",", ".")
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
         RepAutorizacionDePagoFF.Show vbModal
       End If
         
       If FrmMensaje.Retorno = vbNuevo Then
         Call LimpiarAutorizacion
         TxtNroAutorizacion.Text = ""
       End If
       
       If FrmMensaje.Retorno = vbExportesPDF Then
          Call CmdExpPdf_Click
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
                                                    ", @A_CodigoProveedor = " & VecProveedores(CmbProv.ListIndex).Codigo & _
                                                    ", @U_Usuario ='" & Usuario & _
                                                   "', @A_Observaciones = '" & TxtObs.Text & _
                                                   "', @A_FechaEmision = " & FechaSQL(CalFechaEmitida, "SQL") & _
                                                   " , @@A_FondoFijoNro= " & VecFondosFijoPendiente(CmbFondoFijoNro.ListIndex).R_NumeroFondoFijo & _
                                                   " , @@A_IvaFF=" & Replace(Val(TxtIVA), ",", ".")
      Conec.Execute Sql
      
      Sql = "SpOCAutorizacionesDePagoRenglonesBorrar @A_NumeroDeAutorizacionDePago=" & NroAutorizacion
      Conec.Execute Sql
    
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
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
         TxtNroAutorizacion.Text = ""
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
    FrameAsig.Enabled = True
    ReDim VecRecepcion(0)
    ReDim VecCentroCta(0)
    Proveedor = 0
    NroAutorizacion = 0
    LvCenCostoCtas.ListItems.Clear
    CmbProv.Enabled = True
    CmbProv.ListIndex = 0
    CmbFondoFijoNro.Enabled = True
    CmbFondoFijoNro.ListIndex = 0
    TxtIVA.Text = ""
    CalFechaEmitida.Enabled = True
    TxtIVA.Enabled = True

    CalFecha.Value = ValidarPeriodo(Date, False)
    CalFecha.Format = dtpCustom
    CalFecha.CustomFormat = " "
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)
    CalFecha.Enabled = True
    CmdConfirnar.Visible = True
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    CmdCambiar.Visible = False
    CmdAnular.Visible = False
    Modificado = False
    CmbCuentas.ListIndex = 0
    CmbCentrosDeCostos.ListIndex = 0
    TxtObs.Text = ""
    TxtObs.Enabled = True
    TxtPrecioU.Text = ""
    TxtTotal.Text = ""
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
End Sub

Private Function ValidarEncabezado() As Boolean
    
    ValidarEncabezado = True
    
    If CalFecha.CustomFormat = " " Then
        MsgBox "Debe Período Fecha de Imputación"
        CalFecha.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
    
    If CmbProv.ListIndex <= 0 Then
        MsgBox "Debe Seleccionar un Proveedor", vbExclamation
        CmbProv.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
    
    If CmbFondoFijoNro.ListIndex <= 0 Then
        MsgBox "Debe Seleccionar un Fonfo Fijo", vbExclamation
        CmbFondoFijoNro.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
    
    If Val(Replace(TxtTotal.Text, ",", ".")) = 0 Then
       MsgBox "La autorización de pago debe tener un importe mayor a 0", vbExclamation, "Total"
       LvCenCostoCtas.SetFocus
       ValidarEncabezado = False
       Exit Function
    End If
    
    If Abs(ValN(TxtIVA) + ValN(TxtTotal) - ValN(TxtTotalFF)) > 0.01 Then
       MsgBox "La autorización de pago debe tener El mismo importe que el F.F. sin IVA", vbExclamation, "Total"
       LvCenCostoCtas.SetFocus
       ValidarEncabezado = False
       Exit Function
    End If
End Function

Private Sub CmdEliminar_Click()
    Dim IndexBorrar As Integer
    'guarda el índice del vector
    IndexBorrar = LvCenCostoCtas.SelectedItem.Index
    
    'borra del LV
     LvCenCostoCtas.ListItems.Remove (IndexBorrar)
       
    'borrar del vector
    While IndexBorrar < UBound(VecCentroCta)
        VecCentroCta(IndexBorrar) = VecCentroCta(IndexBorrar + 1)
        IndexBorrar = IndexBorrar + 1
    Wend
        
    ReDim Preserve VecCentroCta(UBound(VecCentroCta) - 1)
      'reposiciona el lv
    LvCenCostoCtas.ListItems(IndexBorrar).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
    TotalMontoSinPres = CalcularTotalSinPresupuestar
    Call CalcularTotal
End Sub

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeAutorizacion
         RepAutorizacionDePagoFF.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepAutorizacionDePagoFF.Pages
         Unload RepAutorizacionDePagoFF
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
    RepAutorizacionDePagoFF.Show
End Sub

Private Sub ConfImpresionDeAutorizacion()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
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
     
    RepAutorizacionDePagoFF.TxtFecha = Format(CalFecha.Value, "MM/yyyy")
    RepAutorizacionDePagoFF.TxtFEmision = CStr(CalFechaEmitida.Value)
    RepAutorizacionDePagoFF.TxtNroOrden.Text = TxtNroAutorizacion.Text
    RepAutorizacionDePagoFF.BarNroAutorizacion.Caption = TxtNroAutorizacion.Text
    RepAutorizacionDePagoFF.TxtProv.Text = CmbProv.Text & " (Cód. " & Proveedor & ")"
    RepAutorizacionDePagoFF.TxtAnulada.Visible = LBAnulada.Visible
    RepAutorizacionDePagoFF.TxtAnulada.Text = LBAnulada.Caption
    RepAutorizacionDePagoFF.TxtObservaciones.Text = TxtObs.Text
    RepAutorizacionDePagoFF.TxtFondoFijoNro = CmbFondoFijoNro.Text
    RepAutorizacionDePagoFF.TxtIVA = TxtIVA
    RepAutorizacionDePagoFF.TxtTotalConIva = Format(ValN(TxtIVA) + ValN(TxtTotal), "0.00")
    RepAutorizacionDePagoFF.TxtTotalFF = TxtTotalFF
    RepAutorizacionDePagoFF.LbObservacionesFF = IIf(OptRecepcionada(1).Value, "NOTA: El Fondo Fijo aún no ha sido recibido", "")
    RepAutorizacionDePagoFF.Zoom = -1
    RepAutorizacionDePagoFF.DataControl1.Recordset = RsListado
End Sub

Private Sub CmdModif_Click()
On Error GoTo Errores
Dim i As Integer
   SinPres = False
    i = LvCenCostoCtas.SelectedItem.Index
    TotalMontoSinPres = TotalMontoSinPres - VecCentroCta(i).MontoSinPresupuestar
   
   If ValidarCarga Then
        Modificado = True
        With VecCentroCta(i)
            .O_CentroDeCostoEmisor = CentroEmisor
            .O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
            .O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
            .PrecioReal = Val(TxtPrecioU)
            .MontoSinPresupuestar = MontoSinPres
    
           'lo pone en el LV
            LvCenCostoCtas.ListItems(i).Text = BuscarDescCentroEmisor(CentroEmisor)
            LvCenCostoCtas.ListItems(i).SubItems(1) = "Centro: " & BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto)) & " - " & BuscarDescCentro(.O_CentroDeCosto)
            LvCenCostoCtas.ListItems(i).SubItems(2) = BuscarDescCta(.O_CuentaContable)
            LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.PrecioReal, "0.00")
        End With

       'limpia los controles
        'CmbCentrosDeCostos.ListIndex = 0
        'CmbCuentas.ListIndex = 0
        'TxtPrecioU.Text = ""
        'le da el foco el combo de centros de costo
        CmbCentrosDeCostos.SetFocus
   End If
   TotalMontoSinPres = CalcularTotalSinPresupuestar
   Call CalcularTotal
Errores:
    If Err.Number <> 0 Then
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub

Private Sub CmdNuevo_Click()
    Call LimpiarAutorizacion
    'FramePrecio.Enabled = True
    LBAnulada.Visible = False
    'FechaMin = Date
    TxtNroAutorizacion.Text = ""
    'Call CargarCmbFondoFijo(CmbFondoFijoNro)
    OptRecepcionada(0).Value = True
    TotalMontoSinPres = 0
    OptRecepcionada(0).Enabled = True
    OptRecepcionada(1).Enabled = True
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    CalFecha.Value = Date 'ValidarPeriodo(Date, False)
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)

    Call CargarVecCuentas(CentroEmisor)
    Call CargarCmbFondoFijo(CmbFondoFijoNro)
    Call CargarCmbCentrosDeCostos(CmbCentrosDeCostos)
    'Call CargarComboCuentasContables(CmbCuentas)
    ReDim Preserve VecCentroCta(0)
    LvCenCostoCtas.ListItems.Add
    Modificado = False
End Sub

Private Sub OptRecepcionada_Click(Index As Integer)
    TxtTotalFF.Enabled = Index = 1
    If Index = 0 Then
        Call CargarCmbFondoFijo(CmbFondoFijoNro)
    Else
        Call CargarCmbFondoFijoSinRecepcionar(CmbFondoFijoNro)
    End If
End Sub

Private Sub TxtIVA_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtIVA, KeyAscii)
End Sub

Private Sub TxtNroAutorizacion_LostFocus()
  If Val(TxtNroAutorizacion.Text) <> NroAutorizacion Then
    CmdConfirnar.Visible = TxtNroAutorizacion.Text = ""
    CmdCambiar.Visible = TxtNroAutorizacion.Text <> ""
    Call LimpiarAutorizacion
  End If

End Sub

Private Sub CargarCmbFondoFijo(Cmb As ComboEsp)
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    Dim i As Integer
On Error GoTo ErrorCarga
    Cmb.Clear
    Sql = "SpOcAutorizacionesDePagoFfPendientes @CentroDeCosto ='" & CentroEmisor & _
                                            "', @NumeroDeAutorizacionDePago=" & Val(TxtNroAutorizacion)
    With RsCargar
        .Open Sql, Conec
        ReDim VecFondosFijoPendiente(.RecordCount)
        Cmb.AddItem "Seleccione un F. F."
        For i = 1 To .RecordCount
            VecFondosFijoPendiente(i).R_Numero = !R_Numero
            VecFondosFijoPendiente(i).R_NumeroFondoFijo = !R_NumeroFondoFijo
            VecFondosFijoPendiente(i).R_TotalARendir = !R_TotalARendir
            VecFondosFijoPendiente(i).R_Fecha = !R_Fecha

            Cmb.AddItem !R_NumeroFondoFijo
            .MoveNext
        Next
        Cmb.ListIndex = 0
    End With
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarCmbFondoFijoSinRecepcionar(Cmb As ComboEsp)
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    Dim i As Integer
On Error GoTo ErrorCarga
    Cmb.Clear
    Sql = "SpOcAutorizacionesDePagoFfSinRecepcionar @CentroDeCosto ='" & CentroEmisor & "', @NumeroDeAutorizacionDePago =" & Val(TxtNroAutorizacion)
    With RsCargar
        .Open Sql, Conec
        ReDim VecFondosFijoPendiente(.RecordCount)
        Cmb.AddItem "Seleccione un F. F."
        For i = 1 To .RecordCount
            VecFondosFijoPendiente(i).R_Numero = !R_Numero
            VecFondosFijoPendiente(i).R_NumeroFondoFijo = !R_NumeroFondoFijo
            VecFondosFijoPendiente(i).R_TotalARendir = !R_TotalARendir
            VecFondosFijoPendiente(i).R_Fecha = !R_Fecha
            Cmb.AddItem !R_NumeroFondoFijo
            .MoveNext
        Next
        Cmb.ListIndex = 0
    End With
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub UbicarCmbFondoFijo(NumeroFF As Long, Cmb As ComboEsp)
    Dim i As Integer
    For i = 1 To UBound(VecFondosFijoPendiente)
        If VecFondosFijoPendiente(i).R_NumeroFondoFijo = NumeroFF Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
    'si no se encuentra es porque no está recepcionado
    OptRecepcionada(1).Value = True
    For i = 1 To UBound(VecFondosFijoPendiente)
        If VecFondosFijoPendiente(i).R_NumeroFondoFijo = NumeroFF Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
    Cmb.ListIndex = 0
End Sub

Private Sub TxtPrecioU_KeyPress(KeyAscii As Integer)
    Call TxtNumericoNeg(TxtPrecioU, KeyAscii)
End Sub

Private Sub CrearEncabezados()
    'LvCenCostoCtas.ColumnHeaders.Add , , "Nº de Orden de Contratación", 1350
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos Emisor", 2650
    LvCenCostoCtas.ColumnHeaders.Add , , "Sub-Centros de Costos", 3450
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", 2550
    LvCenCostoCtas.ColumnHeaders.Add , , "Precio sin IVA", 1300, 1
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
        TxtTotal.Text = Format(Total, "0.00##")

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
    If Item.Index < LvCenCostoCtas.ListItems.Count Then
        Call BuscarCuentaContable(VecCentroCta(Item.Index).O_CuentaContable, CmbCuentas)
        Call BuscarCentro(VecCentroCta(Item.Index).O_CentroDeCosto, CmbCentrosDeCostos)
        TxtPrecioU.Text = Replace(VecCentroCta(Item.Index).PrecioReal, ",", ".")
        
        CmdModif.Enabled = True
        CmdAgregar.Enabled = False
        CmdEliminar.Enabled = True
    Else
        CmbCentrosDeCostos.ListIndex = 0
        CmbCuentas.ListIndex = 0
        TxtPrecioU.Text = ""
        CmdModif.Enabled = False
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
    End If
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

Private Sub CargarVecCuentas(CentroEmisor As String)

 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
      ReDim VecCuentasContables(0)
      'en esta sección carga las cuentas
      'que están asociadas a algún centro de costo
  With RsCargar
      Sql = "SpOcRelacionCentroDeCostoCuentaContableTodos"
      '"SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CentroEmisor & "'"
      .Open Sql, Conec, adOpenStatic, adLockReadOnly
        For i = 1 To UBound(Ayudas.VecCuentasContables)
            .Find "R_CuentaContable = '" & Ayudas.VecCuentasContables(i).Codigo & "'", , , 1
           If Not .EOF Then
              ReDim Preserve VecCuentasContables(UBound(VecCuentasContables) + 1)
              VecCuentasContables(UBound(VecCuentasContables)) = Ayudas.VecCuentasContables(i)
           
           End If
        Next
                        
  End With
    'carga los combos con los valores de los vectores locales
    Call CargarCmbCuentasContables(CmbCuentas)
End Sub

Public Sub CargarCmbCuentasContables(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    
    If Tipo = "Elegir" Then
       Cmb.AddItem "Seleccione una Cuenta Contable"
    Else
       Cmb.AddItem "Todas las Cuentas Contables"
    End If

    For i = 1 To UBound(VecCuentasContables)
        Cmb.AddItem Trim(VecCuentasContables(i).Descripcion)
    Next
        
    Cmb.ListIndex = 0
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub BuscarCuentaContable(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer

    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = C_Codigo Then
            Cmb.ListIndex = i
        End If
    Next
    
End Sub

Private Function CalcularTotalSinPresupuestar() As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        Total = Total + VecCentroCta(i).MontoSinPresupuestar
    Next
       
    CalcularTotalSinPresupuestar = Total
End Function

Private Function CalcularTotalPorCuenta(Cuenta As String) As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        If VecCentroCta(i).O_CuentaContable = Cuenta And _
           Not LvCenCostoCtas.ListItems(i).Selected Then
             Total = Total + VecCentroCta(i).PrecioReal
        End If
    Next
       
    CalcularTotalPorCuenta = Total
End Function

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call BuscarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub CmbCuentas_Click()
    TxtCodCuenta.Text = VecCuentasContables(CmbCuentas.ListIndex).Codigo
End Sub

Private Sub TxtTotalFF_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtTotalFF, KeyAscii)
End Sub
