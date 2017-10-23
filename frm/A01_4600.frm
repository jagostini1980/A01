VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_4600 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificación Turismo"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
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
      TabIndex        =   42
      Top             =   5685
      Width           =   960
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Height          =   1410
      Left            =   60
      TabIndex        =   31
      Top             =   5190
      Width           =   9240
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   5670
         MaxLength       =   4
         TabIndex        =   11
         Top             =   795
         Width           =   645
      End
      Begin VB.CommandButton CmdBuscarSubCentro 
         Caption         =   "Buscar"
         Height          =   300
         Left            =   5580
         TabIndex        =   9
         Top             =   382
         Width           =   735
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7965
         TabIndex        =   13
         Top             =   180
         Width           =   1150
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7965
         TabIndex        =   15
         Top             =   990
         Width           =   1150
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   7965
         TabIndex        =   14
         Top             =   585
         Width           =   1150
      End
      Begin Controles.ComboEsp CmbCentrosDeCostos 
         Height          =   315
         Left            =   1980
         TabIndex        =   8
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
         TabIndex        =   10
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
         TabIndex        =   12
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   435
         Width           =   1830
      End
      Begin VB.Label LbPrecio 
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   855
         Width           =   1485
      End
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3030
      TabIndex        =   17
      Top             =   6675
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Autorización"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5775
      TabIndex        =   19
      Top             =   6675
      Width           =   1815
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   1725
      TabIndex        =   16
      Top             =   6660
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   4380
      TabIndex        =   18
      Top             =   6675
      Width           =   1230
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   7650
      TabIndex        =   20
      Top             =   6660
      Width           =   1320
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   9090
      TabIndex        =   22
      Top             =   6645
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   6630
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
      TabIndex        =   27
      Text            =   "0"
      Top             =   6255
      Width           =   960
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   3060
      Left            =   60
      TabIndex        =   7
      Top             =   2175
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   5398
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
      Height          =   2070
      Left            =   60
      TabIndex        =   23
      Top             =   60
      Width           =   10320
      Begin VB.CommandButton CmdAplicarAnticipo 
         Caption         =   "&Aplicar Anticipos"
         Height          =   315
         Left            =   8895
         TabIndex        =   53
         Top             =   990
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtFileNro 
         Height          =   315
         Left            =   1500
         TabIndex        =   44
         Top             =   997
         Width           =   1140
      End
      Begin VB.TextBox TxtFactura 
         Height          =   315
         Left            =   6840
         TabIndex        =   41
         Top             =   997
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.OptionButton OptComprobante 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Factura"
         Height          =   240
         Index           =   1
         Left            =   5925
         TabIndex        =   40
         Top             =   1034
         Width           =   915
      End
      Begin VB.OptionButton OptComprobante 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anticipo"
         Height          =   240
         Index           =   0
         Left            =   5010
         TabIndex        =   39
         Top             =   1020
         Width           =   960
      End
      Begin VB.TextBox TxtTotalAnt 
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
         Left            =   4095
         TabIndex        =   37
         Text            =   "0"
         Top             =   997
         Width           =   885
      End
      Begin VB.TextBox TxtObs 
         Height          =   315
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1695
         Width           =   8715
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
         Height          =   315
         Left            =   8955
         TabIndex        =   3
         Top             =   225
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   108199939
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1500
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
         Format          =   108199937
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbPaquetes 
         Height          =   330
         Left            =   1500
         TabIndex        =   45
         Top             =   1350
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   582
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
      Begin MSComCtl2.DTPicker CalDesde 
         Height          =   315
         Left            =   6990
         TabIndex        =   49
         Top             =   1350
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   108199937
         CurrentDate     =   38993
      End
      Begin MSComCtl2.DTPicker CalHasta 
         Height          =   315
         Left            =   8940
         TabIndex        =   51
         Top             =   1350
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   108199937
         CurrentDate     =   38993
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   255
         TabIndex        =   46
         Top             =   915
         Width           =   1365
         Begin VB.OptionButton OptTipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "File Nº:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   48
            Top             =   120
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton OptTipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Destinos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   47
            Top             =   450
            Width           =   1125
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta:"
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
         Left            =   8325
         TabIndex        =   52
         Top             =   1410
         Width           =   570
      End
      Begin VB.Label Label3 
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
         Left            =   5745
         TabIndex        =   50
         Top             =   1410
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Anticipos:"
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
         Left            =   2700
         TabIndex        =   38
         Top             =   1050
         Width           =   1350
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
         TabIndex        =   36
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
         TabIndex        =   24
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
         Left            =   120
         TabIndex        =   30
         Top             =   1755
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
         TabIndex        =   29
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
         Left            =   510
         TabIndex        =   28
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
         TabIndex        =   25
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   7665
      TabIndex        =   21
      Top             =   6660
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label LbTotalIva 
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
      TabIndex        =   43
      Top             =   5460
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
      TabIndex        =   26
      Top             =   6045
      Width           =   510
   End
End
Attribute VB_Name = "A01_4600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Proveedor As Integer
'este vector se aparea con el Lv
Private VecCentroCta() As TipoAutorizacionDePago
Private VecFiles() As TipoFiles
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
    txtTotal = ""
End Sub

Private Sub CalFecha_GotFocus()
    CalFecha.CustomFormat = "MM/yyyy"
End Sub

Private Sub CmbPaquetes_Validate(Cancel As Boolean)
Dim Sql As String
Dim RsCargar As New ADODB.Recordset
    If CmbPaquetes.ListIndex = 0 Then
        Exit Sub
    End If
    
    ' Sql = "SpOcAutorizacionesDePagoPaquete @Paquete=" & VecPaquetes(CmbPaquetes.ListIndex).P_Codigo & _
    '                                     ", @Proveedor=" & VecProveedores(CmbProv.ListIndex).Codigo
    'With RsCargar
    '    .Open Sql, Conec
    '    If Not .EOF Then
    '        TxtTotalAnt = Replace(ValN(!Anticipo), ",", ".")
    '    End If
    'End With

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
    
    If Val(TxtPrecioU) = 0 And TxtPrecioU.Enabled Then
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

Private Sub CmdAplicarAnticipo_Click()
    A01_4610.Proveedor = Val(VecProveedores(CmbProv.ListIndex).Codigo)
    A01_4610.Destino = VecPaquetes(Me.CmbPaquetes.ListIndex).P_Codigo
    A01_4610.FileNro = Val(TxtFileNro)
    A01_4610.Show vbModal
    If A01_4610.Aceptar Then
       TxtTotalAnt = A01_4610.TotalAnticipo
    End If
End Sub

Private Sub CMDBuscar_Click()
    Unload BuscarAutorizacionDePagoFiles
    BuscarAutorizacionDePagoFiles.Show vbModal
    
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
    Dim RsAnticipos As New ADODB.Recordset
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
    
    If VerificarNulo(!A_FileNro, "N") = 0 And ValN(!A_Paquete) = 0 Then
        MsgBox "La Autorización de Pago No es de File"
        Call CmdNuevo_Click
        Exit Sub
    End If
    
    If ValN(!A_FileNro) <> 0 Then
        OptTipo(1).Value = False
        OptTipo(0).Value = False
        OptTipo(1).Value = True
        TxtFileNro = ValN(!A_FileNro)
        CmbPaquetes.ListIndex = 0
        CalDesde.Value = Date
        CalHasta.Value = Date
    End If
    
    OptComprobante(!A_Comprobante).Value = True
    If ValN(!A_Paquete) <> 0 Then
       OptTipo(0).Value = True
       TxtFileNro = ""
       Call UbicarCmbPaquetes(ValN(!A_Paquete), CmbPaquetes)
       CalDesde.Value = IIf(IsNull(!A_Desde), Date, !A_Desde)
       CalHasta.Value = IIf(IsNull(!A_Hasta), Date, !A_Hasta)
    End If
    
    Sql = "SpOcAutorizacionesDePagoCabeceraTraerAnticipor @NumeroDeAutorizacionDePago=" & NroAutorizacion
    RsAnticipos.Open Sql, Conec
    If Not RsAnticipos.EOF Then
       TxtTotalAnt = Replace(ValN(RsAnticipos!A_Importe), ",", ".")
    End If
    'TxtTotalFactura.Enabled = False
    TxtFactura = !A_Factura
    'TxtTotalFactura = Replace(Format(ValN(!A_NetoFactura), "0.00"), ",", ".")
    OptComprobante(ValN(!A_Comprobante)).Enabled = True
    
    TxtFileNro.Enabled = False
    TxtIVA.Text = Replace(Format(!A_IvaFF, "0.00"), ",", ".")
    If Not IsNull(!A_FechaAnulacion) Then
        If Not IsNull(!A_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!A_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
        OptTipo(0).Enabled = False
        OptTipo(1).Enabled = False
        TxtFileNro.Enabled = False
        CmbPaquetes.Enabled = False
        CmdCambiar.Enabled = False
        CmbProv.Enabled = False
        CalFecha.Enabled = False
        CmdAnular.Visible = False
        TxtObs.Enabled = False
        FrameAsig.Enabled = False
        CalFechaEmitida.Enabled = False
        TxtIVA.Enabled = False
    Else
        OptTipo(0).Enabled = True
        OptTipo(1).Enabled = True
        'TxtFileNro.Enabled = True
        'CmbPaquetes.Enabled = True
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
    'TxtTotalFF = Format(ValN(TxtTotal) + ValN(TxtIVA), "0.00")
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

'On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
       Sql = "SpOCAutorizacionesDePagoCabeceraAgregar @A_Fecha= " & FechaSQL(CStr(CalFecha.Value), "SQL") & _
                                                  " , @A_CodigoProveedor = " & VecProveedores(CmbProv.ListIndex).Codigo & _
                                                  " , @U_Usuario ='" & Usuario & _
                                                  "', @A_Observaciones = '" & TxtObs.Text & _
                                                  "', @A_FechaEmision = " & FechaSQL(CalFechaEmitida, "SQL") & _
                                                  " , @@A_IvaFF=" & Replace(Val(TxtIVA), ",", ".") & _
                                                  " , @A_FileNro =" & Val(TxtFileNro) & _
                                                  " , @A_Factura ='" & TxtFactura.Text & _
                                                  "', @A_Comprobante=" & IIf(Me.OptComprobante(0).Value, 0, 1) & _
                                                  " , @A_Paquete =" & VecPaquetes(CmbPaquetes.ListIndex).P_Codigo & _
                                                  " , @A_Desde=" & IIf(OptTipo(0).Value, FechaSQL(CalDesde, "SQL"), "Null") & _
                                                  " , @A_Hasta=" & IIf(OptTipo(0).Value, FechaSQL(CalHasta, "SQL"), "Null")
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
    
    For i = 1 To UBound(VecAutorizacionesAnticiposApli)
        Sql = "SpOcAutorizacionesDePagoCabeceraAplicarAnticipo @AplicadoEnAutorizacion =" & NroAutorizacion & _
                                                            ", @NumeroDeAutorizacionDePago =" & VecAutorizacionesAnticiposApli(i)
        Conec.Execute Sql
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
         RepAutorizacionDePagoFile.Show vbModal
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
                                                   " , @@A_IvaFF=" & Replace(Val(TxtIVA), ",", ".") & _
                                                   " , @A_FileNro =" & Val(TxtFileNro) & _
                                                   " , @A_Factura ='" & TxtFactura.Text & _
                                                   "', @A_Comprobante=" & IIf(Me.OptComprobante(0).Value, 0, 1) & _
                                                   " , @A_Paquete =" & VecPaquetes(CmbPaquetes.ListIndex).P_Codigo & _
                                                  " , @A_Desde=" & IIf(OptTipo(0).Value, FechaSQL(CalDesde, "SQL"), "Null") & _
                                                  " , @A_Hasta=" & IIf(OptTipo(0).Value, FechaSQL(CalHasta, "SQL"), "Null")
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
    
    For i = 1 To UBound(VecAutorizacionesAnticiposApli)
        Sql = "SpOcAutorizacionesDePagoCabeceraAplicarAnticipo @AplicadoEnAutorizacion =" & NroAutorizacion & _
                                                            ", @NumeroDeAutorizacionDePago =" & VecAutorizacionesAnticiposApli(i)
        Conec.Execute Sql
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
    TxtFileNro.Enabled = True
    TxtFileNro = ""
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
    txtTotal.Text = ""
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
    If OptTipo(1).Value Then
        If Val(TxtFileNro) = 0 Then
            MsgBox "Debe Ingresar un File", vbExclamation
            TxtFileNro.SetFocus
            ValidarEncabezado = False
            Exit Function
        End If
    Else
        If CmbPaquetes.ListIndex <= 0 Then
            MsgBox "Debe Seleccionar un Destino", vbExclamation
            CmbPaquetes.SetFocus
            ValidarEncabezado = False
            Exit Function
        End If
    End If
    
    If OptComprobante(0).Value Then
        If Val(Replace(txtTotal.Text, ",", ".")) = 0 Then
           MsgBox "La autorización de pago debe tener un importe mayor a 0", vbExclamation, "Total"
           LvCenCostoCtas.SetFocus
           ValidarEncabezado = False
           Exit Function
        End If
    End If
    'If Val(TxtIVA) + ValN(TxtTotal) <> ValN(TxtTotalFF) Then
    '   MsgBox "La autorización de pago debe tener El mismo importe que el F.F. sin IVA", vbExclamation, "Total"
    '   LvCenCostoCtas.SetFocus
    '   ValidarEncabezado = False
    '   Exit Function
    'End If
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
         RepAutorizacionDePagoFile.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepAutorizacionDePagoFile.Pages
         Unload RepAutorizacionDePagoFile
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
    RepAutorizacionDePagoFile.Show
End Sub

Private Sub ConfImpresionDeAutorizacion()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Dim RsAnticipos As New ADODB.Recordset
  Dim Sql As String
  Dim ObservacionesFile As String
  
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
    If OptTipo(1).Value Then
        RepAutorizacionDePagoFile.LbFilePequete = "File Nº:"
        RepAutorizacionDePagoFile.TxtFilePaquete = TxtFileNro
        RepAutorizacionDePagoFile.LbImporte = LbPrecio.Caption
    Else
        RepAutorizacionDePagoFile.LbFilePequete = "Paquete:"
        RepAutorizacionDePagoFile.TxtFilePaquete = CmbPaquetes.Text
    End If
    
    If OptComprobante(1).Value Then
        'Si es Factura
        Sql = "SpOcAutorizacionesDePagoCabeceraTraerAnticipor @NumeroDeAutorizacionDePago=" & NroAutorizacion
        With RsAnticipos
            .Open Sql, Conec
            ObservacionesFile = "Anticipos Realizados" & Chr(13)
            While Not .EOF
                ObservacionesFile = ObservacionesFile & "Autorizacion Nº: " & !A_NumeroDeAutorizacionDePago & _
                                    " - Fecha: " & Format(!A_Fecha, "dd/MM/yyyy") & " - Importe: " & Format(!A_Importe, "0.00") & _
                                    " - Usuario: " & !U_Usuario & Chr(13)
                .MoveNext
            Wend
        End With
        RepAutorizacionDePagoFile.TxtObservacionesAnticipos = ObservacionesFile
    End If

    RepAutorizacionDePagoFile.TxtFecha = Format(CalFecha.Value, "MM/yyyy")
    RepAutorizacionDePagoFile.TxtFEmision = CStr(CalFechaEmitida.Value)
    RepAutorizacionDePagoFile.TxtNroOrden.Text = TxtNroAutorizacion.Text
    RepAutorizacionDePagoFile.BarNroAutorizacion.Caption = TxtNroAutorizacion.Text
    RepAutorizacionDePagoFile.TxtProv.Text = CmbProv.Text & " (Cód. " & VecProveedores(CmbProv.ListIndex).Codigo & ")"
    RepAutorizacionDePagoFile.TxtAnulada.Visible = LBAnulada.Visible
    RepAutorizacionDePagoFile.TxtAnulada.Text = LBAnulada.Caption
    RepAutorizacionDePagoFile.TxtObservaciones.Text = TxtObs.Text

    RepAutorizacionDePagoFile.TxtIVA = TxtIVA
    RepAutorizacionDePagoFile.TxtTotalConIva = Format(ValN(TxtIVA) + ValN(txtTotal), "0.00")
    RepAutorizacionDePagoFile.TxtTotalAnticipos = Format(ValN(TxtTotalAnt), "0.00")
    'RepAutorizacionDePagoFile.LbNetoFactura = IIf(OptComprobante(0).Value, "", "Neto Factura: " & TxtTotalFactura)
    RepAutorizacionDePagoFile.LbFactura = TxtFactura
    RepAutorizacionDePagoFile.LbTipo = IIf(OptComprobante(0).Value, OptComprobante(0).Caption, OptComprobante(1).Caption)
    RepAutorizacionDePagoFile.Zoom = -1
    RepAutorizacionDePagoFile.DataControl1.Recordset = RsListado
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
    TxtTotalAnt = ""
    'TxtTotalFactura = ""
    TxtFactura = ""
    'FechaMin = Date
    TxtNroAutorizacion.Text = ""
    'Call CargarCmbFondoFijo(CmbFondoFijoNro)
    OptComprobante(0).Value = True
    TotalMontoSinPres = 0
    OptComprobante(0).Enabled = True
    OptComprobante(1).Enabled = True
    OptTipo(1).Value = False
    OptTipo(0).Value = False
    OptTipo(1).Value = True
    
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    CalFecha.Value = Date 'ValidarPeriodo(Date, False)
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)
    Call CargarCmbPaquetes(CmbPaquetes)
    Call CargarVecCuentas(CentroEmisor)
    Call CargarCmbCentrosDeCostos(CmbCentrosDeCostos)
    'Call CargarComboCuentasContables(CmbCuentas)
    ReDim Preserve VecCentroCta(0)
    ReDim VecAutorizacionesAnticiposApli(0)

    LvCenCostoCtas.ListItems.Add
    Modificado = False
    CalDesde.Value = Date
    CalHasta.Value = Date
    OptComprobante(0).Value = True
End Sub

Private Sub OptComprobante_Click(Index As Integer)
    TxtFactura.Visible = Index = 1
    'LbFactura.Visible = Index = 1
    'TxtTotalFactura.Visible = Index = 1
    'TxtPrecioU.Enabled = Index <> 1
    'TxtTotalFactura.Enabled = True
    TxtFactura = ""
    'TxtTotalFactura = ""
    LbTotalIva.Visible = Index = 1
    TxtIVA.Visible = Index = 1
    CmdAplicarAnticipo.Visible = Index = 1
    If Index = 0 Then
       LbPrecio.Caption = "Imp. Anticipo"
    Else
        LbPrecio.Caption = "Precio sin IVA"
    End If
    
End Sub

Private Sub OptTipo_Click(Index As Integer)
    TxtFileNro.Enabled = Index = 1
    TxtFileNro = 0
    CmbPaquetes.Enabled = Index = 0
    CalDesde.Enabled = Index = 0
    CalHasta.Enabled = Index = 0
    CmbPaquetes.ListIndex = 0
    CalDesde.Value = Date
    CalHasta.Value = Date
End Sub

Private Sub TxtFileNro_KeyPress(KeyAscii As Integer)
       If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
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

