VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Begin VB.Form A01_6100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Orden de Contratación"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMulti 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   45
      TabIndex        =   52
      Top             =   6075
      Width           =   11805
      Begin VB.CommandButton CmdCrear 
         Caption         =   "Cr&ear"
         Height          =   350
         Left            =   10350
         TabIndex        =   46
         Top             =   1710
         Width           =   1300
      End
      Begin VB.TextBox TxtCantOrdenes 
         Height          =   315
         Left            =   90
         TabIndex        =   44
         Text            =   "12"
         Top             =   1035
         Width           =   780
      End
      Begin Controles.ComboEsp CmbFrecuencia 
         Height          =   315
         Left            =   90
         TabIndex        =   43
         Top             =   450
         Width           =   1950
         _ExtentX        =   3440
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
      Begin MSComctlLib.ListView LvMeses 
         Height          =   1905
         Left            =   2430
         TabIndex        =   45
         Top             =   180
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad de órdenes"
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
         TabIndex        =   54
         Top             =   810
         Width           =   2220
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frecuencia:"
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
         TabIndex        =   53
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4905
      TabIndex        =   51
      Top             =   5760
      Width           =   1230
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   3510
      TabIndex        =   48
      Top             =   5745
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10620
      TabIndex        =   42
      Top             =   5745
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7740
      TabIndex        =   40
      Top             =   5745
      Width           =   1230
   End
   Begin VB.CommandButton CmdNueva 
      Caption         =   "&Nueva"
      Height          =   350
      Left            =   6300
      TabIndex        =   39
      Top             =   5745
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   630
      Top             =   5490
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación"
      Height          =   2940
      Left            =   8190
      TabIndex        =   32
      Top             =   2565
      Width           =   3660
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1080
         Width           =   645
      End
      Begin VB.CommandButton CmdBuscarSubCentro 
         Caption         =   "Buscar"
         Height          =   300
         Left            =   2835
         TabIndex        =   16
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox TxtPrecioU 
         Height          =   315
         Left            =   2655
         TabIndex        =   19
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton CmdAgregarCentroCta 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   1935
         TabIndex        =   21
         Top             =   1980
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminarCentroCta 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   405
         TabIndex        =   22
         Top             =   2430
         Width           =   1300
      End
      Begin VB.CommandButton CmdModifCentroCta 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   405
         TabIndex        =   20
         Top             =   1980
         Width           =   1300
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   1125
         Width           =   2715
         _ExtentX        =   4789
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
      Begin Controles.ComboEsp CmbCentrosDeCostos 
         Height          =   330
         Left            =   90
         TabIndex        =   15
         Top             =   495
         Width           =   2715
         _ExtentX        =   4789
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
         Left            =   2880
         TabIndex        =   55
         Top             =   855
         Width           =   405
      End
      Begin VB.Label LbCant 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe sin IVA:"
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
         Left            =   1260
         TabIndex        =   35
         Top             =   1620
         Width           =   1365
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
         Top             =   225
         Width           =   1830
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
         Left            =   90
         TabIndex        =   33
         Top             =   855
         Width           =   1485
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
      Left            =   7020
      TabIndex        =   31
      Text            =   "0"
      Top             =   5385
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   2760
      Left            =   45
      TabIndex        =   14
      Top             =   2565
      Width           =   8070
      _ExtentX        =   14235
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Orden de Contratación"
      Height          =   2445
      Left            =   45
      TabIndex        =   23
      Top             =   45
      Width           =   11850
      Begin VB.OptionButton OptMulti 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Multiples Períodos"
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
         Index           =   1
         Left            =   3960
         TabIndex        =   13
         Top             =   2115
         Width           =   1950
      End
      Begin VB.OptionButton OptMulti 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Para Este Período"
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
         Index           =   0
         Left            =   1485
         TabIndex        =   12
         Top             =   2115
         Value           =   -1  'True
         Width           =   2445
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1755
         Width           =   10230
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   2565
         TabIndex        =   1
         Top             =   270
         Width           =   900
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   3555
         TabIndex        =   2
         Top             =   270
         Width           =   900
      End
      Begin VB.TextBox TxtLugar 
         Height          =   315
         Left            =   7065
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1035
         Width           =   4650
      End
      Begin VB.TextBox TxtResp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   5
         Top             =   675
         Width           =   3480
      End
      Begin VB.TextBox TxtNroOrden 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   270
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   8460
         TabIndex        =   3
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   103546883
         UpDown          =   -1  'True
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1395
         Width           =   5100
         _ExtentX        =   8996
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
      Begin Controles.ComboEsp CmbEmp 
         Height          =   315
         Left            =   8325
         TabIndex        =   10
         Top             =   1395
         Width           =   3390
         _ExtentX        =   5980
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
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   7065
         TabIndex        =   6
         Top             =   675
         Width           =   4695
         _ExtentX        =   8281
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
      Begin MSComCtl2.DTPicker CalFechaEmitida 
         Height          =   330
         Left            =   10485
         TabIndex        =   4
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   103546881
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbFormaDePago 
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Top             =   1035
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período Imputacion:"
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
         Left            =   6660
         TabIndex        =   24
         Top             =   315
         Width           =   1740
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
         Left            =   9540
         TabIndex        =   56
         Top             =   345
         Width           =   915
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
         Left            =   90
         TabIndex        =   50
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label LBPerCerrado 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período Cerrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Left            =   4545
         TabIndex        =   49
         Top             =   405
         Visible         =   0   'False
         Width           =   2280
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
         TabIndex        =   38
         Top             =   135
         Visible         =   0   'False
         Width           =   2325
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
         Left            =   4950
         TabIndex        =   37
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lugar del Servicio:"
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
         Left            =   5385
         TabIndex        =   36
         Top             =   1080
         Width           =   1620
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
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Orden:"
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
         TabIndex        =   28
         Top             =   330
         Width           =   1125
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
         TabIndex        =   27
         Top             =   720
         Width           =   1170
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
         TabIndex        =   26
         Top             =   1080
         Width           =   1350
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
         Left            =   6435
         TabIndex        =   25
         Top             =   1440
         Width           =   1860
      End
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   9180
      TabIndex        =   41
      Top             =   5745
      Width           =   1300
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   9180
      TabIndex        =   47
      Top             =   5745
      Visible         =   0   'False
      Width           =   1300
   End
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   1440
      Top             =   5475
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   2205
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      Left            =   6480
      TabIndex        =   30
      Top             =   5445
      Width           =   510
   End
End
Attribute VB_Name = "A01_6100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoCentroCta
   O_CuentaContable As String
   CentroPadre As String
   Cta_Descripcion As String
   O_CentroDeCosto As String
   Centro_Descripcion As String
   O_PrecioPactado As Double
   O_PendienteAutorizacionDePago As Boolean
   O_SinPresupuestar As Boolean
   O_MontoSinPresupuestar As Double
End Type

Private VecCentroCta() As TipoCentroCta

Private TotalMontoSinPres As Double
Private MontoSinPres As Double
Private SinPres As Boolean

Private Modificado As Boolean
Public NroOrden As Integer
Private Nivel As Integer
Public CentroEmisorActual As String
Private VecCuentasContables() As CuentasContables

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalFecha_Change()
    CalFecha.Value = ValidarPeriodo(CalFecha.Value)
    txtTotal = "0"
     
    LvCenCostoCtas.ListItems.Clear
    LvMeses.ListItems.Clear
    OptMulti(0).Value = True
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
    ReDim VecCentroCta(0)

End Sub

Private Sub CmbCentroDeCostoEmisor_Click()
'    Modificado = True
    txtTotal = "0"
     
    LvCenCostoCtas.ListItems.Clear
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    ReDim VecCentroCta(0)
    Call CargarVecCentroEmisor(CentroEmisorActual)
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)

End Sub

Private Sub CmbCuentas_Click()
    TxtCodCuenta.Text = VecCuentasContables(CmbCuentas.ListIndex).Codigo
End Sub

Private Sub CmbEmp_Click()
    Modificado = True
End Sub

Private Sub CmbFormaPago_Change()
    Modificado = True
End Sub

Private Sub CmbFrecuencia_Click()
    Select Case CmbFrecuencia.ListIndex
    Case 0
        TxtCantOrdenes.Text = "12"
    Case 1
        TxtCantOrdenes.Text = "6"
    Case 2
        TxtCantOrdenes.Text = "4"
    Case 3
        TxtCantOrdenes.Text = "3"
    Case 4
        TxtCantOrdenes.Text = "2"
    End Select
End Sub

Private Sub CmbProv_Click()
    Modificado = True
End Sub

Private Sub CmdAgregarCentroCta_Click()
'On Error GoTo errores
Dim i As Integer
 
 SinPres = False
 
 If ValidarCargaCentroCta Then
    Modificado = True
        
    If LvCenCostoCtas.SelectedItem.Index = LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Index Then
        
       'agrega al vector
        i = LvCenCostoCtas.SelectedItem.Index
        ReDim Preserve VecCentroCta(UBound(VecCentroCta) + 1)
        
        VecCentroCta(UBound(VecCentroCta)).Centro_Descripcion = Trim(CmbCentrosDeCostos.Text)
        VecCentroCta(UBound(VecCentroCta)).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
        VecCentroCta(UBound(VecCentroCta)).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
        VecCentroCta(UBound(VecCentroCta)).Cta_Descripcion = Trim(CmbCuentas.Text)
        VecCentroCta(UBound(VecCentroCta)).O_PrecioPactado = Val(TxtPrecioU.Text)
        VecCentroCta(UBound(VecCentroCta)).O_PendienteAutorizacionDePago = True
        VecCentroCta(UBound(VecCentroCta)).CentroPadre = BuscarDescCentroEmisor(BuscarCentroPadre(VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo))
        VecCentroCta(UBound(VecCentroCta)).O_SinPresupuestar = SinPres
        VecCentroCta(UBound(VecCentroCta)).O_MontoSinPresupuestar = MontoSinPres

        TotalMontoSinPres = CalcularTotalSinPresupuestar

        'lo pone en el LV
        With VecCentroCta(UBound(VecCentroCta))
           LvCenCostoCtas.ListItems(i).Text = .CentroPadre
           LvCenCostoCtas.ListItems(i).SubItems(1) = .Centro_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(2) = .Cta_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.O_PrecioPactado, "0.00")
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
     Call CalcularTotal
  End If
    TotalMontoSinPres = CalcularTotalSinPresupuestar

End Sub

Private Function ValidarCargaCentroCta() As Boolean
On Error GoTo Error
    ValidarCargaCentroCta = True
    Dim i As Integer
    Dim Index As Integer
    MontoSinPres = 0
    'recorro el lv para ver si no está la combinación de centro - cta
  If CmbCentroDeCostoEmisor.ListIndex = 0 Then
     MsgBox "Debe Seleccionar un Centro de Costo Emisor"
     CmbCentroDeCostoEmisor.SetFocus
     ValidarCargaCentroCta = False
     Exit Function
  End If
    
    For i = 1 To LvCenCostoCtas.ListItems.Count - 1
      If i <> LvCenCostoCtas.SelectedItem.Index Then
        'Index = LvCenCostoCtas.SelectedItem.Index
       If VecCentroCta(i).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo _
          And VecCentroCta(i).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
          
            MsgBox "Ya existe esta Combinación de Centro de Costo - Cuenta", , "Orden de Contratacion"
            ValidarCargaCentroCta = False
            Exit Function
       End If
      End If
    Next
    
    If CmbCentrosDeCostos.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Centro de Costo"
       CmbCentrosDeCostos.SetFocus
       ValidarCargaCentroCta = False
    Else
        If CmbCuentas.ListIndex = 0 Then
            MsgBox "Debe Seleccionar una Cuenta"
            CmbCuentas.SetFocus
            ValidarCargaCentroCta = False
        Else
           If Val(TxtPrecioU.Text) = 0 Then
              MsgBox "Debe Ingresar un Importe mayor a 0"
              TxtPrecioU.SetFocus
              ValidarCargaCentroCta = False
           End If
        End If
    End If

    Dim Sql As String
    Dim Rta As Integer
    Dim ImportePres As Double
    Dim RsValidar As New ADODB.Recordset
    Dim MontoDisponible As Double
    
    If VecCuentasContables(CmbCuentas.ListIndex).Codigo = "5121" Then
         Sql = "SpOCPresupuestosDistribucionValidarServicio" & _
                     "   @CuentaContable ='" & VecCuentasContables(CmbCuentas.ListIndex).Codigo & _
                     "', @CentroEmisor ='" & CentroEmisor & _
                     "', @Periodo ='" & Format(CalFecha, "MM/yyyy") & _
                     "', @NroOrden=" & NroOrden & _
                     " , @SubCentroDeCosto='" & VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo & "'"
                     
         RsValidar.Open Sql, Conec
         'acá está el problema
         MontoDisponible = ValN(RsValidar!MontoDisponible)  'TotalMontoSinPres
        
         If MontoDisponible < Val(TxtPrecioU.Text) Then
            MsgBox "El monto disponible para este período es de $" & Format(MontoDisponible, "0.00"), vbInformation
            ValidarCargaCentroCta = False
            Exit Function
        End If
    Else
        Sql = "SpOCPresupuestosRenglonesValidarContratacion @CuentaContable = '" & VecCuentasContables(CmbCuentas.ListIndex).Codigo & _
                                                        "', @NroOrden = " & NroOrden & _
                                                         ", @Periodo ='" & Format(CalFecha.Value, "MM/yyyy") & _
                                                        "', @CentroEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
        RsValidar.Open Sql, Conec
        'si se hace una sola orden se puede ing serv sin presupuestar
      If OptMulti(0).Value Then
         With RsValidar
            If Val(TxtPrecioU.Text) > !MontoDisponible - CalcularTotalCuenta(VecCuentasContables(CmbCuentas.ListIndex).Codigo) Then
                Rta = MsgBox("El Importe aprobada para esa Cuenta Contable de $ " & RsValidar!MontoDisponible & _
                             " Para el Período " & Format(CalFecha.Value, "MMMM/yyyy") & _
                             " ¿Desea ingresar el Servicio sin presupuestar?", vbYesNo, "Importe")
                ImportePres = !MontoDisponible - CalcularTotalCuenta(VecCuentasContables(CmbCuentas.ListIndex).Codigo)
                .Close
                SinPres = False
                
                If Rta = vbYes Then
                    
                   Sql = "SpOCImporteSinPresupuestarContratacion @CentroDeCosto ='" & CentroEmisorActual & _
                                                              "', @NroOrden =" & NroOrden & _
                                                              " , @Periodo =" & FechaSQL(CalFecha.Value, "SQL")
                   .Open Sql, Conec
                   
                   If Val(TxtPrecioU) - ImportePres > !MontoSinPresupuestarMensual - TotalMontoSinPres - !MontoSinPres Then
                        MsgBox "Solo está autorizado por un monto de: $" & !MontoSinPresupuestarMensual - TotalMontoSinPres - !MontoSinPres
                                 
                        TxtPrecioU.SetFocus
                        ValidarCargaCentroCta = False
                        Exit Function
                   Else
                        MontoSinPres = Val(TxtPrecioU) - ImportePres
                        SinPres = True
    
                        Exit Function
                   End If
        
                End If
        
                TxtPrecioU.SetFocus
                ValidarCargaCentroCta = False
                Exit Function
            End If
         End With
      Else
       'si se hacen multiples ordenes
         'If Val(TxtPrecioU.Text) > RsValidar!MontoDisponible Then
         '   MsgBox "El Importe aprobada para esa Cuenta Contable/Centro de Costo es de $ " & RsValidar!MontoDisponible & _
                   " Para el Período " & Format(CalFecha.Value, "MMMM/yyyy"), vbInformation, "Importe"
                        
         '   TxtPrecioU.SetFocus
         '   ValidarCargaCentroCta = False
         '   Exit Function
         'End If
    
      End If
    End If
Error:
    If Err.Number <> 0 Then
       Call ManipularError(Err.Number, Err.Description)
       ValidarCargaCentroCta = False
    End If
End Function

Private Sub CmdAnular_Click()
 Dim Sql As String
 Dim Rta As Integer
 Dim RsAnular As ADODB.Recordset
 Set RsAnular = New ADODB.Recordset
 
 On Error GoTo Error
    Rta = MsgBox("¿Está seguro de que desea Anular La Orden de Contratación?", vbYesNo)
    If Rta = vbYes Then
        Sql = "SpOCOrdenesDeContratacionCabeceraAnular @O_NumeroOrdenDeContratacion =" & NroOrden & _
                                                    ", @O_CentroDeCostoEmisor ='" & CentroEmisorActual & "'"
        RsAnular.Open Sql, Conec
        If RsAnular!Ok = "OK" Then
            MsgBox RsAnular!Mensaje, vbInformation
        Else
            MsgBox RsAnular!Mensaje, vbExclamation
        End If
    End If
Error:

  If Err.Number <> 0 Then
     Call ManipularError(Err.Number, Err.Description)
  Else
     Rta = MsgBox("¿Desea realizar otra acción?", vbYesNo)
     If Rta = vbYes Then
        Call LimpiarOrden
     Else
        Unload Me
     End If
  End If
End Sub

Private Sub CMDBuscar_Click()
    NroOrden = 0
    BuscarOrdenDeContratacion.LbCentroEmisor.Visible = Nivel = 2
    BuscarOrdenDeContratacion.CmbCentroDeCostoEmisor.Visible = Nivel = 2
    BuscarOrdenDeContratacion.Show vbModal
    Timer1.Enabled = True
End Sub

Private Sub CmdBuscarSubCentro_Click()
    BuscarSubCentro.Show vbModal
    Call BuscarCentro(BuscarSubCentro.CodigoSubCentro, CmbCentrosDeCostos)

End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar la Orden de Contratacion?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarOrden
    End If
End Sub

Private Function ModificarOrden() As Boolean
  Dim Sql As String
  Dim i As Integer
  Dim Precio As String
  
On Error GoTo ErrorUpdate

  NroOrden = Val(TxtNroOrden.Text)
  ModificarOrden = False
  
  If ValidarEncabezado Then
     ModificarOrden = True
    
    Conec.BeginTrans
     Sql = "SpOCOrdenesDeContratacionCabeceraModificar @O_NumeroOrdenDeContratacion =" + CStr(NroOrden) + _
           " , @O_Fecha=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
           " , @O_Responsable ='" + TxtResp.Text & _
           "', @O_CodigoProveedor ='" + CStr(VecProveedores(CmbProv.ListIndex).Codigo) + _
           "', @O_LugarDelServicio = '" + TxtLugar.Text & _
           "', @O_FormaDePagoPactada ='" + CmbFormaDePago.Text & _
           "', @O_EmpresaFacturaANombreDe ='" + VecEmpresas(CmbEmp.ListIndex).Codigo + _
           "', @U_Usuario = '" + Usuario + _
           "', @O_CentroDeCostoEmisor= '" & CentroEmisorActual & _
           "', @O_Observaciones = '" & TxtObs & _
           "', @O_FechaEmision = " & FechaSQL(CalFechaEmitida, "SQL") & _
           " , @O_CodigoFormaDePago=" & VecFormasDePago(CmbFormaDePago.ListIndex).F_Codigo & _
           " , @O_Autorizado=" & IIf(ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion, 0, 1)
           
        Conec.Execute Sql
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
      
        Sql = "SpOCOrdenesDeContratacionRenglonesAgregar @O_NumeroOrdenDeContratacion = " & NroOrden & _
                                                      ", @O_CuentaContable = '" & .O_CuentaContable & _
                                                     "', @O_CentroDeCosto = '" & .O_CentroDeCosto & _
                                                     "', @O_PrecioPactado = " & Replace(.O_PrecioPactado, ",", ".") & _
                                                      ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & _
                                                     "', @O_SinPresupuestar = " & IIf(.O_SinPresupuestar, 1, 0) & _
                                                      ", @O_MontoSinPresupuestar=" & Replace(.O_MontoSinPresupuestar, ",", ".")
      End With
        Conec.Execute Sql
    Next
    Conec.CommitTrans
ErrorUpdate:
    If Err.Number = 0 Then
        If ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion Then
            MsgBox "La orden necesita autorización", vbInformation, "Supera los $" & MaxSinAutorizacion
            FrmMensaje.CmdImprimir.Enabled = False
            FrmMensaje.CmdExportar.Enabled = False
            Call EnviarMailAutorizacion(NroOrden)
        Else
            MsgBox "La Orden de Contratacion se Grabó correctamente"
        End If
       Modificado = False
    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Function

Private Sub CmdCargar_Click()
    Call CargarOrden(Val(TxtNroOrden))
    Modificado = False

End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar la Orden de Contratación?", vbYesNo)
    If Rta = vbYes Then
        Call GrabarOrden
    End If
End Sub

Private Function GrabarOrden() As Boolean
  Dim Sql As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

On Error GoTo ErrorInsert
  GrabarOrden = False
  If ValidarEncabezado Then
   'retorna verdadero si se puede graba bien
       GrabarOrden = True

    Conec.BeginTrans
     Sql = "SpOCOrdenesDeContratacionCabeceraAgregar @O_Fecha=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
            ", @O_Responsable ='" + TxtResp.Text + _
           "', @O_CodigoProveedor ='" + CStr(VecProveedores(CmbProv.ListIndex).Codigo) + _
           "', @O_LugarDelServicio = '" & TxtLugar.Text & _
           "', @O_FormaDePagoPactada ='" & CmbFormaDePago.Text & _
           "', @O_EmpresaFacturaANombreDe ='" & VecEmpresas(CmbEmp.ListIndex).Codigo + _
           "', @U_Usuario = '" & Usuario & _
           "', @O_CentroDeCostoEmisor= '" & CentroEmisorActual & _
           "', @O_Observaciones = '" & TxtObs & _
           "', @O_FechaEmision = " & FechaSQL(CalFechaEmitida, "SQL") & _
           " , @O_CodigoFormaDePago=" & VecFormasDePago(CmbFormaDePago.ListIndex).F_Codigo & _
           " , @O_Autorizado=" & IIf(ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion, 0, 1)
     'graba el encabezado y retorna el Nro de Orden
        RsGrabar.Open Sql, Conec
        NroOrden = RsGrabar!O_NumeroOrdenDeContratacion
        
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
      
        Sql = "SpOCOrdenesDeContratacionRenglonesAgregar @O_NumeroOrdenDeContratacion = " & NroOrden & _
                                                ", @O_CuentaContable = '" & .O_CuentaContable & _
                                               "', @O_CentroDeCosto = '" & .O_CentroDeCosto & _
                                               "', @O_PrecioPactado = " & Replace(.O_PrecioPactado, ",", ".") & _
                                                ", @O_CentroDeCostoEmisor ='" & CentroEmisorActual & _
                                               "', @O_SinPresupuestar = " & IIf(.O_SinPresupuestar, 1, 0) & _
                                                ", @O_MontoSinPresupuestar = " & Replace(.O_MontoSinPresupuestar, ",", ".")
      End With
        Conec.Execute Sql
    Next
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       CmdConfirmar.Visible = False
       CmdCambiar.Visible = True
       CmdImprimir.Enabled = True
       Modificado = False
       
        If ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion Then
            MsgBox "La orden necesita autorización", vbInformation, "Supera los $" & MaxSinAutorizacion
            FrmMensaje.CmdImprimir.Enabled = False
            FrmMensaje.CmdExportar.Enabled = False
            Call EnviarMailAutorizacion(NroOrden)
       End If

       FrmMensaje.LbMensaje = "La Orden de Contratación se Grabó correctamente con el Nº: " + CStr(NroOrden) & _
                              Chr(13) & " ¿Que desea hacer?"
       FrmMensaje.Show vbModal
       
       Modificado = False
       If FrmMensaje.Retorno = vbimprimir Then
            Call ConfImpresionDeOrden
            RepOrdenDeContratacion.Show vbModal
       End If
         
       If FrmMensaje.Retorno = vbNuevo Then
         Call LimpiarOrden
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

Private Sub EnviarMailAutorizacion(NroOrden As Integer)
Dim Mensaje As String
Dim i As Integer
Dim EMail As String
Dim inicio As Integer
Dim Fin As Integer
'omite los errores
On Error GoTo ErrorEMail

    EMail = BuscarEMailAutorizacionCentroEmisor(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo)
    If EMail = "" Then
       MsgBox "El Centro de costo no tiene configurado E-Mail de Autorización", vbInformation
       Exit Sub
    End If
    
    Call ConfImpresionDeOrden
    RepOrdenDeContratacion.Run
    'guarda la orden de compra como un PDF
    Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
    Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
    myPDFExport.AcrobatVersion = DDACR40
     
    myPDFExport.Filename = "C:\Orden " & TxtNroOrden.Text & " " & CmbCentroDeCostoEmisor.Text & ".pdf"
    myPDFExport.JPGQuality = 100
    myPDFExport.SemiDelimitedNeverEmbedFonts = ""
    myPDFExport.Export RepOrdenDeContratacion.Pages
    Unload RepOrdenDeContratacion
    
    MAPISession.DownLoadMail = False
    MAPISession.SignOn
    MAPIMessages.SessionID = MAPISession.SessionID
    MAPIMessages.Compose
    'MAPIMessages.RecipAddress = VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail
    
    '******** varias E-Mail **********
    
    MAPIMessages.RecipIndex = i
    MAPIMessages.RecipType = 1
    MAPIMessages.RecipAddress = EMail
    '****** Fin Varios E-Mail **********
    
    'MAPIMessages.AddressResolveUI = False
    'MAPIMessages.ResolveName
    MAPIMessages.MsgSubject = "Autorizar Orden de Contratación Nº: " & Format(NroOrden, "0000000000") & " ,Centro de Costo Emisor: " & CmbCentroDeCostoEmisor.Text
    
    'MAPIMessages.MsgNoteText = Mensaje
   
    MAPIMessages.AttachmentPathName = myPDFExport.Filename
    MAPIMessages.Send False
    MAPISession.SignOff
    
ErrorEMail:
    'If Err.Number <> 0 Then
        Call EnviarEmail(EMail, EMail, "Autorizar Orden de Contratación Nº: " & Format(NroOrden, "0000000000") & " ,Centro de Costo Emisor: " & CmbCentroDeCostoEmisor.Text, "", "C:\Orden " & TxtNroOrden.Text & " " & CmbCentroDeCostoEmisor.Text & ".pdf")
    'End If
    Call Kill(myPDFExport.Filename)
End Sub



Private Function ValidarEncabezado() As Boolean
Dim i As Integer
Dim Asignado As Boolean

    ValidarEncabezado = True
    
    If CalFecha.CustomFormat = " " Then
        MsgBox "Debe Período Fecha de Imputación"
        CalFecha.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If

    If CmbFormaDePago.ListIndex = 0 Then
       MsgBox "Debe Seleccionar una Forma de Pago"
       CmbFormaDePago.SetFocus
       ValidarEncabezado = False
       Exit Function
    Else
      If TxtLugar.Text = "" Then
         MsgBox "Debe ingresar el Lugar del Servicio"
         TxtLugar.SetFocus
         ValidarEncabezado = False
         Exit Function
      Else
        If CmbProv.ListIndex = 0 Then
           MsgBox "Debe Seleccionar un Proveedor"
           CmbProv.SetFocus
           ValidarEncabezado = False
           Exit Function
          End If
        End If
    End If

    If LvCenCostoCtas.ListItems.Count <= 1 Then
       MsgBox "La órden debe tener algún Servico", vbExclamation, "Servicio"
       LvCenCostoCtas.ListItems(1).Selected = True
       Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
       ValidarEncabezado = False
       Exit Function
       
    End If
    
    Dim Sql As String
    Dim RsValidarPeriodo As ADODB.Recordset
    Set RsValidarPeriodo = New ADODB.Recordset
    Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(CalFecha.Value, "MM/yyyy")) & "'"
    RsValidarPeriodo.Open Sql, Conec
    
    If RsValidarPeriodo!Cerrado > 0 Then
       MsgBox "El período está Cerrado", vbExclamation, "Período Cerrado"
       CalFecha.SetFocus
       ValidarEncabezado = False
       Exit Function
    End If
        
    'si se crean multiples órdenes se debe verificar que exista por lo menos una
    Dim Chekeo As Boolean
    If OptMulti(1).Value Then
        For i = 1 To LvMeses.ListItems.Count
            If LvMeses.ListItems(i).Checked Then
                Chekeo = True
                Exit For
            End If
        Next
        If Not Chekeo Then
           MsgBox "Dese Selecconar al menos un Período", vbExclamation, "Período"
           LvMeses.SetFocus
           ValidarEncabezado = False
           Exit Function
        End If
    End If
End Function

Private Sub CmdCrear_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar la Multiples Ordenes de Contratación?", vbYesNo)
    If Rta = vbYes Then
        Call CrearMultiOrdenes
    End If

End Sub

Private Function CrearMultiOrdenes() As Boolean
  Dim Sql As String
  Dim Mensaje As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer
  Dim j As Integer
  Dim Precio As String

'On Error GoTo ErrorInsert
  CrearMultiOrdenes = False
  If ValidarEncabezado Then
    CrearMultiOrdenes = True
    
    Conec.BeginTrans
    
    For j = 1 To LvMeses.ListItems.Count
      If LvMeses.ListItems(j).Checked Then
         Sql = "SpOCOrdenesDeContratacionCabeceraAgregar @O_Fecha=" & FechaSQL(Format(Day(CalFecha.Value) & "/" & LvMeses.ListItems(j).Text, "DD/MM/YYYY"), "SQL") & _
                ", @O_Responsable ='" & TxtResp.Text & _
               "', @O_CodigoProveedor ='" & VecProveedores(CmbProv.ListIndex).Codigo & _
               "', @O_LugarDelServicio = '" & TxtLugar.Text & _
               "', @O_FormaDePagoPactada ='" & CmbFormaDePago.Text & _
               "', @O_EmpresaFacturaANombreDe ='" & VecEmpresas(CmbEmp.ListIndex).Codigo & _
               "', @U_Usuario = '" & Usuario & _
               "', @O_CentroDeCostoEmisor= '" & CentroEmisorActual & _
               "', @O_Observaciones = '" & TxtObs & _
               "', @O_FechaEmision =" & FechaSQL(CalFechaEmitida, "SQL") & _
               " , @O_CodigoFormaDePago=" & VecFormasDePago(CmbFormaDePago.ListIndex).F_Codigo & _
               " , @O_Autorizado=" & IIf(ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion, 0, 1)
               
         'graba el encabezado y retorna el Nro de Orden
            RsGrabar.Open Sql, Conec
            NroOrden = RsGrabar!O_NumeroOrdenDeContratacion
            RsGrabar.Close
            Mensaje = Mensaje & " - " & LvMeses.ListItems(j).Text & " Nº: " & Format(NroOrden, "0000000000")
        For i = 1 To UBound(VecCentroCta)
          With VecCentroCta(i)
          
            Sql = "SpOCOrdenesDeContratacionRenglonesAgregar @O_NumeroOrdenDeContratacion = " & NroOrden & _
                                    ", @O_CuentaContable = '" & .O_CuentaContable & _
                                   "', @O_CentroDeCosto = '" & .O_CentroDeCosto & _
                                   "', @O_PrecioPactado = " & Replace(.O_PrecioPactado, ",", ".") & _
                                    ", @O_CentroDeCostoEmisor ='" & CentroEmisorActual & _
                                   "', @O_SinPresupuestar = " & IIf(.O_SinPresupuestar, 1, 0) & _
                                    ", @O_MontoSinPresupuestar = " & .O_MontoSinPresupuestar

          End With
            Conec.Execute Sql
        Next
      End If
    Next
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       
       Rta = MsgBox("Las Ordenes de Contratación se Grabaron correctamente " & Mensaje & " ¿Desea Crear otra Orden de Contratación?", vbYesNo)
       Modificado = False
        
      
      If Rta = vbYes Then
         Call LimpiarOrden
         CrearMultiOrdenes = False
      Else
         Unload Me
      End If

    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Function

Private Sub CmdEliminarCentroCta_Click()
    Dim Index As Integer
   
    Index = LvCenCostoCtas.SelectedItem.Index
   If VecCentroCta(Index).O_PendienteAutorizacionDePago Then
      If VecCentroCta(Index).O_SinPresupuestar Then
         TotalMontoSinPres = TotalMontoSinPres - VecCentroCta(Index).O_PrecioPactado
      End If
    'borra del LV
    LvCenCostoCtas.ListItems.Remove (Index)
    'borrar del vector
    While Index < UBound(VecCentroCta)
        VecCentroCta(Index) = VecCentroCta(Index + 1)
        Index = Index + 1
    Wend
        
    ReDim Preserve VecCentroCta(UBound(VecCentroCta) - 1)
      'reposiciona el lv
      LvCenCostoCtas.ListItems(Index).Selected = True
      
    If LvCenCostoCtas.ListItems.Count = LvCenCostoCtas.SelectedItem.Index Then
        CmdEliminarCentroCta.Enabled = False
        CmdModifCentroCta.Enabled = False
    End If
        
    Modificado = True
    Call CalcularTotal
 Else
    MsgBox "El Servicio ya fue autorizado"
 End If
     TotalMontoSinPres = CalcularTotalSinPresupuestar

End Sub

Private Sub ConfImpresionDeOrden()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "Centro", adVarChar, 100
    RsListado.Fields.Append "CentroPadre", adVarChar, 100
    RsListado.Fields.Append "Importe", adDouble
    RsListado.Open
    i = 1
    For i = 1 To UBound(VecCentroCta)
        RsListado.AddNew
      With VecCentroCta(i)
        RsListado!Centro = .Centro_Descripcion
        RsListado!CentroPadre = .CentroPadre
        RsListado!Cuenta = .Cta_Descripcion
        RsListado!Importe = .O_PrecioPactado
        
      End With
    Next
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If

    TxtNroOrden.Text = Format(NroOrden, "0000000000")
    
    RepOrdenDeContratacion.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    RepOrdenDeContratacion.TxtFormaDePago.Text = CmbFormaDePago.Text
    RepOrdenDeContratacion.TxtFecha = Format(CalFecha.Value, "MM/yyyy")
    RepOrdenDeContratacion.TxtFEmision = CalFechaEmitida
    
    RepOrdenDeContratacion.LbFactura.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeContratacion.TxtFactNombre.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeContratacion.TxtFactNombre.Text = CmbEmp.Text
    RepOrdenDeContratacion.LbCUIT.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeContratacion.TxtCuit.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeContratacion.TxtCuit.Text = VecEmpresas(CmbEmp.ListIndex).CUIT
    
    RepOrdenDeContratacion.TxtNroOrden.Text = TxtNroOrden.Text
    RepOrdenDeContratacion.TxtLugarDeEntrega = TxtLugar.Text
    RepOrdenDeContratacion.TxtProv.Text = CmbProv.Text
    RepOrdenDeContratacion.TxtResp.Text = TxtResp.Text
    RepOrdenDeContratacion.TxtAnulada.Visible = LBAnulada.Visible
    RepOrdenDeContratacion.TxtAnulada.Text = LBAnulada.Caption
    RepOrdenDeContratacion.TxtObservaciones.Text = TxtObs
    
    RepOrdenDeContratacion.DataControl1.Recordset = RsListado
    RepOrdenDeContratacion.Zoom = -1
End Sub

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeOrden
         RepOrdenDeContratacion.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepOrdenDeContratacion.Pages
         Unload RepOrdenDeContratacion
  End If
Error:
    If Err.Number = 0 Then
        MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeOrden
    RepOrdenDeContratacion.Show
End Sub

Private Sub CmdModifCentroCta_Click()
On Error GoTo Errores
  Dim i As Integer
  Dim PosVec As Integer
  i = LvCenCostoCtas.SelectedItem.Index
  TotalMontoSinPres = TotalMontoSinPres - IIf(VecCentroCta(i).O_SinPresupuestar, VecCentroCta(i).O_MontoSinPresupuestar, 0)
  SinPres = False
 If VecCentroCta(i).O_PendienteAutorizacionDePago Then
   If ValidarCargaCentroCta Then
        
       'agrega al vector
       PosVec = i
       Modificado = True
        With VecCentroCta(PosVec)
           .Centro_Descripcion = Trim(CmbCentrosDeCostos.Text)
           .O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
           .O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
           .CentroPadre = BuscarDescCentroEmisor(BuscarCentroPadre(VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo))
           .Cta_Descripcion = Trim(CmbCuentas.Text)
           .O_PrecioPactado = Val(TxtPrecioU.Text)
           .O_SinPresupuestar = SinPres
           .O_MontoSinPresupuestar = MontoSinPres

       'lo pone en el LV
           LvCenCostoCtas.ListItems(i).Text = .CentroPadre
           LvCenCostoCtas.ListItems(i).SubItems(1) = .Centro_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(2) = .Cta_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.O_PrecioPactado, "0.00")
        End With
      Call CalcularTotal
    End If
 Else
    MsgBox "El Servicio ya fue autorizado"
 End If
    TotalMontoSinPres = CalcularTotalSinPresupuestar

Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdNueva_Click()
    CentroEmisorActual = CentroEmisor
    TxtNroOrden.Text = ""
    CmdConfirmar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
    'esto es por que cuando se trae una orden anulada los
    'valores están invertidos
    LBAnulada.Visible = False
    FrameAsig.Enabled = True
    CmdCambiar.Enabled = True
    TxtLugar.Enabled = True
    CmbEmp.Enabled = True
    CmbFormaDePago.Enabled = True
    CmbProv.Enabled = True
    CalFecha.Enabled = True
    TxtObs.Enabled = True
    LBPerCerrado.Visible = False
    TotalMontoSinPres = 0
End Sub

Private Sub LvMeses_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim Sql As String
    Dim i As Integer
    Dim RsValidar As ADODB.Recordset
    Set RsValidar = New ADODB.Recordset
    
  For i = 1 To UBound(VecCentroCta)
      Sql = "SpOCPresupuestosRenglonesValidarContratacion @CuentaContable = '" & VecCentroCta(i).O_CuentaContable & _
                                                      "', @NroOrden = " & NroOrden & _
                                                       ", @Periodo ='" & Format(Item.Text, "MM/yyyy") & _
                                                      "', @CentroEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
      RsValidar.Open Sql, Conec
      
      If VecCentroCta(i).O_PrecioPactado > RsValidar!MontoDisponible Then
          MsgBox "No Existe un presupuesto aprobado para todas las Cuenta Contable/Centro de Costo" & _
                 " Para el Período " & Format(Item.Text, "MMMM/yyyy"), vbInformation, "Importe"
        Item.Checked = False
        Exit For
      End If
      RsValidar.Close
  Next
    Set RsValidar = Nothing
End Sub

Private Sub OptMulti_Click(Index As Integer)
    
    If Index = 0 Then
       Height = 6530
       
       FrameMulti.Enabled = False
       CmdConfirmar.Enabled = True
    Else
       Height = 8700
       CmbFrecuencia.ListIndex = 0
       Call TxtCantOrdenes_Change
       LvCenCostoCtas.ListItems.Clear
       LvCenCostoCtas.ListItems.Add
       ReDim VecCentroCta(0)
       FrameMulti.Enabled = True
       CmdConfirmar.Enabled = False
    End If
    Call CentrarFormulario(Me)
End Sub

Private Sub TxtCantOrdenes_Change()
    Dim i As Integer
    Dim cant As Integer
    cant = Val(TxtCantOrdenes.Text)
    
    LvMeses.ListItems.Clear
    
    Select Case CmbFrecuencia.ListIndex
    Case 0
        For i = 0 To cant - 1
            LvMeses.ListItems.Add , , Format(DateAdd("M", i, CalFecha.Value), "MMMM/YYYY")
        Next
    Case 1
        For i = 0 To cant - 1
            LvMeses.ListItems.Add , , Format(DateAdd("M", i * 2, CalFecha.Value), "MMMM/YYYY")
        Next
    Case 2
        For i = 0 To cant - 1
            LvMeses.ListItems.Add , , Format(DateAdd("M", i * 3, CalFecha.Value), "MMMM/YYYY")
        Next
    Case 3
        For i = 0 To cant - 1
            LvMeses.ListItems.Add , , Format(DateAdd("M", i * 4, CalFecha.Value), "MMMM/YYYY")
        Next
    Case 4
        For i = 0 To cant - 1
            LvMeses.ListItems.Add , , Format(DateAdd("M", i * 6, CalFecha.Value), "MMMM/YYYY")
        Next
    End Select
End Sub

Private Sub TxtCantOrdenes_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
    End If

End Sub

Private Sub TxtObs_Change()
    Modificado = True
End Sub

Private Sub TxtPrecioU_KeyPress(KeyAscii As Integer)
    Call TxtNumericoNeg(TxtPrecioU, KeyAscii)
End Sub

Private Sub Form_Load()
    CentroEmisorActual = CentroEmisor
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    Call CargarCmbFormasDePago(CmbFormaDePago)
    Call CargarComboEmpresas(CmbEmp, "Sin")
    Call CargarCmbCentrosDeCostos(CmbCentrosDeCostos)
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    Nivel = TraerNivel("A016100")
    TxtResp.Text = NombreUsuario
    TxtNroOrden.Text = ""
    CalFecha.Value = ValidarPeriodo(Date, False)
    Modificado = False
    TotalMontoSinPres = 0
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)

    With CmbFrecuencia
        .Clear
        .AddItem "Mensual"
        .AddItem "Bimestral"
        .AddItem "Trimestral"
        .AddItem "Cuatrimestral"
        .AddItem "Semestral"
        '.ListIndex = 0
    End With
    Call OptMulti_Click(0)

End Sub

Private Sub CrearEncabezados()
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos", (LvCenCostoCtas.Width - 1600) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Sub-Centros de Costos", (LvCenCostoCtas.Width - 1600) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", (LvCenCostoCtas.Width - 1600) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Importe sin IVA", 1300, 1
    
    LvMeses.ColumnHeaders.Add , , "Período", LvMeses.Width - 300
End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To UBound(VecCentroCta)
        Total = Total + VecCentroCta(i).O_PrecioPactado
    Next
        txtTotal.Text = Format(Total, "$ 0.00##")
    Call DesCheckingMeses
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
            If NroOrden = 0 Then
                If OptMulti(0).Value Then
                    If Not GrabarOrden Then
                       Cancel = 1
                       Exit Sub
                    End If
                Else
                    If Not CrearMultiOrdenes Then
                       Cancel = 1
                       Exit Sub
                    End If
                End If
            Else
                If Not ModificarOrden Then
                   Cancel = 1
                   Exit Sub
                End If

            End If
         End If
       End If
    End If
End Sub

Private Sub LvCenCostoCtas_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
'NO SE TOCA
   If Item.Index < LvCenCostoCtas.ListItems.Count Then
       Call CargarEnModificarCentorCta(Item.Index)
        CmdModifCentroCta.Enabled = True
        CmdEliminarCentroCta.Enabled = True
        CmdAgregarCentroCta.Enabled = False
    Else
        CmdAgregarCentroCta.Enabled = True
        CmdModifCentroCta.Enabled = False
        CmdEliminarCentroCta.Enabled = False
        
        CmbCentrosDeCostos.ListIndex = 0
        CmbCuentas.ListIndex = 0
        TxtPrecioU.Text = ""
   End If

End Sub

Private Sub CargarEnModificarCentorCta(Index As Integer)
    Dim i As Integer
    
    i = Index
    Call UbicarCuentaContable(VecCentroCta(i).O_CuentaContable, CmbCuentas)
    Call BuscarCentro(VecCentroCta(i).O_CentroDeCosto, CmbCentrosDeCostos)
    TxtPrecioU.Text = Replace(VecCentroCta(i).O_PrecioPactado, ",", ".")
End Sub

Private Sub Timer1_Timer()
   If NroOrden <> 0 Then
      TxtNroOrden.Text = CStr(NroOrden)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub LimpiarOrden()
   
    TxtLugar.Text = ""
    TxtResp.Text = NombreUsuario
    txtTotal = "0"
    NroOrden = 0
     
    LvCenCostoCtas.ListItems.Clear
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    
    ReDim VecCentroCta(0)

    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCentrosDeCostos.ListIndex = 0
    CmbCuentas.ListIndex = 0
    CmbEmp.ListIndex = 0
    CmbFormaDePago.ListIndex = 0
    TxtObs.Text = ""
    CmbProv.ListIndex = 0
    
    CalFecha.Value = ValidarPeriodo(Date, False)
    CalFecha.Format = dtpCustom
    CalFecha.CustomFormat = " "
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)
    CmdConfirmar.Visible = True
    CmdCambiar.Visible = False
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    Modificado = False
    
    'muestra los option button para hacer pultiples ordenes
    OptMulti(0).Visible = True
    OptMulti(1).Visible = True
    OptMulti(0).Value = True
    
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
End Sub

Private Sub TxtLugar_Change()
    Modificado = True
End Sub

Private Sub TxtNroOrden_KeyPress(KeyAscii As Integer)
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

Private Sub TxtNroOrden_LostFocus()
  If Val(TxtNroOrden.Text) <> NroOrden Then
    CmdConfirmar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
  End If
End Sub

Private Sub CargarOrden(NroOrden As Integer)
    Dim Sql As String
    Dim i As Integer
    Dim PeriodoCerrado As Boolean
    Dim Autorizado As Boolean
    Dim RsCargar As New ADODB.Recordset
    Dim RsValidarPeriodo As New ADODB.Recordset
'oculta los option button para hacer pultiples ordenes
    OptMulti(0).Visible = False
    OptMulti(1).Visible = False
    OptMulti(0).Value = True
    TotalMontoSinPres = 0
'On Error GoTo Error
    LBAnulada.Visible = False
    LBPerCerrado.Visible = False
    
  With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    Sql = "SpOCOrdenesDeContratacionCabeceraTraerNro @NroOrden= " + CStr(NroOrden) + _
            ", @Usuario='" + Usuario + "', @O_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
    .Open Sql, Conec
    
  If .EOF Then
      MsgBox "No existe una orden de Contratacion con esa numeración", vbInformation
      Call CmdNueva_Click
      Exit Sub
  Else
      CalFecha.Enabled = False
      CmbCentroDeCostoEmisor.Enabled = False

      Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(!O_Fecha, "MM/yyyy")) & "'"
      RsValidarPeriodo.Open Sql, Conec
      PeriodoCerrado = RsValidarPeriodo!Cerrado > 0

     If Not IsNull(!O_FechaAnulacion) Or PeriodoCerrado Then
        If Not IsNull(!O_FechaAnulacion) Then
            LBAnulada.Caption = "Anulada " + Mid(CStr(!O_FechaAnulacion), 1, 10)
            LBAnulada.Visible = True
        End If
        
        LBPerCerrado.Visible = PeriodoCerrado

        CmdAnular.Visible = False
        FrameAsig.Enabled = False
        CmdCambiar.Enabled = False
        TxtLugar.Enabled = False
        TxtObs.Enabled = False
        CmbEmp.Enabled = False
        CmbFormaDePago.Enabled = False
        CmbProv.Enabled = False
        'CalFecha.Enabled = False
    Else
        LBAnulada.Visible = False
        FrameAsig.Enabled = True
        CmdAnular.Visible = True
        CmdCambiar.Enabled = True
        TxtLugar.Enabled = True
        TxtObs.Enabled = True
        CmbEmp.Enabled = True
        CmbFormaDePago.Enabled = True
        CmbProv.Enabled = True
        'CalFecha.Enabled = True
    End If
    
    TxtNroOrden.Text = Format(!O_NumeroOrdenDeContratacion, "0000000000")
    Me.NroOrden = !O_NumeroOrdenDeContratacion
    Autorizado = VerificarNulo(!O_Autorizado, "B")


    CmdConfirmar.Visible = False
    CmdCambiar.Visible = True
    CmdImprimir.Enabled = True
    CmdExpPdf.Enabled = True
        
    CalFecha.Value = !O_Fecha
    CalFecha.CustomFormat = "MM/yyyy"
    
    If IsNull(!O_FechaEmision) Then
        CalFechaEmitida.Value = !O_Fecha
     Else
        CalFechaEmitida.Value = !O_FechaEmision
     End If

    TxtResp = !O_Responsable
    TxtObs.Text = VerificarNulo(!O_Observaciones)
    Call BuscarProveedor(!O_CodigoProveedor, CmbProv)
    TxtLugar.Text = !O_LugarDelServicio
    
    If ValN(!O_CodigoFormaDePago) = 0 Then
        CmbFormaDePago.Text = !O_FormaDePagoPactada
    Else
        Call UbicarCmbFormasDePago(ValN(!O_CodigoFormaDePago), CmbFormaDePago)
    End If
    
    Call UbicarEmpresa(!O_EmpresaFacturaANombreDe, CmbEmp)
    Call BuscarCentroEmisor(!O_CentroDeCostoEmisor, CmbCentroDeCostoEmisor)
    
    .Close
     
    Sql = "SpOCOrdenesDeContratacionRenglonesTraer @NroOrden=" & NroOrden & _
                                                ", @O_CentroDeCostoEmisor= '" & CentroEmisorActual & "'"
    .Open Sql, Conec
        ReDim VecCentroCta(.RecordCount)
    i = 1
    LvCenCostoCtas.ListItems.Clear
    While Not .EOF
        VecCentroCta(i).Centro_Descripcion = BuscarDescCentro(!O_CentroDeCosto)
        VecCentroCta(i).Cta_Descripcion = BuscarDescCta(!O_CuentaContable)
        VecCentroCta(i).O_CentroDeCosto = !O_CentroDeCosto
        VecCentroCta(i).O_CuentaContable = !O_CuentaContable
        VecCentroCta(i).O_PrecioPactado = !O_PrecioPactado
        VecCentroCta(i).O_PendienteAutorizacionDePago = !O_PendienteAutorizacionDePago
        VecCentroCta(i).CentroPadre = BuscarDescCentroEmisor(BuscarCentroPadre(!O_CentroDeCosto))
        VecCentroCta(i).O_SinPresupuestar = VerificarNulo(!O_SinPresupuestar, "B")
        VecCentroCta(i).O_MontoSinPresupuestar = VerificarNulo(!O_MontoSinPresupuestar, "N")
        
        LvCenCostoCtas.ListItems.Add
        LvCenCostoCtas.ListItems(i).Text = VecCentroCta(i).CentroPadre
        LvCenCostoCtas.ListItems(i).SubItems(1) = VecCentroCta(i).Centro_Descripcion
        LvCenCostoCtas.ListItems(i).SubItems(2) = VecCentroCta(i).Cta_Descripcion
        LvCenCostoCtas.ListItems(i).SubItems(3) = Format(VecCentroCta(i).O_PrecioPactado, "0.00")
        
        i = i + 1
        .MoveNext
    Wend
    LvCenCostoCtas.ListItems.Add
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)

  End If
  End With
  
  TotalMontoSinPres = CalcularTotalSinPresupuestar
  Call CalcularTotal
    
  If Not Autorizado Then
      MsgBox "La Orden no ha sido autorizada ahún", vbInformation
      CmdImprimir.Enabled = False
      CmdExpPdf.Enabled = False
  Else
      CmdImprimir.Enabled = True
      CmdExpPdf.Enabled = True
  End If

Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

'funciones que eran púplical
Private Sub CargarVecCentroEmisor(CentroEmisor As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
 'dependiendo del centro de costo emisor carga las cuentas correspondientes
 'en esta sección carga las cuentas
   
   Sql = "SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CentroEmisor & "'"
    With RsCargar
      .Open Sql, Conec, adOpenStatic, adLockReadOnly
      
      ReDim VecCuentasContables(0)
        
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

Public Sub UbicarCuentaContable(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
   
    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = C_Codigo Then
           Cmb.ListIndex = i
           Exit For
        End If
    Next
    
End Sub

Private Sub DesCheckingMeses()
    Dim i As Integer
    
    For i = 1 To LvMeses.ListItems.Count
        LvMeses.ListItems(i).Checked = False
    Next
End Sub

Private Function CalcularTotalSinPresupuestar() As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        Total = Total + VecCentroCta(i).O_MontoSinPresupuestar
    Next
       
    CalcularTotalSinPresupuestar = Total
End Function

Private Function CalcularTotalCuenta(Cuenta As String) As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        If VecCentroCta(i).O_CuentaContable = Cuenta And i <> LvCenCostoCtas.SelectedItem.Index Then
           Total = Total + VecCentroCta(i).O_PrecioPactado - VecCentroCta(i).O_MontoSinPresupuestar
        End If
    Next
       
    CalcularTotalCuenta = Total
End Function

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call UbicarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub CalFecha_GotFocus()
    CalFecha.CustomFormat = "MM/yyyy"
End Sub

