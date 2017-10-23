VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Begin VB.Form A01_2100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Orden de Compra"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   720
      Top             =   7740
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
      Left            =   1485
      Top             =   7785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5040
      TabIndex        =   58
      Top             =   7930
      Width           =   1230
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular"
      Height          =   330
      Left            =   3600
      TabIndex        =   57
      Top             =   7930
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdNueva 
      Caption         =   "&Nueva"
      Height          =   350
      Left            =   6435
      TabIndex        =   31
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   180
      Top             =   7875
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación"
      Height          =   2220
      Left            =   8235
      TabIndex        =   49
      Top             =   5220
      Width           =   3795
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   2970
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1035
         Width           =   645
      End
      Begin VB.CommandButton CmdBuscarSubCentro 
         Caption         =   "Buscar"
         Height          =   300
         Left            =   2925
         TabIndex        =   24
         Top             =   450
         Width           =   735
      End
      Begin VB.TextBox TxtCantCentroCta 
         Height          =   315
         Left            =   1035
         TabIndex        =   27
         Top             =   1380
         Width           =   870
      End
      Begin VB.CommandButton CmdAgregarCentroCta 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   1305
         TabIndex        =   29
         Top             =   1755
         Width           =   1150
      End
      Begin VB.CommandButton CmdEliminarCentroCta 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   30
         Top             =   1755
         Width           =   1150
      End
      Begin VB.CommandButton CmdModifCentroCta 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   90
         TabIndex        =   28
         Top             =   1755
         Width           =   1150
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   330
         Left            =   135
         TabIndex        =   25
         Top             =   1035
         Width           =   2760
         _ExtentX        =   4868
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
         Left            =   135
         TabIndex        =   23
         Top             =   450
         Width           =   2760
         _ExtentX        =   4868
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
         Left            =   2970
         TabIndex        =   61
         Top             =   810
         Width           =   405
      End
      Begin VB.Label LbCant 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad:"
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
         TabIndex        =   52
         Top             =   1440
         Width           =   825
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
         Left            =   135
         TabIndex        =   51
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
         Left            =   135
         TabIndex        =   50
         Top             =   810
         Width           =   1485
      End
   End
   Begin VB.Frame FraArt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Artículo"
      Height          =   2625
      Left            =   8910
      TabIndex        =   45
      Top             =   2250
      Width           =   3120
      Begin VB.TextBox TxtCodArticulo 
         Height          =   315
         Left            =   900
         TabIndex        =   14
         Top             =   160
         Width           =   1005
      End
      Begin VB.CommandButton CmdRequerimientos 
         Caption         =   "Requerimientos"
         Height          =   350
         Left            =   1620
         TabIndex        =   63
         Top             =   2160
         Width           =   1300
      End
      Begin VB.CommandButton CmdCrearNuevoArt 
         Caption         =   "Crear Nuevo"
         Height          =   315
         Left            =   1935
         TabIndex        =   15
         Top             =   160
         Width           =   1050
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   180
         TabIndex        =   19
         Top             =   1710
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   180
         TabIndex        =   21
         Top             =   2160
         Width           =   1300
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   1620
         TabIndex        =   20
         Top             =   1710
         Width           =   1300
      End
      Begin VB.TextBox TxtCant 
         Height          =   315
         Left            =   2025
         TabIndex        =   17
         Top             =   945
         Width           =   1005
      End
      Begin VB.TextBox TxtPrecioU 
         Height          =   315
         Left            =   2025
         TabIndex        =   18
         Top             =   1305
         Width           =   1005
      End
      Begin Controles.ComboEsp CmbArtCompra 
         Height          =   330
         Left            =   90
         TabIndex        =   16
         Top             =   540
         Width           =   2940
         _ExtentX        =   5186
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio Unit. sin IVA:"
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
         Left            =   225
         TabIndex        =   48
         Top             =   1365
         Width           =   1740
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad:"
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
         Left            =   1140
         TabIndex        =   47
         Top             =   1005
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Artículo:"
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
         TabIndex        =   46
         Top             =   220
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7875
      TabIndex        =   32
      Top             =   7920
      Width           =   1230
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
      Left            =   7650
      TabIndex        =   44
      Text            =   "0"
      Top             =   4845
      Width           =   1140
   End
   Begin MSComctlLib.ListView LvCenCostoCtas 
      Height          =   2580
      Left            =   45
      TabIndex        =   22
      Top             =   5220
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4551
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
      Left            =   10755
      TabIndex        =   35
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Orden de Compra"
      Height          =   2130
      Left            =   45
      TabIndex        =   36
      Top             =   45
      Width           =   11985
      Begin VB.TextBox TxtDescuento 
         Height          =   315
         Left            =   11115
         TabIndex        =   12
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox TxtCuit 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5085
         TabIndex        =   64
         Top             =   1395
         Width           =   1230
      End
      Begin Controles.ComboEsp CmbFormaDePago 
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   1050
         Width           =   4440
         _ExtentX        =   7832
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
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1755
         Width           =   8310
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   2610
         TabIndex        =   1
         Top             =   315
         Width           =   900
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   315
         Width           =   900
      End
      Begin VB.TextBox TxtResp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   5
         Top             =   720
         Width           =   4605
      End
      Begin VB.TextBox TxtNroOrden 
         Height          =   315
         Left            =   1305
         TabIndex        =   0
         Top             =   315
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   8280
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   106496003
         UpDown          =   -1  'True
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1080
         TabIndex        =   9
         Top             =   1395
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
         Left            =   8280
         TabIndex        =   10
         Top             =   1395
         Width           =   3615
         _ExtentX        =   6376
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
         Left            =   8280
         TabIndex        =   6
         Top             =   675
         Width           =   3615
         _ExtentX        =   6376
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
         Left            =   10620
         TabIndex        =   4
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   106496001
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbLugar 
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   1050
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descuento %:"
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
         Left            =   9855
         TabIndex        =   67
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "C.U.I.T.:"
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
         Left            =   4320
         TabIndex        =   65
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "F. Factura:"
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
         Left            =   9630
         TabIndex        =   62
         Top             =   315
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Imputación:"
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
         Left            =   6525
         TabIndex        =   37
         Top             =   315
         Width           =   1710
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
         TabIndex        =   59
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
         Left            =   4590
         TabIndex        =   56
         Top             =   405
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label LBAnulada 
         Alignment       =   2  'Center
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
         Left            =   4590
         TabIndex        =   55
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
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
         Left            =   6195
         TabIndex        =   54
         Top             =   720
         Width           =   2055
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
         Left            =   6720
         TabIndex        =   53
         Top             =   1080
         Width           =   1530
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   360
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
         TabIndex        =   40
         Top             =   765
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
         TabIndex        =   39
         Top             =   1110
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
         Left            =   6390
         TabIndex        =   38
         Top             =   1440
         Width           =   1860
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2535
      Left            =   45
      TabIndex        =   13
      Top             =   2250
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4471
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
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   9315
      TabIndex        =   34
      Top             =   7920
      Width           =   1300
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   9315
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label LbVariacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Variación Precio"
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
      Left            =   8910
      TabIndex        =   66
      Top             =   4935
      Width           =   1410
   End
   Begin VB.Label LbTotAsignado 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cantidad Total Asignada:"
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
      Left            =   8235
      TabIndex        =   60
      Top             =   7605
      Width           =   2160
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
      Left            =   7110
      TabIndex        =   43
      Top             =   4905
      Width           =   510
   End
End
Attribute VB_Name = "A01_2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VecOrdenDeCompra() As TipoOrdenDeCompra
'este vector no se aparea con el Lv
Private VecCentroCta() As TipoCentroCta

Private MontoSinPres As Double
Private TotalMontoSinPres As Double
Private SinPres As Boolean

Private Modificado As Boolean
Private A_Codigo As Long
Public NroOrden As Integer
Private Nivel As Integer
Private CantNoAsignada As Long
Private VecArtCompra() As TipoArticuloCompras
Private VecCuentasContables() As CuentasContables
Public CentroEmisorActual As String
Public TablaArticulos As String
Public TablaRequerimientos As String
Private RsRenglones As ADODB.Recordset
Private Requerimiento As Boolean

Public Sub CargarRequerimineto()
Dim i As Integer
    LvListado.ListItems.Clear
    ReDim VecOrdenDeCompra(UBound(VecRequerimientoCompra))
    For i = 1 To UBound(VecRequerimientoCompra)
        With VecOrdenDeCompra(i)
            .A_Codigo = VecRequerimientoCompra(i).CodArticulo
            .A_Descripcion = VecRequerimientoCompra(i).DescArticulo
            .Cantidad = VecRequerimientoCompra(i).Cantidad
            .CantPendiente = VecRequerimientoCompra(i).Cantidad
            .Requerimiento = True
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = Trim(.A_Descripcion)
            LvListado.ListItems(i).SubItems(1) = Format(.Cantidad, "0.00##")
            LvListado.ListItems(i).SubItems(2) = Format(.PrecioUnit, "0.00##")
            LvListado.ListItems(i).SubItems(3) = Format(.Cantidad * .PrecioUnit, "0.00##")
            
            LvListado.ListItems(i).ForeColor = vbRed
            LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
            LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
            LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
        End With
    Next
    
    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
End Sub

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalFecha_Change()
    CalFecha.Value = ValidarPeriodo(CalFecha.Value)
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    
    LvCenCostoCtas.ListItems.Clear
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True

    ReDim VecOrdenDeCompra(0)
    ReDim VecCentroCta(0)
    
    txtTotal.Text = "$ 0"

End Sub

Private Sub CalFecha_GotFocus()
    CalFecha.CustomFormat = "MM/yyyy"
End Sub



Private Sub CmbCentroDeCostoEmisor_Click()
    'Modificado = True
    TablaArticulos = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_TablaArticulos
    TablaRequerimientos = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_TablaRequerimientosDeCompra
    CmdRequerimientos.Visible = TablaRequerimientos <> ""
    TxtCodArticulo.Visible = TablaArticulos <> ""
    CmdCrearNuevoArt.Enabled = TablaArticulos = ""
    Call CargarVecCentroEmisor(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo)
    CentroEmisorActual = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    
    LvCenCostoCtas.ListItems.Clear

    ReDim VecOrdenDeCompra(0)
    ReDim VecCentroCta(0)
    
    txtTotal.Text = "$ 0"

End Sub

Private Sub CmbEmp_Click()
    Modificado = True
End Sub

Private Sub CmbFormaDePago_Click()
    Modificado = True
End Sub

Private Sub CmbLugar_Click()
    Modificado = True
End Sub

Private Sub CmbProv_Click()
    Modificado = True
    TxtCuit = Format(VecProveedores(CmbProv.ListIndex).CUIT, "##-########-#")

End Sub

Private Sub CmbProv_Validate(Cancel As Boolean)
    If VecProveedores(CmbProv.ListIndex).Calificacion < 40 And VecProveedores(CmbProv.ListIndex).Calificacion <> 0 Then
        MsgBox "El Proveedor no está Aprobado para realizar compras", vbInformation
        CmbProv.ListIndex = 0
        Cancel = True
    End If
End Sub

Private Sub CmdRequerimientos_Click()
On Error GoTo errorprecio
    ReDim Requerimientos(0)
    Dim i As Integer
    Dim R As Integer
    
    A01_2110.Show vbModal
    If UBound(Requerimientos) Then
        ReDim VecOrdenDeCompra(0)
        For i = 1 To UBound(Requerimientos)
            R = Existe(Requerimientos(i).CodArticulo)
            If R = 0 Then
               ReDim Preserve VecOrdenDeCompra(UBound(VecOrdenDeCompra) + 1)
               R = UBound(VecOrdenDeCompra)
            End If
            
            VecOrdenDeCompra(R).A_Codigo = Requerimientos(i).CodArticulo
            VecOrdenDeCompra(R).CantPendiente = Requerimientos(i).Cantidad
            VecOrdenDeCompra(R).Cantidad = VecOrdenDeCompra(R).Cantidad + Requerimientos(i).Cantidad + Requerimientos(i).CantidadExtra
            VecOrdenDeCompra(R).A_Descripcion = Requerimientos(i).DescArticulo
            VecOrdenDeCompra(R).Requerimiento = True
        Next
            A01_2120.LvListado.ListItems.Clear
            For i = 1 To UBound(VecOrdenDeCompra)
                A01_2120.LvListado.ListItems.Add
                A01_2120.LvListado.ListItems(i).Text = VecOrdenDeCompra(i).A_Descripcion
                A01_2120.LvListado.ListItems(i).SubItems(1) = VecOrdenDeCompra(i).Cantidad
                'LvListado.ListItems(i).SubItems(2) = VecOrdenDeCompra(i).PrecioUnit
            Next
         'si no hay artículos no pide los precios y sale de la función
        If UBound(VecOrdenDeCompra) = 0 Then
            Exit Sub
        End If
        ReDim Precios(0)
        A01_2120.Show vbModal
        If UBound(Precios) = UBound(VecOrdenDeCompra) Then
            TxtObs.Text = "Requerimiento del Taller"
            LvListado.ListItems.Clear
            Requerimiento = True
            For R = 1 To UBound(VecOrdenDeCompra)
                LvListado.ListItems.Add
                VecOrdenDeCompra(R).PrecioUnit = Precios(R) - (Precios(R) * Val(TxtDescuento.Text) / 100)
                LvListado.ListItems(R).Text = VecOrdenDeCompra(R).A_Descripcion
                LvListado.ListItems(R).SubItems(1) = Format(VecOrdenDeCompra(R).Cantidad, "0.00##")
                LvListado.ListItems(R).SubItems(2) = Format(VecOrdenDeCompra(R).PrecioUnit, "0.00##")
                LvListado.ListItems(R).SubItems(3) = Format(VecOrdenDeCompra(R).Cantidad * VecOrdenDeCompra(R).PrecioUnit, "0.00##")
                'pone en rojo indicando que se debe asignar a un centro y cuenta las
                'cantidades
                LvListado.ListItems(R).ForeColor = vbRed
                LvListado.ListItems(R).ListSubItems(1).ForeColor = vbRed
                LvListado.ListItems(R).ListSubItems(2).ForeColor = vbRed
                LvListado.ListItems(R).ListSubItems(3).ForeColor = vbRed
            Next
            LvListado.ListItems.Add
            CmdRequerimientos.Enabled = False
        Else
            ReDim VecOrdenDeCompra(0)
        End If
    End If
    
    Call CalcularTotal
    ReDim Precios(0)
errorprecio:
    If Err.Number > 0 Then
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub

Private Function Existe(Articulo As Long) As Integer
    Dim i As Integer
    Existe = 0
    For i = 1 To UBound(VecOrdenDeCompra)
        If VecOrdenDeCompra(i).A_Codigo = Articulo Then
            Existe = i
            Exit Function
        End If
    Next
End Function


Private Sub TxtCodArticulo_KeyPress(KeyAscii As Integer)
    'controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 _
       And KeyAscii <> Asc(".") Then
          Beep
          KeyAscii = 0
    End If
End Sub

Private Sub TxtCodArticulo_Validate(Cancel As Boolean)
    If TxtCodArticulo.Text <> "" Then
       Call BuscarArt(Val(TxtCodArticulo), CmbArtCompra)
    End If
End Sub

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call UbicarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub CmbCuentas_Click()
    If CmbCuentas.ListIndex > 0 Then
       TxtCodCuenta.Text = VecCuentasContables(CmbCuentas.ListIndex).Codigo
    End If
End Sub

Private Sub CmdAgregarCentroCta_Click()
On Error GoTo Errores
Dim i As Integer
  SinPres = False
  MontoSinPres = 0
  
 If ValidarCargaCentroCta Then
    Modificado = True
        
    If LvCenCostoCtas.SelectedItem.Index = LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Index Then
        
       'agrega al vector
        i = LvCenCostoCtas.SelectedItem.Index
        ReDim Preserve VecCentroCta(UBound(VecCentroCta) + 1)
        TotalMontoSinPres = TotalMontoSinPres + MontoSinPres
        With VecCentroCta(UBound(VecCentroCta))
            .O_CodigoArticulo = A_Codigo
            .Centro_Descripcion = Trim(CmbCentrosDeCostos.Text)
            .O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
            .O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
            .Cta_Descripcion = Trim(CmbCuentas.Text)
            .O_CantidadPedida = Val(TxtCantCentroCta.Text)
            .O_CantidadPendiente = Val(TxtCantCentroCta.Text)
            .O_SinPresupuestar = SinPres
            .O_MontoSinPresupuestar = MontoSinPres
            
            'lo pone en el LV
            LvCenCostoCtas.ListItems(i).Text = BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
            LvCenCostoCtas.ListItems(i).SubItems(1) = .Centro_Descripcion
            LvCenCostoCtas.ListItems(i).SubItems(2) = .Cta_Descripcion
            LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.O_CantidadPedida, "0.00##")
            LvCenCostoCtas.ListItems(i).SubItems(4) = UBound(VecCentroCta)
        End With

         'es el último registro, por lo tanto quería agregar uno nuevo
            LvCenCostoCtas.ListItems.Add
         'pocisiona en el último
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Selected = True
          'limpia los controles
            CmbCentrosDeCostos.ListIndex = 0
            CmbCuentas.ListIndex = 0
            TxtCantCentroCta.Text = ""
     End If
     Call ValidarCantArticulos(A_Codigo)
     Call CalcularTotalAsignado
     'le da el foco el combo de centros de costo
     CmbCentrosDeCostos.SetFocus
  End If
      TotalMontoSinPres = CalcularTotalSinPresupuestar

Errores:
    If Err.Number <> 0 Then
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub

Private Function ValidarCargaCentroCta() As Boolean
    ValidarCargaCentroCta = True
    Dim i As Integer
    Dim Index As Integer
    Dim Rta As Integer
    MontoSinPres = 0
    
    If VecOrdenDeCompra(LvListado.SelectedItem.Index).PrecioUnit = 0 Then
        MsgBox "Debe Ingresar Precio unitario", vbInformation
        TxtPrecioU.SetFocus
        ValidarCargaCentroCta = False
        Exit Function
   End If

    If VecCuentasContables(CmbCuentas.ListIndex).Codigo = "5121" Then
       MsgBox "La Cuenta " & CmbCuentas.Text & " no puede ser utilizada en órdenes de compra", vbInformation
       CmbCuentas.SetFocus
       ValidarCargaCentroCta = False
       Exit Function
    End If
    
    'recorro el lv para ver si no está la combinación de centro - cta
    For i = 1 To LvCenCostoCtas.ListItems.Count - 1
      If i <> LvCenCostoCtas.SelectedItem.Index Then
        Index = Val(LvCenCostoCtas.ListItems(i).SubItems(4))
        If VecCentroCta(Index).O_CodigoArticulo = A_Codigo And _
           VecCentroCta(Index).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo _
           And VecCentroCta(Index).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
           
             MsgBox "Para este artículo ya existe esta Combinación de Centro de Costo - Cuenta"
             ValidarCargaCentroCta = False
             Exit Function
        End If
      End If
    Next
    
    If CmbCentrosDeCostos.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Centro de Costo"
       CmbCentrosDeCostos.SetFocus
       ValidarCargaCentroCta = False
       Exit Function
    Else
        If CmbCuentas.ListIndex = 0 Then
            MsgBox "Debe Seleccionar una Cuenta"
            CmbCuentas.SetFocus
            ValidarCargaCentroCta = False
            Exit Function
        Else
           If Val(TxtCantCentroCta.Text) = 0 Then
              MsgBox "Debe Ingresar una cantidad mayor que 0"
              TxtCantCentroCta.SetFocus
              ValidarCargaCentroCta = False
              Exit Function
          End If
        End If
    End If

    Dim Sql As String
    Dim RsValidar As New ADODB.Recordset
    
    Sql = "SpOCPresupuestosRenglonesValidarCompra @CuentaContable = '" & VecCuentasContables(CmbCuentas.ListIndex).Codigo & _
                                              "', @NroOrden = " & NroOrden & _
                                               ", @CentroEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                                              "', @Periodo = '" & Format(CalFecha.Value, "MM/yyyy") & "'"
    RsValidar.Open Sql, Conec
        
    If Val(TxtCantCentroCta.Text) * VecOrdenDeCompra(LvListado.SelectedItem.Index).PrecioUnit + _
       TotalCtaContable(VecCuentasContables(CmbCuentas.ListIndex).Codigo) > RsValidar!MontoDisponible Then
    
        Rta = MsgBox("¿Desea imputar la cantidad Sin Presupuestar?", vbYesNo, "Fuera de Presupuesto")
        If Rta = vbYes Then
            SinPres = True
            
            MontoSinPres = Val(TxtCantCentroCta.Text) * VecOrdenDeCompra(LvListado.SelectedItem.Index).PrecioUnit _
                           - (RsValidar!MontoDisponible - (TotalCtaContable(VecCuentasContables(CmbCuentas.ListIndex).Codigo) - _
                              CalcularTotalSinPresupuestarPorCuenta(VecCuentasContables(CmbCuentas.ListIndex).Codigo)))

        Else
            TxtCantCentroCta.SetFocus
            ValidarCargaCentroCta = False
            Exit Function
        End If
       
        RsValidar.Close
        Sql = "SpOCImporteSinPresupuestarCompras @CentroDeCosto ='" & CentroEmisorActual & _
                                             "', @NroOrden =" & NroOrden & _
                                              ", @Periodo =" & FechaSQL(CalFecha.Value, "SQL")
        RsValidar.Open Sql, Conec
        
        If MontoSinPres + CalcularTotalSinPresupuestarPorCuenta(VecCuentasContables(CmbCuentas.ListIndex).Codigo) _
           > RsValidar!MontoSinPresupuestarMensual - RsValidar!MontoSinPres Then

            MsgBox "El Importe (precio x cantidad) aprobada para esa Cuenta Contable es de $" & RsValidar!MontoSinPresupuestarMensual - TotalMontoSinPres - RsValidar!MontoSinPres & _
                   " Para el Período " & Format(CalFecha.Value, "MMMM/yyyy"), vbInformation, "Importe"
            TxtCantCentroCta.SetFocus
            ValidarCargaCentroCta = False
            Exit Function
        End If
    End If

End Function

Private Function TotalCtaContable(CuentaContable As String) As Double
    Dim i As Integer
    Dim Total As Double
    
        For i = 1 To UBound(VecCentroCta)
            If i <> Val(LvCenCostoCtas.SelectedItem.SubItems(4)) And _
               VecCentroCta(i).O_CuentaContable = CuentaContable Then
                Total = Total + (VecCentroCta(i).O_CantidadPedida * BuscarPrecio(VecCentroCta(i).O_CodigoArticulo))
            End If
        Next
    TotalCtaContable = Total
End Function

Private Sub CmdAnular_Click()
  On Error GoTo ErrorAnulacion
    Dim Sql As String
    Dim Rta As Integer

    Dim RsAnular As ADODB.Recordset
    Set RsAnular = New ADODB.Recordset
       
    Rta = MsgBox("¿Confirma que desea anular la Orden de Compra?", vbYesNo)
    
    If Rta = vbYes Then
        
        Sql = "SpOCOrdenesDeCompraCabeceraAnular @O_NumeroOrdenDeCompra = " & NroOrden & _
                                              ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
        RsAnular.Open Sql, Conec
        'If RsAnular!OK = "OK" Then
            MsgBox RsAnular!Mensaje
        'End If
    End If
ErrorAnulacion:
 If Err.Number <> 0 Then
    Call ManipularError(Err.Number, Err.Description)
 End If

End Sub

Private Sub CMDBuscar_Click()
    NroOrden = 0
    Unload A01_2200
    A01_2200.LbCentroEmisor.Visible = Nivel = 2
    A01_2200.CmbCentroDeCostoEmisor.Visible = Nivel = 2
    A01_2200.Show vbModal, Me
    Timer1.Enabled = True
End Sub

Private Sub CmdBuscarSubCentro_Click()
    BuscarSubCentro.Show vbModal
    Call BuscarCentro(BuscarSubCentro.CodigoSubCentro, CmbCentrosDeCostos)
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar la Orden de Compra?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarOrden
    End If
End Sub

Private Sub ModificarOrden()
  Dim Sql As String
  Dim i As Integer
  Dim Precio As String
  Dim Rta As Integer
  
'On Error GoTo ErrorUpdate

  NroOrden = Val(TxtNroOrden.Text)
  If ValidarEncabezado Then
     If Not ValidarIntegridad Then
        Exit Sub
     End If
     
     Conec.BeginTrans
     Sql = "SpOCOrdenesDeCompraCabeceraModificar @O_NumeroOrdenDeCompra=" + CStr(NroOrden) + _
           ", @O_Fecha=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
           ", @O_Responsable ='" + TxtResp.Text + _
           "', @O_CodigoProveedor ='" + CStr(VecProveedores(CmbProv.ListIndex).Codigo) + _
           "', @O_LugarDeEntrega = '" + CmbLugar.Text + _
           "', @O_FormaDePagoPactada ='" + CmbFormaDePago.Text + _
           "', @O_EmpresaFacturaANombreDe ='" + VecEmpresas(CmbEmp.ListIndex).Codigo + _
           "', @U_Usuario = '" + Usuario + _
           "', @O_CentroDeCostoEmisor = '" + VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
           "', @O_Observaciones = '" & TxtObs.Text & _
           "', @O_FechaEmision =" & FechaSQL(CalFechaEmitida, "SQL") & _
           " , @O_CodigoLugarDeEntrega = " & VecLugaresDeEntrega(IIf(CmbLugar.ListIndex = -1, 0, CmbLugar.ListIndex)).L_Codigo & _
           " , @O_CodigoFormaDePago=" & VecFormasDePago(CmbFormaDePago.ListIndex).F_Codigo & _
           " , @O_Autorizado= " & IIf(ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion, 0, 1) & _
           " , @O_Descuento =" & Val(TxtDescuento.Text)
        
        Conec.Execute Sql
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
      
        Precio = Replace(CStr(BuscarPrecio(.O_CodigoArticulo)), ",", ".")
        
        Sql = "SpOCOrdenesDeCompraRenglonesAgregar @O_NumeroOrdenDeCompra = " & NroOrden & _
              ", @O_CodigoArticulo = " & .O_CodigoArticulo & _
              ", @O_CuentaContable = '" & .O_CuentaContable & _
              "', @O_CentroDeCosto = '" & .O_CentroDeCosto & _
              "', @O_CantidadPedida =" & Replace(.O_CantidadPedida, ",", ".") & _
              ", @O_CantidadPendiente = " & Replace(.O_CantidadPendiente, ",", ".") & _
              ", @O_PrecioPactado = " & Precio & _
              ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & _
              "', @O_SinPresupuestar = " & IIf(.O_SinPresupuestar, 1, 0) & _
              ", @O_MontoSinPresupuestar = " & Replace(.O_MontoSinPresupuestar, ",", ".")

      End With
        Conec.Execute Sql
    Next
    Conec.CommitTrans
ErrorUpdate:
 On Error Resume Next

    If Err.Number = 0 Then
       If ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion Then
            MsgBox "La Orden de compra se Grabó correctamente, pero necesita autorización", vbInformation, "Supera los $" & MaxSinAutorizacion
            Call EnviarMailAutorizacion(NroOrden)
       Else
            MsgBox "La Orden de compra se Grabó correctamente"
       End If

       Call EnviarMail(NroOrden)

       Modificado = False
       Call CmdCargar_Click
    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Sub

Private Sub CmdCargar_Click()
    
    Call CargarOrden(Val(TxtNroOrden))
    Modificado = False
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirnar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar la Orden de Compra?", vbYesNo)
    If Rta = vbYes Then
        Call GrabarOrden
    End If
End Sub

Private Sub GrabarOrden()
  Dim Sql As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer
  Dim Precio As String

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
     Sql = "SpOCOrdenesDeCompraCabeceraAgregar @O_Fecha=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
           " , @O_Responsable ='" + TxtResp.Text + _
           "', @O_CodigoProveedor ='" + CStr(VecProveedores(CmbProv.ListIndex).Codigo) + _
           "', @O_LugarDeEntrega = '" + CmbLugar.Text + _
           "', @O_FormaDePagoPactada ='" + CmbFormaDePago.Text + _
           "', @O_EmpresaFacturaANombreDe ='" + VecEmpresas(CmbEmp.ListIndex).Codigo + _
           "', @U_Usuario = '" + Usuario + _
           "', @O_CentroDeCostoEmisor= '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
           "', @O_Observaciones= '" & TxtObs.Text & _
           "', @O_FechaEmision =" & FechaSQL(CalFechaEmitida, "SQL") & _
           " , @O_CodigoLugarDeEntrega = " & VecLugaresDeEntrega(IIf(CmbLugar.ListIndex = -1, 0, CmbLugar.ListIndex)).L_Codigo & _
           " , @O_CodigoFormaDePago=" & VecFormasDePago(CmbFormaDePago.ListIndex).F_Codigo & _
           " , @O_Autorizado= " & IIf(ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion, 0, 1) & _
           " , @O_Descuento =" & Val(TxtDescuento.Text)
           
     'graba el encabezado y retorna el Nro de Orden
        RsGrabar.Open Sql, Conec
        NroOrden = RsGrabar!O_NumeroOrdenDeCompra
        
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
      
        Precio = Replace(CStr(BuscarPrecio(.O_CodigoArticulo)), ",", ".")
        
        Sql = "SpOCOrdenesDeCompraRenglonesAgregar @O_NumeroOrdenDeCompra = " & NroOrden & _
                                                ", @O_CodigoArticulo = " & .O_CodigoArticulo & _
                                                ", @O_CuentaContable = '" & .O_CuentaContable & _
                                               "', @O_CentroDeCosto = '" & .O_CentroDeCosto & _
                                               "', @O_CantidadPedida =" & Replace(.O_CantidadPedida, ",", ".") & _
                                                ", @O_CantidadPendiente = " & Replace(.O_CantidadPedida, ",", ".") & _
                                                ", @O_PrecioPactado = " & Precio & _
                                                ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & _
                                                "', @O_SinPresupuestar = " & IIf(.O_SinPresupuestar, 1, 0) & _
                                                ", @O_MontoSinPresupuestar = " & Replace(.O_MontoSinPresupuestar, ",", ".")
      End With
        Conec.Execute Sql
    Next
        For i = 1 To UBound(Requerimientos)
            With Requerimientos(i)
                Sql = "SpOcOrdenesDeCompraRequerimientosRestarCantidad @Cantidad =" & .Cantidad & _
                                                                   " , @Articulo =" & .CodArticulo & _
                                                                   " , @Taller =" & .Taller & _
                                                                   " , @Marca ='" & .Marca & _
                                                                   "', @Numero =" & .Numero
                Conec.Execute Sql
            End With
        Next
        
        For i = 1 To UBound(VecRequerimientoCompra)
            With VecRequerimientoCompra(i)
                Sql = "SpOcRequerimientosCompradosAgregar @R_Cantidad =" & .Cantidad & _
                                                      " , @R_Articulo =" & .CodArticulo & _
                                                      " , @R_CentroDeCostoEmisor ='" & CentroEmisorActual & _
                                                      "', @R_NumeroOc =" & NroOrden & _
                                                      " , @R_Numero =" & .Numero
                Conec.Execute Sql
            End With
        Next

    Conec.CommitTrans
    
ErrorInsert:
 On Error Resume Next

    If Err.Number = 0 Then
       CmdConfirnar.Visible = False
       CmdCambiar.Visible = True
       CmdImprimir.Enabled = True
       CmdExpPdf.Enabled = True
       
       Call EnviarMail(NroOrden)
       If ValN(Replace(txtTotal, "$", "")) > MaxSinAutorizacion Then
            MsgBox "La orden necesita autorización", vbInformation, "Supera los $" & MaxSinAutorizacion
            FrmMensaje.CmdImprimir.Enabled = False
            FrmMensaje.CmdExportar.Enabled = False
            Call EnviarMailAutorizacion(NroOrden)
       End If
       
       FrmMensaje.LbMensaje = "La Orden de compra se Grabó correctamente con el Nº: " + CStr(NroOrden) & _
                              Chr(13) & " ¿Que desea hacer?"
       FrmMensaje.Show vbModal
       
       Modificado = False
       If FrmMensaje.Retorno = vbimprimir Then
         Call ConfImpresionDeOrden
         RepOrdenDeCompra.Show vbModal
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
End Sub

Private Function BuscarPrecio(A_Codigo As Long) As Double
  Dim i As Integer
    i = 1
    While VecOrdenDeCompra(i).A_Codigo <> A_Codigo
        i = i + 1
    Wend
    BuscarPrecio = VecOrdenDeCompra(i).PrecioUnit
End Function

Private Function ValidarEncabezado() As Boolean
Dim i As Integer
Dim Asignado As Boolean

    ValidarEncabezado = True
    If CalFecha.CustomFormat = " " Then
        MsgBox "Debe Ingresar Período de Imputación"
        CalFecha.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
    
    If TxtResp.Text = "" Then
        MsgBox "Debe ingresar el responsable"
        TxtResp.SetFocus
        ValidarEncabezado = False
        Exit Function
    Else
       If CmbCentroDeCostoEmisor.ListIndex = 0 Then
          MsgBox "Debe Seleccionar un Centro de Costo Emisor"
          CmbCentroDeCostoEmisor.SetFocus
          ValidarEncabezado = False
          Exit Function
       Else
        If CmbFormaDePago.ListIndex <= 0 Then
           MsgBox "Debe Seleccionar una Forma de Pago"
           CmbFormaDePago.SetFocus
           ValidarEncabezado = False
           Exit Function
        Else
          If CmbLugar.ListIndex = 0 Then
             MsgBox "Debe Seleccionar el Lugar de Entrega"
             CmbLugar.SetFocus
             ValidarEncabezado = False
             Exit Function
          Else
            If CmbProv.ListIndex = 0 Then
               MsgBox "Debe Seleccionar un Proveedor"
               CmbProv.SetFocus
               ValidarEncabezado = False
               Exit Function
            Else
                'If CmbEmp.ListIndex = 0 Then
                '   MsgBox "Debe Seleccionar a nombre de quien vendrá la Factura"
                '   CmbEmp.SetFocus
                '   ValidarEncabezado = False
                '   Exit Function
                'Else
                    If LvListado.ListItems.Count <= 1 Then
                       MsgBox "Debe ingresar artículos a la Órden de Compra"
                       LvListado.SetFocus
                       ValidarEncabezado = False
                       Exit Function
                    Else
                       Asignado = True
                        For i = 1 To LvListado.ListItems.Count - 1
                         'si alguna fina está en rojo es por que
                         'no estan asignados todos los artículos a
                         'un centro - cta
                            If LvListado.ListItems(i).ForeColor = vbRed Then
                               Asignado = False
                               Exit For
                            End If
                        Next
                        If Not Asignado Then
                           MsgBox "Todos los artículos deben estar asignados a Centros de costos / Cuentas Contables"
                           LvListado.SetFocus
                           ValidarEncabezado = False
                           Exit Function
                        End If
                    End If
                End If
              End If
            End If
        End If
      End If
    'End If
    
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
        
End Function

Private Sub CmdEliminar_Click()
    Dim IndexBorrar As Integer
    
    IndexBorrar = LvListado.SelectedItem.Index
    
    'borra del LV
 If VecOrdenDeCompra(IndexBorrar).Cantidad = VecOrdenDeCompra(IndexBorrar).CantPendiente Then
      LvListado.ListItems.Remove (IndexBorrar)
      'borrar del vector haciento un corrimiento
      While IndexBorrar < UBound(VecOrdenDeCompra)
          VecOrdenDeCompra(IndexBorrar) = VecOrdenDeCompra(IndexBorrar + 1)
          IndexBorrar = IndexBorrar + 1
      Wend
      
      Call EleminarArticuloDeVecCentroCta(A_Codigo)
      LvCenCostoCtas.ListItems.Clear
      Call HabilitarAsignacion(False)

      ReDim Preserve VecOrdenDeCompra(UBound(VecOrdenDeCompra) - 1)
         
    'calcula el total de la orden
      Call CalcularTotal
          
      Modificado = True
      
      If LvListado.ListItems.Count = LvListado.SelectedItem.Index Then
          CmdEliminar.Enabled = False
          CmdModif.Enabled = False
      End If
      
      TotalMontoSinPres = CalcularTotalSinPresupuestar

  Else
      MsgBox "El renglón no puede ser borrado por Haber artículos recibidos", , "Borrar"
  End If
End Sub

Private Sub EleminarArticuloDeVecCentroCta(A_Codigo As Long)
    Dim i As Integer
    Dim VecTemp() As TipoCentroCta
    
    i = 1
    ReDim VecTemp(0)
 'paso todo a un temporal y luego lo pongo en el original
    While i <= UBound(VecCentroCta)
      If VecCentroCta(i).O_CodigoArticulo <> A_Codigo Then
         ReDim Preserve VecTemp(UBound(VecTemp) + 1)
         VecTemp(UBound(VecTemp)) = VecCentroCta(i)
      End If
      i = i + 1
    Wend
    
    ReDim VecCentroCta(UBound(VecTemp))
    i = 1
    While i <= UBound(VecTemp)
        VecCentroCta(i) = VecTemp(i)
        i = i + 1
    Wend
    LvCenCostoCtas.ListItems.Clear
End Sub

Private Sub CmdEliminarCentroCta_Click()
    Dim IndexBorrar As Integer
    Dim IndexLv As Integer
    'guarda el índice del vector
    IndexBorrar = LvCenCostoCtas.SelectedItem.SubItems(4)
    
    IndexLv = LvCenCostoCtas.SelectedItem.Index
    
    
 If VecCentroCta(IndexBorrar).O_CantidadPedida = VecCentroCta(IndexBorrar).O_CantidadPendiente Then
    'borra del LV
     LvCenCostoCtas.ListItems.Remove (IndexLv)
       
    'borrar del vector
    While IndexBorrar < UBound(VecCentroCta)
        VecCentroCta(IndexBorrar) = VecCentroCta(IndexBorrar + 1)
        IndexBorrar = IndexBorrar + 1
    Wend
        
    ReDim Preserve VecCentroCta(UBound(VecCentroCta) - 1)
    Call CargarLvCenCostoCtas(A_Codigo)
      'reposiciona el lv
    LvCenCostoCtas.ListItems(IndexLv).Selected = True
      
    If LvCenCostoCtas.ListItems.Count = LvCenCostoCtas.SelectedItem.Index Then
        CmdEliminarCentroCta.Enabled = False
        CmdModifCentroCta.Enabled = False
    End If
        
    Modificado = True
    Call ValidarCantArticulos(A_Codigo)
    Call CalcularTotalAsignado
    TotalMontoSinPres = CalcularTotalSinPresupuestar

 Else
    MsgBox "El renglón no puede ser borrado por Haber artículos recibidos", , "Borrar"
 End If

End Sub

Private Sub ConfImpresionDeOrden()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "Articulo", adVarChar, 120
    RsListado.Fields.Append "Cantidad", adDouble
    RsListado.Fields.Append "Precio", adDouble
    RsListado.Fields.Append "Importe", adDouble
    RsListado.Open
    i = 1
    While i < LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
        RsListado!Articulo = .Text
        RsListado!Cantidad = ValN(.SubItems(1))
        RsListado!Precio = Val(Replace(.SubItems(2), ",", "."))
        RsListado!Importe = Val(Replace(.SubItems(3), ",", "."))
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    TxtNroOrden.Text = Format(NroOrden, "0000000000")
    RepOrdenDeCompra.TxtObservaciones = TxtObs
    RepOrdenDeCompra.TxtDescuento = TxtDescuento.Text
    RepOrdenDeCompra.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    RepOrdenDeCompra.TxtFormaDePago.Text = CmbFormaDePago.Text
    RepOrdenDeCompra.TxtFecha = Format(CalFecha.Value, "MM/yyyy")
    
    RepOrdenDeCompra.TxtFactNombre.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeCompra.LbCUIT.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeCompra.TxtCuit.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeCompra.LbFactura.Visible = CmbEmp.ListIndex > 0
    RepOrdenDeCompra.TxtCuit.Text = VecEmpresas(CmbEmp.ListIndex).CUIT
    RepOrdenDeCompra.TxtFactNombre.Text = CmbEmp.Text
    
    RepOrdenDeCompra.TxtFEmitida = CalFechaEmitida
    RepOrdenDeCompra.TxtNroOrden.Text = TxtNroOrden.Text
    RepOrdenDeCompra.TxtLugarDeEntrega = CmbLugar.Text
    RepOrdenDeCompra.TxtProv.Text = CmbProv.Text
    RepOrdenDeCompra.TxtCuitProv = TxtCuit.Text
    RepOrdenDeCompra.TxtResp.Text = TxtResp.Text
    RepOrdenDeCompra.TxtAnulada.Visible = LBAnulada.Visible
    RepOrdenDeCompra.TxtAnulada.Text = LBAnulada.Caption
    RepOrdenDeCompra.DataControl1.Recordset = RsListado
    
    RepOrdenDeCompra.Zoom = -1
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeOrden
    RepOrdenDeCompra.Show
End Sub

Private Sub CmdModif_Click()
On Error GoTo Errores
Dim i As Integer
    i = LvListado.SelectedItem.Index
    
 If ValidarCargaOC Then

   If VecOrdenDeCompra(i).CantPendiente + (Val(TxtCant.Text) _
       - VecOrdenDeCompra(i).Cantidad) >= 0 Then
        Modificado = True
        Call EleminarArticuloDeVecCentroCta(A_Codigo)
        A_Codigo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
       'agrega al vector
        VecOrdenDeCompra(i).A_Codigo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
        VecOrdenDeCompra(i).A_Descripcion = CmbArtCompra.Text
        VecOrdenDeCompra(i).CantPendiente = VecOrdenDeCompra(i).CantPendiente + (Val(TxtCant.Text) - VecOrdenDeCompra(i).Cantidad)
        VecOrdenDeCompra(i).MontoSinPres = MontoSinPres
        'TotalMontoSinPres = TotalMontoSinPres + MontoSinPres
        MontoSinPres = 0
        
        VecOrdenDeCompra(i).Cantidad = Val(TxtCant.Text)
        VecOrdenDeCompra(i).PrecioUnit = Val(TxtPrecioU.Text)
        'lo pone en el LV
        With VecOrdenDeCompra(i)
           LvListado.ListItems(i).Text = Trim(.A_Descripcion)
           LvListado.ListItems(i).SubItems(1) = Format(.Cantidad, "0.00##")
           LvListado.ListItems(i).SubItems(2) = Format(.PrecioUnit, "0.00##")
           LvListado.ListItems(i).SubItems(3) = Format(.Cantidad * .PrecioUnit, "0.00##")
        End With
        Call HabilitarAsignacion(False)

    Else
        MsgBox "La cantidad pedida de artículos debe ser mayor a la entregad para permitir su modificación"
    End If
 End If
 ' End If
  'calcula el total de la orden
  Call CalcularTotal
  TotalMontoSinPres = CalcularTotalSinPresupuestar
  Call ValidarCantArticulos(A_Codigo)

Errores:
  Call ManipularError(Err.Number, Err.Description, Timer1)

End Sub

Private Sub CmdModifCentroCta_Click()
On Error GoTo Errores
  Dim i As Integer
  Dim PosVec As Integer
  i = LvCenCostoCtas.SelectedItem.Index
  SinPres = False
  PosVec = Val(LvCenCostoCtas.ListItems(i).SubItems(4))

  TotalMontoSinPres = TotalMontoSinPres - VecCentroCta(PosVec).O_MontoSinPresupuestar
 If ValidarCargaCentroCta Then
     
       'agrega al vector
   If VecCentroCta(PosVec).O_CantidadPendiente + Val(TxtCantCentroCta.Text) _
        - VecCentroCta(PosVec).O_CantidadPedida >= 0 Then
        
        If VecCentroCta(PosVec).O_CantidadPendiente <> VecCentroCta(PosVec).O_CantidadPedida And _
           VecCentroCta(PosVec).O_CentroDeCosto <> VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo Then

            MsgBox "No Se puede modificar El Sub-Centro de Costo por tener Artículos recibidos", vbInformation
            Exit Sub
        End If
        
        If VecCentroCta(PosVec).O_CantidadPendiente <> VecCentroCta(PosVec).O_CantidadPedida And _
           VecCentroCta(PosVec).O_CuentaContable <> VecCuentasContables(CmbCuentas.ListIndex).Codigo Then

            MsgBox "No Se puede modificar la Cuenta Contable por tener Artículos recibidos", vbInformation
            Exit Sub
        End If
       
        Modificado = True
        VecCentroCta(PosVec).O_CodigoArticulo = A_Codigo
        VecCentroCta(PosVec).Centro_Descripcion = Trim(CmbCentrosDeCostos.Text)
        VecCentroCta(PosVec).O_CentroDeCosto = VecCentroDeCosto(CmbCentrosDeCostos.ListIndex).C_Codigo
        VecCentroCta(PosVec).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
        VecCentroCta(PosVec).Cta_Descripcion = Trim(CmbCuentas.Text)
        VecCentroCta(PosVec).O_CantidadPendiente = VecCentroCta(PosVec).O_CantidadPendiente + Val(TxtCantCentroCta.Text) - VecCentroCta(PosVec).O_CantidadPedida
        VecCentroCta(PosVec).O_CantidadPedida = Val(TxtCantCentroCta.Text)
        VecCentroCta(PosVec).O_SinPresupuestar = SinPres
        VecCentroCta(PosVec).O_MontoSinPresupuestar = MontoSinPres
        'lo pone en el LV
        With VecCentroCta(PosVec)
           LvCenCostoCtas.ListItems(i).Text = BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
           LvCenCostoCtas.ListItems(i).SubItems(1) = .Centro_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(2) = .Cta_Descripcion
           LvCenCostoCtas.ListItems(i).SubItems(3) = Format(.O_CantidadPedida, "0.00##")
           'LvCenCostoCtas.ListItems(i).SubItems(3) = UBound(VecCentroCta)
        End With
           Call ValidarCantArticulos(A_Codigo)
           Call CalcularTotalAsignado
    Else
        MsgBox "La cantidad pedida de artículos debe ser mayor a la entregada para permitir su modificación"
    End If
        
  End If
    TotalMontoSinPres = CalcularTotalSinPresupuestar
Errores:
  Call ManipularError(Err.Number, Err.Description, Timer1)

End Sub

Private Sub CmdNueva_Click()
    TxtNroOrden.Text = ""
    
    TotalMontoSinPres = 0
    MontoSinPres = 0
    
    CmdConfirnar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
    'esto es por que cuando se trae una orden anulada los
    'valores están invertidos
     LBAnulada.Visible = False
     FraArt.Enabled = True
     FrameAsig.Enabled = True
     'CmdCambiar.Enabled = True
     CmbLugar.Enabled = True
    
    'busca y establece el centro de costo emisor del usuario actual del sistema
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
     CmbEmp.Enabled = True
     CmbFormaDePago.Enabled = True
     CmbProv.Enabled = True
     CalFecha.Enabled = True
     Modificado = False
     Requerimiento = False
End Sub

Private Sub CmdCrearNuevoArt_Click()
    FrmCrearArticulo.Show vbModal
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
 'dependiendo del centro de costo emisor carga las cuentas y artículos correspondientes
    Sql = "SpOCRelacionCentroDeCostoArticulosTraer @R_CentroDeCosto='" & CentroEmisor & "'"
    With RsCargar
      ReDim VecArtCompra(0)
        If TablaArticulos = "" Then
             .Open Sql, Conec, adOpenStatic, adLockReadOnly
            'en esta sección carga los art
              For i = 1 To UBound(VariablesYFunciones.VecArtCompra)
                  .Find "R_Articulo = " & VariablesYFunciones.VecArtCompra(i).A_Codigo, , , 1
                 If Not .EOF Then
                    ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                    VecArtCompra(UBound(VecArtCompra)) = VariablesYFunciones.VecArtCompra(i)
                 End If
              Next
            .Close
         Else
            For i = 1 To UBound(VecArtTaller)
                ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                VecArtCompra(UBound(VecArtCompra)) = VecArtTaller(i)
            Next
         End If
    End With
    'carga los combos con los valores de los vectores locales
    Call CargarCmbArtCompra(CmbArtCompra)

End Sub

Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtDescuento, KeyAscii)
End Sub

Private Sub TxtPrecioU_KeyPress(KeyAscii As Integer)
    Call TxtNumericoNeg(TxtPrecioU, KeyAscii)
End Sub

Private Sub Form_Load()
    CentroEmisorActual = CentroEmisor
    TotalMontoSinPres = 0
    MontoSinPres = 0
    ReDim Requerimientos(0)
    Call CrearEncabezados
    Call CargarComboProveedores(CmbProv)
    Call CargarCmbFormasDePago(CmbFormaDePago)
    Call CargarComboEmpresas(CmbEmp, "Sin")
    Call CargarCmbLugaresDeEntrega(CmbLugar)
    Call CargarCmbCentrosDeCostos(CmbCentrosDeCostos)
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    Call HabilitarAsignacion(False)
    Nivel = TraerNivel("A012100")
    TxtNroOrden.Text = ""
    txtTotal.Text = Format(0, "$ 0.00##")
    CalFecha.Value = ValidarPeriodo(Date, False)
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)
    
    TxtResp.Text = NombreUsuario
    Modificado = False
End Sub

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeOrden
         RepOrdenDeCompra.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepOrdenDeCompra.Pages
         Unload RepOrdenDeCompra
         MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
  End If
Error:
     Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CalcularTotalAsignado()
    Dim i As Integer
    Dim cant As Double
    
   For i = 1 To UBound(VecCentroCta)
      If VecCentroCta(i).O_CodigoArticulo = A_Codigo Then
         cant = cant + VecCentroCta(i).O_CantidadPedida
      End If
   Next

   LbTotAsignado.Caption = "Cantidad Total Asignada: " & cant
End Sub

Private Sub ValidarCantArticulos(A_Codigo As Long)
   Dim i As Integer
   Dim cant As Double
   
   For i = 1 To UBound(VecCentroCta)
      If VecCentroCta(i).O_CodigoArticulo = A_Codigo Then
         cant = cant + VecCentroCta(i).O_CantidadPedida
      End If
   Next
   
   i = IIf(LvListado.SelectedItem.Index < LvListado.ListItems.Count, LvListado.SelectedItem.Index, LvListado.SelectedItem.Index - 1)
   
   If VecOrdenDeCompra(i).Cantidad = cant Then
     LvListado.ListItems(i).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(1).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(2).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(3).ForeColor = vbBlack
   Else
     LvListado.ListItems(i).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
   End If
   
End Sub

Private Sub CmdAgregar_Click()
'On Error GoTo errores
Dim i As Integer
    
 If ValidarCargaOC Then
    Modificado = True
        
    If LvListado.SelectedItem.Index = LvListado.ListItems(LvListado.ListItems.Count).Index Then
        
       'agrega al vector
        i = LvListado.SelectedItem.Index
        ReDim Preserve VecOrdenDeCompra(UBound(VecOrdenDeCompra) + 1)
        
        VecOrdenDeCompra(i).A_Codigo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
        VecOrdenDeCompra(i).A_Descripcion = CmbArtCompra.Text
        VecOrdenDeCompra(i).Cantidad = Val(TxtCant.Text)
        VecOrdenDeCompra(i).CantPendiente = Val(TxtCant.Text)
        VecOrdenDeCompra(i).PrecioUnit = Val(TxtPrecioU.Text)
        VecOrdenDeCompra(i).MontoSinPres = MontoSinPres
        
       ' TotalMontoSinPres = TotalMontoSinPres + MontoSinPres
        MontoSinPres = 0
        
        A_Codigo = VecOrdenDeCompra(i).A_Codigo
        'lo pone en el LV
        With VecOrdenDeCompra(i)
           LvListado.ListItems(i).Text = Trim(.A_Descripcion)
           LvListado.ListItems(i).SubItems(1) = Format(.Cantidad, "0.00##")
           LvListado.ListItems(i).SubItems(2) = Format(.PrecioUnit, "0.00##")
           LvListado.ListItems(i).SubItems(3) = Format(.Cantidad * .PrecioUnit, "0.00##")
           
        End With
       'pone en rojo indicando que se debe asignar a un centro y cuenta las
       'cantidades
           LvListado.ListItems(i).ForeColor = vbRed
           LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
           LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
           LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
           
    'es el último registro, por lo tanto quería agregar uno nuevo
            LvListado.ListItems.Add
         'pocisiona en el último
            LvListado.ListItems(LvListado.ListItems.Count).Selected = True
            Call LimpiarOC
     End If
     'le da el foco al combo de artículoa
      CmbArtCompra.SetFocus
  End If
  'calcula el total de la orden
   Call CalcularTotal
   Call HabilitarAsignacion(False)
   TotalMontoSinPres = CalcularTotalSinPresupuestar

Errores:
  Call ManipularError(Err.Number, Err.Description)

End Sub

Private Function ValidarCargaOC() As Boolean
On Error GoTo Error
    ValidarCargaOC = True
    Dim i As Integer
    Dim Rta As Integer
    
    If CmbCentroDeCostoEmisor.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Centro de Costo Emisor"
       CmbCentroDeCostoEmisor.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
    
    For i = 1 To UBound(VecOrdenDeCompra)
      If i <> LvListado.SelectedItem.Index Then
       If VecOrdenDeCompra(i).A_Codigo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo Then
          MsgBox "Ese artículo ya existe en esta Orden de compra"
          ValidarCargaOC = False
          Exit Function
       End If
      End If
    Next
    
    If CmbArtCompra.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Artículo"
       CmbArtCompra.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
       
    If Val(TxtCant.Text) = 0 Then
       MsgBox "Debe Ingresar una cantidad mayor que 0"
       TxtCant.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
           
    If Val(TxtPrecioU.Text) = 0 Then
       MsgBox "Debe ingresar un precio mayor que 0"
       TxtPrecioU.SetFocus
       ValidarCargaOC = False
        Exit Function
    End If
    
    Exit Function
Error:
    ValidarCargaOC = False
    Call ManipularError(Err.Number, Err.Description)
End Function

Private Sub CrearEncabezados()
    LvListado.ColumnHeaders.Add , , "Descripción Artículo", LvListado.Width - 3550
    LvListado.ColumnHeaders.Add , , "Cant.", 750, 1
    LvListado.ColumnHeaders.Add , , "P. Unitario sin IVA", 1450, 1
    LvListado.ColumnHeaders.Add , , "Importe", 1000, 1
    
    LvCenCostoCtas.ColumnHeaders.Add , , "Centros de Costos", (LvCenCostoCtas.Width - 1300) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Sub-Centros de Costos", (LvCenCostoCtas.Width - 1300) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Cuenta Contable", (LvCenCostoCtas.Width - 1300) / 3
    LvCenCostoCtas.ColumnHeaders.Add , , "Cantidad", 1000, 1
    LvCenCostoCtas.ColumnHeaders.Add , , "Index vector", 0

End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To LvListado.ListItems.Count
        Total = Total + Val(Replace(LvListado.ListItems(i).SubItems(3), ",", "."))
    Next
        txtTotal.Text = Format(Total, "$ 0.00##")

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
            Call ModificarOrden
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
        Dim CuentaPorDefecto As String
        CuentaPorDefecto = BuscarCtaPorDefectoArt(VecOrdenDeCompra(LvListado.SelectedItem.Index).A_Codigo)
        
        CmdAgregarCentroCta.Enabled = True
        CmdModifCentroCta.Enabled = False
        CmdEliminarCentroCta.Enabled = False
        
        CmbCentrosDeCostos.ListIndex = 0
        'CmbCuentas.ListIndex = 0
        Call UbicarCuentaContable(CuentaPorDefecto, CmbCuentas)
        TxtCantCentroCta.Text = ""
        'Call HabilitarAsignacion(False)
   End If

End Sub

Private Sub CargarEnModificarCentorCta(Index As Integer)
    Dim i As Integer
    
    i = Val(LvCenCostoCtas.ListItems(Index).SubItems(4))
    Call UbicarCuentaContable(VecCentroCta(i).O_CuentaContable, CmbCuentas)
    Call BuscarCentro(VecCentroCta(i).O_CentroDeCosto, CmbCentrosDeCostos)
    TxtCantCentroCta.Text = Replace(VecCentroCta(i).O_CantidadPedida, ",", ".")
End Sub

Private Sub Timer1_Timer()
   If NroOrden <> 0 Then
      TxtNroOrden.Text = CStr(NroOrden)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtCant, KeyAscii)
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
'NO SE TOCA
   If Item.Index < LvListado.ListItems.Count Then
        Call CargarEnModificar(Item.Index)
        Call CargarLvCenCostoCtas(A_Codigo)
        CmdModif.Enabled = True
        CmdEliminar.Enabled = True
        CmdAgregar.Enabled = False
    Else
        LvCenCostoCtas.ListItems.Clear
        LbTotAsignado.Caption = ""
        Call HabilitarAsignacion(False)
        Call LimpiarOC
        CmdAgregar.Enabled = True
        CmdModif.Enabled = False
        CmdEliminar.Enabled = False
        TxtCant.Enabled = True
        CmbArtCompra.Enabled = True
   End If
'errores:
 '   ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub HabilitarAsignacion(Habilitar As Boolean)
    LbCant.Enabled = Habilitar
    LBCC.Enabled = Habilitar
    LbCta.Enabled = Habilitar
    CmbCentrosDeCostos.Enabled = Habilitar
    CmdBuscarSubCentro.Enabled = Habilitar
    CmbCuentas.Enabled = Habilitar
    TxtCantCentroCta.Enabled = Habilitar
    CmdAgregarCentroCta.Enabled = Habilitar
    TxtCodCuenta.Enabled = Habilitar
    LbCodCuenta.Enabled = Habilitar
    If Habilitar Then
      ' LvCenCostoCtas.ListItems.Add
    Else
       LvCenCostoCtas.ListItems.Clear
       CmbCentrosDeCostos.ListIndex = 0
       CmbCuentas.ListIndex = 0
       CmdModifCentroCta.Enabled = False
       CmdEliminarCentroCta.Enabled = False
    End If
End Sub

Private Sub CargarLvCenCostoCtas(A_Codigo As Long)
  Dim i As Integer
    LvCenCostoCtas.ListItems.Clear
    For i = 1 To UBound(VecCentroCta)
      With VecCentroCta(i)
        If .O_CodigoArticulo = A_Codigo Then
            LvCenCostoCtas.ListItems.Add
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Text = BuscarDescCentroEmisor(BuscarCentroPadre(.O_CentroDeCosto))
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(1) = .Centro_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(2) = .Cta_Descripcion
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(3) = Format(.O_CantidadPedida, "0.00##")
            LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).SubItems(4) = i
        End If
      End With
    Next
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(LvCenCostoCtas.ListItems.Count).Selected = True
    Call LvCenCostoCtas_ItemClick(LvCenCostoCtas.SelectedItem)
    Call HabilitarAsignacion(True)
    Call CalcularTotalAsignado
End Sub

Private Sub LimpiarOC()
    CmbArtCompra.ListIndex = 0
    TxtCant.Text = ""
    TxtPrecioU.Text = ""
    TxtCodArticulo.Text = ""
    LbVariacion.Caption = ""
End Sub

Public Sub BuscarArt(Codigo As Long, CmbArt As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = Codigo Then
           Exit For
        End If
   Next
   If i <= UBound(VecArtCompra) Then
      CmbArt.ListIndex = i
   End If
End Sub

Public Function BuscarCtaPorDefectoArt(Codigo As Long) As String
 Dim i As Integer
   For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = Codigo Then
           BuscarCtaPorDefectoArt = VecArtCompra(i).A_CuentaPorDefecto
           Exit Function
        End If
   Next
End Function

Private Sub CargarEnModificar(Index As Integer)
    TxtCant.Enabled = Not VecOrdenDeCompra(Index).Requerimiento
    CmbArtCompra.Enabled = Not VecOrdenDeCompra(Index).Requerimiento
   Call BuscarArt(VecOrdenDeCompra(Index).A_Codigo, CmbArtCompra)
  'esta variable se usa para cargar el lvCentroCta
   A_Codigo = VecOrdenDeCompra(Index).A_Codigo
   
   TxtCant.Text = Replace(VecOrdenDeCompra(Index).Cantidad, ",", ".")
   TxtPrecioU.Text = Replace(VecOrdenDeCompra(Index).PrecioUnit, ",", ".")
   Call CalcularVariacionPrecio(VecOrdenDeCompra(Index).PrecioUnit)
End Sub

Private Sub CalcularVariacionPrecio(PrecioAct As Double)
        Dim Sql As String
        Dim PrecioAnt As Double
        Dim RsPrecios As New ADODB.Recordset
        
        Sql = "SpOcOrdenesDeCompraRenglonesUltimoPrecioArticulo @O_CodigoArticulo=" & A_Codigo & _
                                                                                              ", @O_CentroDeCostoEmisor='" & CentroEmisor & "'"
                                                                                              
        RsPrecios.Open Sql, Conec
        If Not RsPrecios.EOF Then
            PrecioAnt = RsPrecios!O_PrecioPactado
            LbVariacion.Caption = "Precio Ant. $" & Format(PrecioAnt, "0.00") & " var. " & Format((PrecioAct - PrecioAnt) / PrecioAnt, "0.00%")
        Else
            LbVariacion.Caption = ""
        End If
        
End Sub

Private Sub TxtCantCentroCta_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtCantCentroCta, KeyAscii)
End Sub

Private Sub LimpiarOrden()
    TxtCant.Text = ""
    TxtCantCentroCta.Text = ""
    CmbLugar.ListIndex = 0
    TxtResp.Text = NombreUsuario

    TxtPrecioU.Text = ""
    TxtObs.Text = ""
    TxtDescuento.Text = ""
    TxtDescuento.Enabled = True
    TxtObs.Enabled = True
    txtTotal.Text = Format(0, "$ 0.00##")
    'CmbCentroDeCostoEmisor.Enabled = True
    'CalFecha.Enabled = False
    NroOrden = 0
    
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    
    LvCenCostoCtas.ListItems.Clear
    LvCenCostoCtas.ListItems.Add
    LvCenCostoCtas.ListItems(1).Selected = True
    
    ReDim VecOrdenDeCompra(0)
    ReDim VecCentroCta(0)
    ReDim VecRequerimientoCompra(0)

    CmbArtCompra.ListIndex = 0
   ' CmbCentroDeCostoEmisor.ListIndex = 0
    CmbCentrosDeCostos.ListIndex = 0
    CmbCuentas.ListIndex = 0
    CmbEmp.ListIndex = 0
    CmbFormaDePago.ListIndex = 0
    CmbProv.ListIndex = 0

    CalFecha.Value = ValidarPeriodo(Date, False)
    
    CalFecha.Format = dtpCustom
    CalFecha.CustomFormat = " "
    CalFechaEmitida.Value = ValidarPeriodo(Date, False)
    CmdConfirnar.Visible = True
    CmdAnular.Visible = False
    CmdCambiar.Visible = False
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    LBPerCerrado.Visible = False
    
    CmdRequerimientos.Enabled = True
    TxtCodArticulo.Visible = TablaArticulos <> ""
    Call HabilitarAsignacion(False)
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
    CmdConfirnar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
  End If
End Sub

Private Sub CargarOrden(NroOrden As Long)
    Dim Sql As String
    Dim i As Integer

    Dim RsCargar As New ADODB.Recordset
    Dim PeriodoCerrado As Boolean
    Dim RsValidarPeriodo As New ADODB.Recordset
    Dim Autorizado As Boolean
    LBAnulada.Visible = False
    LBPerCerrado.Visible = False
    TotalMontoSinPres = 0
    
    If RsRenglones Is Nothing Then
        Set RsRenglones = New ADODB.Recordset
        RsRenglones.CursorType = adOpenKeyset
        RsRenglones.LockType = adLockBatchOptimistic
    Else
        If RsRenglones.State = adStateOpen Then
            RsRenglones.Close
        End If
    End If
    
  With RsCargar
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        '.LockType = adLockPessimistic
        
        Sql = "SpOCOrdenesDeCompraCabeceraTraerNro @NroOrden= " & NroOrden & _
                ", @Usuario='" & Usuario & "', @O_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
        .Open Sql, Conec
        
      If .EOF Then
          MsgBox "No existe una orden de compra con esa numeración", vbInformation
          Exit Sub
      Else
      
          CalFecha.Enabled = False
          'CalFecha.Format = dtpShortDate
          CalFecha.CustomFormat = "MM/yyyy"
          CmbCentroDeCostoEmisor.Enabled = False
          Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(!O_Fecha, "MM/yyyy")) & "'"
          RsValidarPeriodo.Open Sql, Conec
          PeriodoCerrado = RsValidarPeriodo!Cerrado > 0
    
        If Not IsNull(!O_FechaAnulacion) Or PeriodoCerrado Then
            If Not IsNull(!O_FechaAnulacion) Then
                LBAnulada.Caption = "Anulada " + Mid(CStr(!O_FechaAnulacion), 1, 10)
                LBAnulada.Visible = True
            End If
            CmdAnular.Visible = False
            LBPerCerrado.Visible = PeriodoCerrado
            FraArt.Enabled = False
            FrameAsig.Enabled = False
            CmdCambiar.Enabled = False
            CmbLugar.Enabled = False
            'TxtResp.Enabled = False
            'CmbCentroDeCostoEmisor.Enabled = False
            CmbEmp.Enabled = False
            CmbFormaDePago.Enabled = False
            CmbProv.Enabled = False
            CalFecha.Enabled = False
            TxtObs.Enabled = False
            TxtDescuento.Enabled = False
        Else
            CmdAnular.Visible = True
            LBAnulada.Visible = False
            FraArt.Enabled = True
            FrameAsig.Enabled = True
            CmdCambiar.Enabled = True
            CmbLugar.Enabled = True
            'TxtResp.Enabled = True
            'CmbCentroDeCostoEmisor.Enabled = True
            CmbEmp.Enabled = True
            CmbFormaDePago.Enabled = True
            TxtObs.Enabled = True
            TxtDescuento.Enabled = True
            CmbProv.Enabled = True
            CalFecha.Enabled = True
        End If
        
        TxtNroOrden.Text = Format(!O_NumeroOrdenDeCompra, "0000000000")
        Me.NroOrden = !O_NumeroOrdenDeCompra
        Autorizado = VerificarNulo(!O_Autorizado, "B")
        CmdConfirnar.Visible = False
        CmdCambiar.Visible = True
        CmdImprimir.Enabled = True
        CmdExpPdf.Enabled = True
        If IsNull(!O_FechaEmision) Then
            CalFechaEmitida.Value = !O_Fecha
        Else
            CalFechaEmitida.Value = !O_FechaEmision
        End If

        CalFecha.Value = !O_Fecha
        TxtResp = RsCargar!O_Responsable
        Call BuscarProveedor(!O_CodigoProveedor, CmbProv)
        
        If ValN(!O_CodigoLugarDeEntrega) = 0 Then
            CmbLugar.Text = !O_LugarDeEntrega
        Else
            Call UbicarCmbLugaresDeEntrega(ValN(!O_CodigoLugarDeEntrega), CmbLugar)
        End If
        
        If ValN(!O_CodigoFormaDePago) = 0 Then
            CmbFormaDePago.Text = !O_FormaDePagoPactada
        Else
            Call UbicarCmbFormasDePago(ValN(!O_CodigoFormaDePago), CmbFormaDePago)
        End If
        
        Call UbicarEmpresa(!O_EmpresaFacturaANombreDe, CmbEmp)
        Call BuscarCentroEmisor(!O_CentroDeCostoEmisor, CmbCentroDeCostoEmisor)
        TxtObs.Text = VerificarNulo(!O_Observaciones)
        TxtDescuento.Text = Replace(ValN(!O_Descuento), ",", ".")
        .Close
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        
        Sql = "SpOCOrdenesDeCompraRenglonesArticulosTraer @NroOrden=" & NroOrden & _
                                                       ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
        .Open Sql, Conec
        ReDim VecOrdenDeCompra(.RecordCount)
        i = 1
        LvListado.ListItems.Clear
        
        While Not .EOF
            
            LvListado.ListItems.Add
            VecOrdenDeCompra(i).A_Codigo = !O_CodigoArticulo
            VecOrdenDeCompra(i).A_Descripcion = BuscarDescArt(!O_CodigoArticulo)
            VecOrdenDeCompra(i).Cantidad = Format(!O_CantidadPedida, "0.00##")
            VecOrdenDeCompra(i).PrecioUnit = !O_PrecioPactado
            VecOrdenDeCompra(i).CantPendiente = !O_CantidadPendiente
            VecOrdenDeCompra(i).MontoSinPres = VerificarNulo(!MontoSinPres, "N")
            LvListado.ListItems(i).Text = VecOrdenDeCompra(i).A_Descripcion
            LvListado.ListItems(i).SubItems(1) = Format(VecOrdenDeCompra(i).Cantidad, "0.00##")
            LvListado.ListItems(i).SubItems(2) = Format(VecOrdenDeCompra(i).PrecioUnit, "0.00##")
            LvListado.ListItems(i).SubItems(3) = Format(VecOrdenDeCompra(i).Cantidad * VecOrdenDeCompra(i).PrecioUnit, "0.00##")
            
            .MoveNext
            i = i + 1
        Wend
        
        Call CalcularTotal
        LvListado.ListItems.Add
        LvListado.ListItems(LvListado.ListItems.Count).Selected = True
        .Close
        
        Sql = "SpOCOrdenesDeCompraRenglonesTraer @NroOrden=" & NroOrden & _
                                              ", @O_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
        With RsRenglones
            .Open Sql, ConecSvr
            ReDim VecCentroCta(0)
            i = 1
            
            While Not .EOF
               ReDim Preserve VecCentroCta(UBound(VecCentroCta) + 1)
                VecCentroCta(i).Centro_Descripcion = BuscarDescCentro(!O_CentroDeCosto)
                VecCentroCta(i).Cta_Descripcion = BuscarDescCta(!O_CuentaContable)
                VecCentroCta(i).O_CantidadPedida = !O_CantidadPedida
                VecCentroCta(i).O_CodigoArticulo = !O_CodigoArticulo
                VecCentroCta(i).O_CentroDeCosto = !O_CentroDeCosto
                VecCentroCta(i).O_CuentaContable = !O_CuentaContable
                VecCentroCta(i).O_CantidadPendiente = !O_CantidadPendiente
                VecCentroCta(i).O_SinPresupuestar = VerificarNulo(!O_SinPresupuestar, "B")
                VecCentroCta(i).O_MontoSinPresupuestar = VerificarNulo(!O_MontoSinPresupuestar, "N")
                If VecCentroCta(i).O_SinPresupuestar Then
                   TotalMontoSinPres = TotalMontoSinPres + VerificarNulo(!O_MontoSinPresupuestar, "N") 'VecCentroCta(i).O_CantidadPedida * BuscarPrecio(VecCentroCta(i).O_CodigoArticulo)
                End If
                i = i + 1
                'esto es para marque el registro como modificado
                'y que detecte eventuales cambios
                RsRenglones.Fields("O_CantidadPendiente").Value = 0
                .MoveNext
            Wend
        End With
      End If
  End With
  
  LvListado.ListItems(1).Selected = True
  Call LvListado_ItemClick(LvListado.ListItems(1))
  'Call HabilitarAsignacion(True)
  If Not Autorizado Then
      MsgBox "La Orden no ha sido autorizada ahún", vbInformation
      CmdImprimir.Enabled = False
      CmdExpPdf.Enabled = False
  Else
      CmdImprimir.Enabled = True
      CmdExpPdf.Enabled = True
  End If
End Sub

Private Sub CargarVecCentroEmisor(CentroEmisor As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
 'dependiendo del centro de costo emisor carga las cuentas y artículos correspondientes
    Sql = "SpOCRelacionCentroDeCostoArticulosTraer @R_CentroDeCosto='" & CentroEmisor & "'"
    With RsCargar
      ReDim VecArtCompra(0)
    If TablaArticulos = "" Then
       .Open Sql, Conec, adOpenStatic, adLockReadOnly
      'en esta sección carga los art
        For i = 1 To UBound(VariablesYFunciones.VecArtCompra)
            .Find "R_Articulo = " & VariablesYFunciones.VecArtCompra(i).A_Codigo, , , 1
           If Not .EOF Then
              ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
              VecArtCompra(UBound(VecArtCompra)) = VariablesYFunciones.VecArtCompra(i)
           End If
        Next
      .Close
     Else
        If TablaArticulos = "MotoVan" Then
            For i = 1 To UBound(VecArtMotoVan)
                ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                VecArtCompra(UBound(VecArtCompra)) = VecArtMotoVan(i)
            Next
        Else
            For i = 1 To UBound(VecArtTaller)
                ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                VecArtCompra(UBound(VecArtCompra)) = VecArtTaller(i)
            Next
        End If
    
     End If
      
      ReDim VecCuentasContables(0)
      'en esta sección carga las cuentas
      Sql = "SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CentroEmisor & "'"
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
    Call CargarCmbArtCompra(CmbArtCompra)
    Call CargarCmbCuentasContables(CmbCuentas)
End Sub

Public Sub CargarCmbArtCompra(CmbAtrCompra As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbAtrCompra.Clear
    For i = 0 To UBound(VecArtCompra)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbAtrCompra.AddItem "Seleccione un Artículo"
           Else
              CmbAtrCompra.AddItem "Todos los Artículos"
           End If
        Else
            CmbAtrCompra.AddItem VecArtCompra(i).A_Descripcion
        End If
    Next
        
    CmbAtrCompra.ListIndex = 0
Errores:
   Call ManipularError(Err.Number, Err.Description)
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

Public Function BuscarDescArt(A_Codigo As Long) As String
    Dim i As Integer
    For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = A_Codigo Then
            BuscarDescArt = VecArtCompra(i).A_Descripcion
            Exit Function
        End If
    Next
End Function

Private Function CalcularTotalSinPresupuestar() As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        Total = Total + VecCentroCta(i).O_MontoSinPresupuestar
    Next
       
    CalcularTotalSinPresupuestar = Total
End Function

Private Function CalcularTotalSinPresupuestarPorCuenta(Cuenta As String) As Double
    Dim i As Integer
    Dim Total As Double
    
    For i = 1 To UBound(VecCentroCta)
        If i <> Val(LvCenCostoCtas.SelectedItem.SubItems(4)) And _
           Cuenta = VecCentroCta(i).O_CuentaContable Then
           Total = Total + VecCentroCta(i).O_MontoSinPresupuestar
        End If
    Next
       
    CalcularTotalSinPresupuestarPorCuenta = Total
End Function

Private Function ValidarIntegridad() As Boolean
    ValidarIntegridad = True
    If Not RsRenglones Is Nothing Then
        With RsRenglones
    
            If .State = adStateOpen Then
                .MoveFirst
                While Not .EOF
                     .Fields("O_CantidadPendiente").Value = 0
                     
                    If .Fields("O_CantidadPendiente").OriginalValue <> .Fields("O_CantidadPendiente").UnderlyingValue Then
                        MsgBox "Los datos fueron modificados por otro usuario intente mas tarde", vbCritical
                        ValidarIntegridad = False
                        Exit Function
                    End If
                    .MoveNext
                Wend
            End If
        End With
    End If
End Function

Private Sub EnviarMail(NroOrden As Integer)
Dim Mensaje As String
Dim i As Integer
Dim EMail As String
Dim inicio As Integer
Dim Fin As Integer
'omite los errores
On Error Resume Next
    If Trim(VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail) = "" Then
        Exit Sub
    End If
    
    Call ConfImpresionDeOrden
    RepOrdenDeCompra.Run
    'guarda la orden de compra como un PDF
    Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
    Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
    myPDFExport.AcrobatVersion = DDACR40
     
    myPDFExport.Filename = "C:\Orden " & TxtNroOrden.Text & " " & CmbCentroDeCostoEmisor.Text & ".pdf"
    myPDFExport.JPGQuality = 100
    myPDFExport.SemiDelimitedNeverEmbedFonts = ""
    myPDFExport.Export RepOrdenDeCompra.Pages
    Unload RepOrdenDeCompra
    
    MAPISession.DownLoadMail = False
    MAPISession.SignOn
    MAPIMessages.SessionID = MAPISession.SessionID
    MAPIMessages.Compose
    'MAPIMessages.RecipAddress = VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail
    
    '******** varios E-Mail **********
    EMail = VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail
    i = 0
    inicio = 1
    While Fin < Len(EMail)
        Fin = Fin + 1
        If Mid(EMail, Fin, 1) = ";" Then
            MAPIMessages.RecipIndex = i
            MAPIMessages.RecipType = 1
            MAPIMessages.RecipAddress = Trim(Mid(EMail, inicio, Fin - inicio))
            inicio = Fin + 1
            i = i + 1
        End If
    Wend
    
    MAPIMessages.RecipIndex = i
    MAPIMessages.RecipType = 1
    MAPIMessages.RecipAddress = Trim(Mid(EMail, inicio))
    '****** Fin Varios E-Mail **********
    
    'MAPIMessages.AddressResolveUI = False
    'MAPIMessages.ResolveName
    MAPIMessages.MsgSubject = "Orden de Compra Nº: " & Format(NroOrden, "0000000000") & " ,Centro de Costo Emisor: " & CmbCentroDeCostoEmisor.Text
    
    'MAPIMessages.MsgNoteText = Mensaje
   
    MAPIMessages.AttachmentPathName = myPDFExport.Filename
    MAPIMessages.Send False
    MAPISession.SignOff
    Call Kill(myPDFExport.Filename)
End Sub

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
    RepOrdenDeCompra.Run
    'guarda la orden de compra como un PDF
    Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
    Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
    myPDFExport.AcrobatVersion = DDACR40
     
    myPDFExport.Filename = "C:\Orden " & TxtNroOrden.Text & " " & CmbCentroDeCostoEmisor.Text & ".pdf"
    myPDFExport.JPGQuality = 100
    myPDFExport.SemiDelimitedNeverEmbedFonts = ""
    myPDFExport.Export RepOrdenDeCompra.Pages
    Unload RepOrdenDeCompra
    
    MAPISession.DownLoadMail = False
    MAPISession.SignOn
    MAPIMessages.SessionID = MAPISession.SessionID
    MAPIMessages.Compose
    'MAPIMessages.RecipAddress = VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail
    
    '******** varios E-Mail **********
    'EMail = VecLugaresDeEntrega(CmbLugar.ListIndex).L_EMail
    i = 0
    inicio = 1
    While Fin < Len(EMail)
        Fin = Fin + 1
        If Mid(EMail, Fin, 1) = ";" Then
            MAPIMessages.RecipIndex = i
            MAPIMessages.RecipType = 1
            MAPIMessages.RecipAddress = Trim(Mid(EMail, inicio, Fin - inicio))
            inicio = Fin + 1
            i = i + 1
        End If
    Wend
    
    MAPIMessages.RecipIndex = i
    MAPIMessages.RecipType = 1
    MAPIMessages.RecipAddress = Trim(Mid(EMail, inicio))
    '****** Fin Varios E-Mail **********
    
    'MAPIMessages.AddressResolveUI = False
    'MAPIMessages.ResolveName
    MAPIMessages.MsgSubject = "Autorizar Orden de Compra Nº: " & Format(NroOrden, "0000000000") & " ,Centro de Costo Emisor: " & CmbCentroDeCostoEmisor.Text
    
    'MAPIMessages.MsgNoteText = Mensaje
   
    MAPIMessages.AttachmentPathName = myPDFExport.Filename
    MAPIMessages.Send False
    MAPISession.SignOff
ErrorEMail:
 On Error Resume Next
    If Err.Number <> 0 Then
        MousePointer = vbHourglass
        Call EnviarEmail(EMail, EMail, "Autorizar Orden de Compra Nº: " & Format(NroOrden, "0000000000") & " ,Centro de Costo Emisor: " & CmbCentroDeCostoEmisor.Text, "", "C:\Orden " & TxtNroOrden.Text & " " & CmbCentroDeCostoEmisor.Text & ".pdf")
        MousePointer = vbNormal
    End If
    Call Kill(myPDFExport.Filename)
End Sub

Private Sub TxtPrecioU_Validate(Cancel As Boolean)
    If Val(TxtDescuento.Text) > 0 Then
        TxtPrecioU.Text = Replace(Val(TxtPrecioU.Text) - (Val(TxtPrecioU.Text) * Val(TxtDescuento.Text) / 100), ",", ".")
    End If
End Sub
