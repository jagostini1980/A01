VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_4100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Presupuestos"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMulti 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   45
      TabIndex        =   38
      Top             =   6570
      Width           =   11085
      Begin VB.TextBox TxtCantOrdenes 
         Height          =   315
         Left            =   1350
         TabIndex        =   40
         Text            =   "12"
         Top             =   900
         Width           =   645
      End
      Begin VB.CommandButton CmdCrear 
         Caption         =   "Cr&ear"
         Height          =   350
         Left            =   9630
         TabIndex        =   39
         Top             =   1710
         Width           =   1300
      End
      Begin Controles.ComboEsp CmbFrecuencia 
         Height          =   315
         Left            =   90
         TabIndex        =   41
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
         Left            =   2070
         TabIndex        =   42
         Top             =   180
         Width           =   7440
         _ExtentX        =   13123
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
         TabIndex        =   44
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad de Presupuestos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   43
         Top             =   855
         Width           =   1230
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10665
      Top             =   5175
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular "
      Height          =   350
      Left            =   4230
      TabIndex        =   34
      Top             =   6210
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdNueva 
      Caption         =   "&Nueva"
      Height          =   350
      Left            =   5625
      TabIndex        =   17
      Top             =   6210
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10125
      Top             =   5220
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación"
      Height          =   2805
      Left            =   7380
      TabIndex        =   27
      Top             =   2250
      Width           =   3795
      Begin VB.TextBox TxtObservaciones 
         Height          =   900
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1395
         Width           =   3615
      End
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   11
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox TxtMonto 
         Height          =   315
         Left            =   2115
         TabIndex        =   12
         Top             =   810
         Width           =   870
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   1290
         TabIndex        =   15
         Top             =   2340
         Width           =   1150
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   16
         Top             =   2340
         Width           =   1150
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   90
         TabIndex        =   14
         Top             =   2340
         Width           =   1150
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   450
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
      Begin VB.Label Label6 
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
         TabIndex        =   48
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   3060
         TabIndex        =   47
         Top             =   225
         Width           =   405
      End
      Begin VB.Label LbCant 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Presupuestado:"
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
         TabIndex        =   29
         Top             =   855
         Width           =   1920
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
         TabIndex        =   28
         Top             =   225
         Width           =   1485
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7020
      TabIndex        =   18
      Top             =   6210
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
      Left            =   9810
      TabIndex        =   26
      Text            =   "0"
      Top             =   5850
      Width           =   1275
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   3930
      Left            =   -15
      TabIndex        =   9
      Top             =   2235
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6932
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
      Left            =   9855
      TabIndex        =   21
      Top             =   6210
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Presupuesta"
      Height          =   1995
      Left            =   45
      TabIndex        =   22
      Top             =   45
      Width           =   11130
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
         Left            =   1530
         TabIndex        =   46
         Top             =   1665
         Value           =   -1  'True
         Width           =   2445
      End
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
         Left            =   4050
         TabIndex        =   45
         Top             =   1665
         Width           =   1950
      End
      Begin VB.CommandButton CmdCopiar 
         Caption         =   "&Copiar"
         Height          =   315
         Left            =   5490
         TabIndex        =   3
         Top             =   225
         Width           =   1000
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1305
         Width           =   9510
      End
      Begin VB.TextBox TxtResp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3105
         MaxLength       =   50
         TabIndex        =   7
         Top             =   945
         Width           =   5730
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   900
         TabIndex        =   5
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM/yyyy"
         Format          =   53870595
         UpDown          =   -1  'True
         CurrentDate     =   38980
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   3225
         TabIndex        =   1
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4365
         TabIndex        =   2
         Top             =   225
         Width           =   1000
      End
      Begin VB.TextBox TxtNroPresupuesto 
         Height          =   315
         Left            =   1890
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   9630
         TabIndex        =   4
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53870593
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   5220
         TabIndex        =   6
         Top             =   600
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   556
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
         TabIndex        =   37
         Top             =   1350
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
         Left            =   6615
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label LBAnulada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anulado"
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
         Left            =   6705
         TabIndex        =   33
         Top             =   135
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Responsable del Centro de Costo:"
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
         TabIndex        =   32
         Top             =   1005
         Width           =   2910
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período:"
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
         TabIndex        =   31
         Top             =   668
         Width           =   750
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
         Left            =   3105
         TabIndex        =   30
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Presupuesto:"
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
         Top             =   270
         Width           =   1665
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
         Left            =   8955
         TabIndex        =   23
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   8400
      TabIndex        =   20
      Top             =   6195
      Width           =   1300
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   8415
      TabIndex        =   19
      Top             =   6210
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuentas Contables"
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
      Left            =   0
      TabIndex        =   35
      Top             =   2025
      Width           =   1605
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
      Left            =   9225
      TabIndex        =   25
      Top             =   5895
      Width           =   510
   End
End
Attribute VB_Name = "A01_4100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoPresupuesto
   O_CuentaContable As String
   Cta_Descripcion As String
   Borrable As Boolean
   FechaAprobacion As String
   Monto As Double
   P_Observaciones As String
   P_ObservacionesPresupuesto As String
End Type

Private VecPresupuesto() As TipoPresupuesto
Private Modificado As Boolean
Private A_Codigo As Long
Public NroPresupuesto As Integer
Dim PeriodoCerrado As Boolean
Dim Anulado As Boolean
Private Nivel As Integer
Public CentroEmisorActual As String
Public TablaArticulos As String
Private VecCuentasContables() As CuentasContables

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalPeriodo_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Sub CalPeriodo_LostFocus()
    CalPeriodo.Value = ValidarPeriodo(CalPeriodo.Value)
End Sub

Private Sub CmbCentroDeCostoEmisor_Click()
    'Modificado = True
    CentroEmisorActual = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo
    TablaArticulos = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_TablaArticulos

    Call CargarVecCentroEmisor(CentroEmisorActual)
End Sub

Private Sub CmbCuentas_Click()
On Error Resume Next
      TxtCodCuenta.Text = VecCuentasContables(CmbCuentas.ListIndex).Codigo
      
End Sub

Private Sub CmbCuentas_Validate(Cancel As Boolean)
      'para gastos de Bar
      TxtMonto.Enabled = VecCuentasContables(CmbCuentas.ListIndex).Codigo <> "5023"
      If VecCuentasContables(CmbCuentas.ListIndex).Codigo = "5023" Then
            A01_4120.TxtCentroDeCostoEmisor = CmbCentroDeCostoEmisor.Text
            A01_4120.CalPeriodo = CalPeriodo
            A01_4120.TxtCuentaContable = CmbCuentas.Text
            A01_4120.TxtNroPresupuesto = TxtNroPresupuesto
            A01_4120.Show vbModal
      End If

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

Private Sub CmdAnular_Click()
 Dim Sql As String
 Dim Rta As Integer
 Dim RsAnular As ADODB.Recordset
 Set RsAnular = New ADODB.Recordset
 
 On Error GoTo Error
    Rta = MsgBox("¿Está seguro de que desea Anular El Presupuesto?", vbYesNo)
    If Rta = vbYes Then
        Sql = "SpOCPresupuestosCabeceraAnular @P_NumeroPresupuesto =" & NroPresupuesto & _
                                           ", @P_CentroDeCostoEmisor = '" & CentroEmisorActual & "'"
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
        Call LimpiarPresupuesto
     Else
        Unload Me
     End If
  End If

End Sub

Private Sub CMDBuscar_Click()
    NroPresupuesto = 0
    
    BuscarPresupuesto.CmbCentroDeCostoEmisor.Visible = Nivel = 2
    BuscarPresupuesto.LbCentroEmisor.Visible = Nivel = 2
    BuscarPresupuesto.Show vbModal
    Timer1.Enabled = True
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar el Presupuesto?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarPresupuesto
    End If
End Sub

Private Sub ModificarPresupuesto()
  Dim Sql As String
  Dim i As Integer
  Dim Monto As String
  
On Error GoTo ErrorUpdate

  NroPresupuesto = Val(TxtNroPresupuesto.Text)
  If ValidarEncabezado Then
    Conec.BeginTrans
     Sql = "SpOCPresupuestosCabeceraModificar @P_NumeroPresupuesto=" + CStr(NroPresupuesto) + _
           ", @P_FechaEmision=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
           ", @P_Responsable ='" + TxtResp.Text + _
           "', @U_Usuario = '" + Usuario + _
           "', @P_CentroDeCostoEmisor= '" + CStr(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo) + _
           "', @P_Periodo = '" + CStr(Format(CalPeriodo.Value, "MM/yyyy")) + _
           "', @P_Observaciones = '" & TxtObs.Text & "'"
           
        Conec.Execute Sql
       Sql = "SpOCPresupuestosRenglonesBorrar @P_NumeroPresupuesto=" & NroPresupuesto & _
                                           ", @P_CentroDeCostoEmisor=" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo
        Conec.Execute Sql
    'RENGLONES DE LAS CUENTAS
    For i = 1 To UBound(VecPresupuesto)
      With VecPresupuesto(i)
      
        Monto = Replace(.Monto, ",", ".")
        
        Sql = "SpOCPresupuestosRenglonesAgregar @P_NumeroPresupuesto = " & CStr(NroPresupuesto) & _
                ", @P_CuentaContable ='" & .O_CuentaContable & _
               "', @P_ImporteUnitario =" & Monto & _
                ", @P_CentroDeCostoEmisor = '" & CentroEmisorActual & _
               "', @@P_FechaAprobacion= " & IIf(.FechaAprobacion = "", "NULL", .FechaAprobacion) & _
                ", @@P_Observaciones = '" & .P_Observaciones & _
               "', @P_ObservacionesPresupuesto= '" & .P_ObservacionesPresupuesto & "'"
                
      End With
        Conec.Execute Sql
    Next
    
    For i = 1 To UBound(VecDistribucionPresupuesto)
      With VecDistribucionPresupuesto(i)
      
        Monto = Replace(.P_Importe, ",", ".")
        
        Sql = "SpOcPresupuestosDistrubucionActualizar @P_NumeroPresupuesto =" & NroPresupuesto & _
                " , @P_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                "', @P_SubCentroDeCosto ='" & .P_SubCentroDeCosto & _
                "', @P_CuentaContable ='" & .P_CuentaContable & _
                "', @P_Periodo ='" & Format(CalFecha.Value, "MM/yyyy") & _
                "', @P_Importe =" & Monto
      End With
        Conec.Execute Sql
    Next
    'distribucion servicio de Bat
    Sql = "SpOcPresupuestosServicioBarBorrar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                          ", @D_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
    Conec.Execute Sql
    For i = 1 To UBound(VecPresServicioDeBar)
      With VecPresServicioDeBar(i)
        
        Sql = "SpOcPresupuestosServicioBarAgregar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                               ", @D_Cuenta ='" & .D_Cuenta & _
                                              "', @D_Linea ='" & .D_Linea & _
                                              "', @D_Proveedor =" & .D_Proveedor & _
                                              " , @D_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                                              "', @D_Cantidad =" & .D_Cantidad & _
                                              " , @D_PrecioUnitario =" & Replace(.D_PrecioUnitario, ",", ".") & _
                                              " , @D_Periodo='" & Format(CalPeriodo.Value, "MM/yyyy") & "'"
      End With
        Conec.Execute Sql
    Next

    Conec.CommitTrans
ErrorUpdate:
    If Err.Number = 0 Then
       MsgBox "El Presupuesto se Grabó correctamente"
       Modificado = False
    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Sub

Private Sub CmdCargar_Click()
    Call CargarPresupuesto(Val(TxtNroPresupuesto))
    Modificado = False

End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar el Presupuesto?", vbYesNo)
    If Rta = vbYes Then
        Call GrabarPresupuesto
    End If
End Sub

Private Sub GrabarPresupuesto()
  Dim Sql As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer
  Dim Monto As String

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
     Conec.BeginTrans
     Sql = "SpOCPresupuestosCabeceraAgregar @P_FechaEmision=" + FechaSQL(CStr(CalFecha.Value), "SQL") + _
            ", @P_Responsable ='" + TxtResp.Text + _
           "', @U_Usuario = '" + Usuario + _
           "', @P_CentroDeCostoEmisor= '" + CStr(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo) + _
           "', @P_Periodo = '" + CStr(Format(CalPeriodo.Value, "MM/yyyy")) + _
           "', @P_Observaciones = '" & TxtObs.Text & "'"
           
     'graba el encabezado y retorna el Nro de presupuesto
        RsGrabar.Open Sql, Conec
        NroPresupuesto = RsGrabar!P_NumeroDePresupuesto
    'RENGLONES DE LAS CUENTAS
    For i = 1 To UBound(VecPresupuesto)
      With VecPresupuesto(i)
      
        Monto = Replace(.Monto, ",", ".")
        
        Sql = "SpOCPresupuestosRenglonesAgregar @P_NumeroPresupuesto = " & CStr(NroPresupuesto) & _
                ", @P_CuentaContable ='" & .O_CuentaContable & _
               "', @P_ImporteUnitario =" & Monto & _
                ", @P_CentroDeCostoEmisor = '" & CentroEmisorActual & _
               "', @P_ObservacionesPresupuesto= '" & .P_ObservacionesPresupuesto & "'"
                
      End With
        Conec.Execute Sql
    Next
    
    For i = 1 To UBound(VecDistribucionPresupuesto)
      With VecDistribucionPresupuesto(i)
      
        Monto = Replace(.P_Importe, ",", ".")
        
        Sql = "SpOcPresupuestosDistrubucionActualizar @P_NumeroPresupuesto =" & NroPresupuesto & _
                " , @P_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                "', @P_SubCentroDeCosto ='" & .P_SubCentroDeCosto & _
                "', @P_CuentaContable ='" & .P_CuentaContable & _
                "', @P_Periodo ='" & Format(CalPeriodo.Value, "MM/yyyy") & _
                "', @P_Importe =" & Monto
      End With
        Conec.Execute Sql
    Next
    'distribucion servicio de Bat
    Sql = "SpOcPresupuestosServicioBarBorrar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                          ", @D_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
    Conec.Execute Sql
    For i = 1 To UBound(VecPresServicioDeBar)
      With VecPresServicioDeBar(i)
        
        Sql = "SpOcPresupuestosServicioBarAgregar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                               ", @D_Cuenta ='" & .D_Cuenta & _
                                              "', @D_Linea ='" & .D_Linea & _
                                              "', @D_Proveedor =" & .D_Proveedor & _
                                              " , @D_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                                              "', @D_Cantidad =" & .D_Cantidad & _
                                              " , @D_PrecioUnitario =" & Replace(.D_PrecioUnitario, ",", ".") & _
                                              " , @D_Periodo='" & Format(CalPeriodo.Value, "MM/yyyy") & "'"
      End With
        Conec.Execute Sql
    Next
    
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       CmdConfirmar.Visible = False
       CmdCambiar.Visible = True
       CmdImprimir.Enabled = True
       
       'Rta = MsgBox("El Presupuesto se Grabó correctamente con el Nº: " + CStr(NroPresupuesto) + " ¿Desea imprimirla?", vbYesNo)
       Modificado = False
       FrmMensaje.LbMensaje = "El Presupuesto se Grabó correctamente con el Nº: " + CStr(NroPresupuesto) & _
                              Chr(13) & " ¿Que desea hacer?"
       FrmMensaje.CmdExportar.Enabled = False
       FrmMensaje.Show vbModal
       
       Modificado = False
       If FrmMensaje.Retorno = vbimprimir Then
         Call ConfImpresionDePresupuesto
         RepPresupuesto.Show vbModal
       End If
         
       If FrmMensaje.Retorno = vbNuevo Then
         Call LimpiarPresupuesto
       End If
       
       'If FrmMensaje.Retorno = vbExportesPDF Then
       '   CmdExpPdf_Click
       'End If
       
       If FrmMensaje.Retorno = vbCerrar Then
          Unload Me
       End If

    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Sub

Private Function ValidarEncabezado() As Boolean
Dim i As Integer
Dim Asignado As Boolean

    ValidarEncabezado = True

    If TxtResp.Text = "" Then
       ' MsgBox "Debe ingresar el responsable"
       ' TxtResp.SetFocus
       ' ValidarEncabezado = False
    Else
       If CmbCentroDeCostoEmisor.ListIndex = 0 Then
          MsgBox "Debe Seleccionar un Centro de Costo Emisor"
          CmbCentroDeCostoEmisor.SetFocus
          ValidarEncabezado = False
       End If
    End If
   If LvListado.ListItems.Count <= 1 Then
        MsgBox "Debe Ingresar Cuentas al Presupuesto"
        LvListado.SetFocus
        ValidarEncabezado = False
   End If
   
   If OptMulti(1).Value Then
       Dim mes As Boolean
       For i = 1 To LvMeses.ListItems.Count
            If LvMeses.ListItems(i).Checked Then
               mes = True
               Exit For
            End If
       Next
       If Not mes Then
            MsgBox "Debe Seleccionas algún mes"
            LvMeses.SetFocus
            ValidarEncabezado = False
       End If
   End If
   
  Dim Sql As String
  Dim RsValidarPeriodo As ADODB.Recordset
  Set RsValidarPeriodo = New ADODB.Recordset
  Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & CStr(Format(CalPeriodo.Value, "MM/yyyy")) & "'"
  RsValidarPeriodo.Open Sql, Conec
    If RsValidarPeriodo!Cerrado > 0 Then
       MsgBox "El período está Cerrado", vbExclamation, "Período Cerrado"
       CalPeriodo.SetFocus
       ValidarEncabezado = False
       Exit Function
    End If

End Function

Private Sub CmdCopiar_Click()
    NroPresupuesto = 0
    
    BuscarPresupuesto.CmbCentroDeCostoEmisor.Visible = False
    BuscarPresupuesto.LbCentroEmisor.Visible = False
    Call BuscarCentroEmisor(CentroEmisor, BuscarPresupuesto.CmbCentroDeCostoEmisor)
    BuscarPresupuesto.Show vbModal
    Timer2.Enabled = True

End Sub

Private Sub CmdCrear_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar la Multiples Presupuestos?", vbYesNo)
    If Rta = vbYes Then
        Call CrearMultiPresupuestos
    End If

End Sub

Private Function CrearMultiPresupuestos() As Boolean
  Dim Sql As String
  Dim Mensaje As String
  Dim RsGrabar As ADODB.Recordset
  Set RsGrabar = New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer
  Dim j As Integer
  Dim Monto As String

On Error GoTo ErrorInsert
  CrearMultiPresupuestos = False
  
  If ValidarEncabezado Then
    CrearMultiPresupuestos = True
    
    Conec.BeginTrans
    
    For j = 1 To LvMeses.ListItems.Count
      If LvMeses.ListItems(j).Checked Then
        Sql = "SpOCPresupuestosCabeceraAgregar @P_FechaEmision=" + FechaSQL(CalFecha.Value, "SQL") + _
                                            ", @P_Responsable ='" + TxtResp.Text + _
                                           "', @U_Usuario = '" + Usuario + _
                                           "', @P_CentroDeCostoEmisor= '" + CStr(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo) + _
                                           "', @P_Periodo = '" & Format("01/" & LvMeses.ListItems(j).Text, "MM/YYYY") & _
                                           "', @P_Observaciones = '" & TxtObs.Text & "'"
         'graba el encabezado y retorna el Nro de presupuesto
            RsGrabar.Open Sql, Conec
            NroPresupuesto = RsGrabar!P_NumeroDePresupuesto
            RsGrabar.Close
            Mensaje = Mensaje & " - " & LvMeses.ListItems(j).Text & " Nº: " & Format(NroPresupuesto, "0000000000")
        For i = 1 To UBound(VecPresupuesto)
            With VecPresupuesto(i)
      
              Monto = Replace(.Monto, ",", ".")
              
              Sql = "SpOCPresupuestosRenglonesAgregar @P_NumeroPresupuesto = " & CStr(NroPresupuesto) & _
                      ", @P_CuentaContable ='" & .O_CuentaContable & _
                     "', @P_ImporteUnitario =" & Monto & _
                      ", @P_CentroDeCostoEmisor = '" & CentroEmisorActual & _
                     "', @P_ObservacionesPresupuesto= '" & .P_ObservacionesPresupuesto & "'"
                      
            End With

            Conec.Execute Sql
        Next
        
        For i = 1 To UBound(VecDistribucionPresupuesto)
          With VecDistribucionPresupuesto(i)
          
            Monto = Replace(.P_Importe, ",", ".")
            
            Sql = "SpOcPresupuestosDistrubucionActualizar @P_NumeroPresupuesto =" & NroPresupuesto & _
                    " , @P_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                    "', @P_SubCentroDeCosto ='" & .P_SubCentroDeCosto & _
                    "', @P_CuentaContable ='" & .P_CuentaContable & _
                    "', @P_Periodo ='" & Format("01/" & LvMeses.ListItems(j).Text, "MM/YYYY") & _
                    "', @P_Importe =" & Monto
          End With
            Conec.Execute Sql
        Next
        'distribucion servicio de Bat
        Sql = "SpOcPresupuestosServicioBarBorrar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                              ", @D_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
        Conec.Execute Sql
        For i = 1 To UBound(VecPresServicioDeBar)
          With VecPresServicioDeBar(i)
            
            Sql = "SpOcPresupuestosServicioBarAgregar @D_NumeroPresupuesto =" & NroPresupuesto & _
                                                   ", @D_Cuenta ='" & .D_Cuenta & _
                                                  "', @D_Linea ='" & .D_Linea & _
                                                  "', @D_Proveedor =" & .D_Proveedor & _
                                                  " , @D_CentroDeCostosEmisor ='" & CentroEmisorActual & _
                                                  "', @D_Cantidad =" & .D_Cantidad & _
                                                  " , @D_PrecioUnitario =" & Replace(.D_PrecioUnitario, ",", ".") & _
                                                  " , @D_Periodo='" & Format("01/" & LvMeses.ListItems(j).Text, "MM/YYYY") & "'"
          End With
            Conec.Execute Sql
        Next

      End If
      
    Next
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
    
      Rta = MsgBox("Los Presupuestos se Grabaron correctamente " & Mensaje & " ¿Desea Crear otro Presupuesto?", vbYesNo)
      Modificado = False
      If Rta = vbYes Then
         Call LimpiarPresupuesto
         CrearMultiPresupuestos = False
      Else
         Unload Me
      End If

    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Function

Private Sub CmdEliminar_Click()
    Dim IndexBorrar As Integer
    
    IndexBorrar = LvListado.SelectedItem.Index
    'borra del LV
   If VecPresupuesto(IndexBorrar).Borrable Then
    LvListado.ListItems.Remove (IndexBorrar)
    'borrar del vector haciento un corrimiento
    While IndexBorrar < UBound(VecPresupuesto)
        VecPresupuesto(IndexBorrar) = VecPresupuesto(IndexBorrar + 1)
        IndexBorrar = IndexBorrar + 1
    Wend
    
    
    ReDim Preserve VecPresupuesto(UBound(VecPresupuesto) - 1)
       
  'calcula el total de la orden
  Call CalcularTotal
        
     Modificado = True
    
    If LvListado.ListItems.Count = LvListado.SelectedItem.Index Then
        CmdEliminar.Enabled = False
        CmdModif.Enabled = False
    End If
    
 Else
    MsgBox "El renglón no puede ser borrado", , "Borrar"
 End If
End Sub

Private Sub ConfImpresionDePresupuesto()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 150
    RsListado.Fields.Append "Observaciones", adVarChar, 150
    RsListado.Fields.Append "Monto", adDouble
    RsListado.Open
    i = 1
    While i < LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
           RsListado!CuentaContable = .Text
           RsListado!Monto = Val(Replace(.SubItems(1), ",", "."))
           RsListado!Observaciones = .SubItems(2)
      End With
        i = i + 1
    Wend
    
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    
    TxtNroPresupuesto.Text = Format(CStr(NroPresupuesto), "0000000000")
    RepPresupuesto.TxtFecha = CStr(CalFecha.Value)
    RepPresupuesto.TxtNroPresupuesto.Text = TxtNroPresupuesto.Text
    RepPresupuesto.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    RepPresupuesto.TxtResp.Text = TxtResp.Text
    RepPresupuesto.TxtAnulada.Visible = LBAnulada.Visible
    RepPresupuesto.TxtAnulada.Text = LBAnulada.Caption
    RepPresupuesto.TxtObservaciones = TxtObs
    RepPresupuesto.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    RepPresupuesto.DataControl1.Recordset = RsListado
    RepPresupuesto.Zoom = -1
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDePresupuesto
    RepPresupuesto.Show
End Sub

Private Sub CmdModif_Click()
'On Error GoTo errores
Dim i As Integer
    
 If ValidarCargaPresupuesto Then
         i = LvListado.SelectedItem.Index

    If VecPresupuesto(i).Borrable Then
       Modificado = True
        
      'agrega al vector
      
       VecPresupuesto(i).Cta_Descripcion = CmbCuentas.Text & " - Cód. " & VecCuentasContables(CmbCuentas.ListIndex).Codigo
       VecPresupuesto(i).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
       VecPresupuesto(i).Monto = Val(TxtMonto.Text)
       VecPresupuesto(i).P_ObservacionesPresupuesto = TxtObservaciones
       LvListado.ListItems(i).Text = CmbCuentas.Text & " - Cód. " & VecCuentasContables(CmbCuentas.ListIndex).Codigo
       LvListado.ListItems(i).SubItems(1) = Replace(Val(TxtMonto.Text), ",", ".")
       LvListado.ListItems(i).SubItems(2) = TxtObservaciones
    Else
       MsgBox "El renglón no puede ser modificado por estar aprobado", , "Modificar"
    End If
 End If
  
  'calcula el total de la orden
   Call CalcularTotal
  
Errores:
   Call ManipularError(Err.Number, Err.Description, Timer1)
End Sub

Private Sub CmdNueva_Click()
    TxtNroPresupuesto.Text = ""
    CmdConfirmar.Visible = TxtNroPresupuesto.Text = ""
    CmdCambiar.Visible = TxtNroPresupuesto.Text <> ""
    Call LimpiarPresupuesto
    LBAnulada.Visible = False
    CmdCambiar.Enabled = True
    CalPeriodo.Enabled = True
    CalFecha.Enabled = True
    PeriodoCerrado = False
    CentroEmisorActual = CentroEmisor
End Sub

Private Sub LVListado_DblClick()
    If LvListado.SelectedItem.Index = LvListado.ListItems.Count Then
        Exit Sub
    End If
    If VecPresupuesto(LvListado.SelectedItem.Index).O_CuentaContable = "5121" Then
        A01_4110.JerarquiaCentro = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia
        A01_4110.TxtNroPresupuesto = TxtNroPresupuesto
        A01_4110.TxtCentroDeCostoEmisor = CmbCentroDeCostoEmisor.Text
        A01_4110.TxtTotalPres = TxtMonto
        A01_4110.CalPeriodo = CalPeriodo
        A01_4110.FrameAsig.Enabled = Not LvListado.SelectedItem.ForeColor = vbBlue
        A01_4110.CmpAceptar.Enabled = Not LvListado.SelectedItem.ForeColor = vbBlue
        A01_4110.TxtCuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Descripcion
        A01_4110.Show vbModal
    End If
    
    If VecPresupuesto(LvListado.SelectedItem.Index).O_CuentaContable = "5023" Then
        A01_4120.TxtCentroDeCostoEmisor = CmbCentroDeCostoEmisor.Text
        A01_4120.CalPeriodo = CalPeriodo
        A01_4120.TxtCuentaContable = VecPresupuesto(LvListado.SelectedItem.Index).Cta_Descripcion
        A01_4120.TxtNroPresupuesto = TxtNroPresupuesto
        A01_4120.Show vbModal
    End If

End Sub

Private Sub OptMulti_Click(Index As Integer)
    
    If Index = 0 Then
       Height = 7000
       
       FrameMulti.Enabled = False
       CmdConfirmar.Enabled = True
    Else
       Height = 9195
       CmbFrecuencia.ListIndex = 0
       Call TxtCantOrdenes_Change
       'LVListado.ListItems.Clear
       'LVListado.ListItems.Add
       'ReDim VecPresupuesto(0)
       FrameMulti.Enabled = True
       CmdConfirmar.Enabled = False
    End If
    Call CentrarFormulario(Me)

End Sub

Private Sub Timer2_Timer()
   If NroPresupuesto <> 0 Then
        Call CopiarPresupuesto(NroPresupuesto)
        Timer2.Enabled = False
   End If

End Sub

Private Sub Form_Load()
    CentroEmisorActual = CentroEmisor
    ReDim VecDistribucionPresupuesto(0)
    ReDim VecPresServicioDeBar(0)
    Call CrearEncabezados
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    TxtResp.Text = NombreUsuario
    Nivel = TraerNivel("A014100")
    
    TxtNroPresupuesto.Text = ""
    
    CalFecha.Value = Date
    CalPeriodo.Value = ValidarPeriodo(Date, False)
        
    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
       
    ReDim VecPresupuesto(0)
    
    Modificado = False
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

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
Dim i As Integer
    
 If ValidarCargaPresupuesto Then
    Modificado = True
    
     'agrega al vector
      i = LvListado.ListItems.Count
      ReDim Preserve VecPresupuesto(UBound(VecPresupuesto) + 1)
      
      VecPresupuesto(i).Cta_Descripcion = CmbCuentas.Text & " - Cód. " & VecCuentasContables(CmbCuentas.ListIndex).Codigo
      VecPresupuesto(i).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo
      VecPresupuesto(i).Monto = Val(TxtMonto.Text)
      VecPresupuesto(i).P_ObservacionesPresupuesto = TxtObservaciones
      VecPresupuesto(i).Borrable = True
      
      LvListado.ListItems(i).Text = CmbCuentas.Text & " - Cód. " & VecCuentasContables(CmbCuentas.ListIndex).Codigo
      LvListado.ListItems(i).SubItems(1) = Replace(Val(TxtMonto.Text), ",", ".")
      LvListado.ListItems(i).SubItems(2) = TxtObservaciones
      LvListado.ListItems.Add
      LvListado.ListItems(LvListado.ListItems.Count).Selected = True
      Call LvListado_ItemClick(LvListado.SelectedItem)
  End If
  'calcula el total
  Call CalcularTotal
  'le da el foco al combo de cuentas
  CmbCuentas.SetFocus
Errores:
  Call ManipularError(Err.Number, Err.Description, Timer1)

End Sub

Private Function ValidarCargaPresupuesto() As Boolean
    ValidarCargaPresupuesto = True
    Dim i As Integer
    
    If CmbCuentas.ListIndex = 0 Then
        MsgBox "Debe Seleccionar una Cuenta Contable"
        CmbCuentas.SetFocus
        ValidarCargaPresupuesto = False
        Exit Function
    End If
    
    For i = 1 To UBound(VecPresupuesto)
      If i <> LvListado.SelectedItem.Index Then
       If VecPresupuesto(i).O_CuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
          MsgBox "Ya existe esa Cuenta en esta presupuesto"
          ValidarCargaPresupuesto = False
          Exit Function
       End If
      End If
    Next
    
    If Val(TxtMonto.Text) = 0 Then
        MsgBox "Debe ingresar un Monto mayor que 0"
        TxtMonto.SetFocus
        ValidarCargaPresupuesto = False
        Exit Function
    End If
    
    If VecCuentasContables(CmbCuentas.ListIndex).Codigo = "5121" Then
        A01_4110.JerarquiaCentro = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia
        A01_4110.TxtCentroDeCostoEmisor = CmbCentroDeCostoEmisor.Text
        A01_4110.CalPeriodo = CalPeriodo
        A01_4110.TxtNroPresupuesto = TxtNroPresupuesto
        A01_4110.TxtTotalPres = TxtMonto
        A01_4110.TxtCuentaContable = VecCuentasContables(CmbCuentas.ListIndex).Descripcion
        A01_4110.Show vbModal
        ValidarCargaPresupuesto = A01_4110.Ok
    End If
End Function

Private Sub CrearEncabezados()
    LvListado.ColumnHeaders.Add , , "Cuenta Contable", (LvListado.Width - 1600) / 2
    LvListado.ColumnHeaders.Add , , "Monto Pres.", 1300, 1
    LvListado.ColumnHeaders.Add , , "Observaciones", (LvListado.Width - 1600) / 2
    LvMeses.ColumnHeaders.Add , , "Período", LvMeses.Width - 300
End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To LvListado.ListItems.Count
        Total = Total + Val(Replace(LvListado.ListItems(i).SubItems(1), ",", "."))
    Next
       
    TxtTotal.Text = Format(Total, "$ 0.00##")

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
            If OptMulti(0).Value Then
                If NroPresupuesto = 0 Then
                    Call GrabarPresupuesto
                Else
                    Call ModificarPresupuesto
                End If
            Else
                If Not CrearMultiPresupuestos Then
                    Cancel = 1
                End If
            End If
         End If
         
       End If
    End If
End Sub

Private Sub CargarEnModificar(Index As Integer)
    Call UbicarCuentaContable(VecPresupuesto(Index).O_CuentaContable, CmbCuentas)
    TxtMonto.Text = Replace(VecPresupuesto(Index).Monto, ",", ".")
    TxtObservaciones.Text = VecPresupuesto(Index).P_ObservacionesPresupuesto
End Sub

Private Sub Timer1_Timer()
   If NroPresupuesto <> 0 Then
      TxtNroPresupuesto.Text = CStr(NroPresupuesto)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Errores
'NO SE TOCA
   If Item.Index < LvListado.ListItems.Count Then
        Call HabilitarAsignacion(VecPresupuesto(Item.Index).Borrable)
        Call CargarEnModificar(Item.Index)
        
        If VecPresupuesto(Item.Index).Borrable Then
            CmdModif.Enabled = True
            CmdEliminar.Enabled = True
            CmdAgregar.Enabled = False
        End If
    Else
        Call HabilitarAsignacion(True)
        CmdAgregar.Enabled = True
        CmdModif.Enabled = False
        CmdEliminar.Enabled = False
        CmbCuentas.ListIndex = 0
        TxtMonto.Text = ""
        TxtObservaciones.Text = ""
   End If
      FrameAsig.Enabled = Not Anulado
      
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub HabilitarAsignacion(Habilitar As Boolean)
    LbCant.Enabled = Habilitar
    LbCta.Enabled = Habilitar
    CmbCuentas.Enabled = Habilitar
    '
    CmdModif.Enabled = Habilitar
    CmdEliminar.Enabled = Habilitar
    CmdAgregar.Enabled = Habilitar
    TxtObservaciones.Enabled = Habilitar
    If LvListado.SelectedItem.Index <= UBound(VecPresupuesto) Then
        TxtMonto.Enabled = VecPresupuesto(LvListado.SelectedItem.Index).O_CuentaContable <> "5023" And Habilitar
    Else
        TxtMonto.Enabled = Habilitar
    End If
End Sub

Private Sub LimpiarPresupuesto()
    TxtMonto.Text = ""
    TxtTotal = "0"
    TxtObs.Text = ""
    TxtResp.Text = NombreUsuario
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    
    ReDim VecPresupuesto(0)
    ReDim VecDistribucionPresupuesto(0)
    ReDim VecPresServicioDeBar(0)

    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCuentas.ListIndex = 0

    CalFecha.Value = Date
    CalPeriodo.Value = ValidarPeriodo(Date, False)

    CmdConfirmar.Visible = True
    CmdCambiar.Visible = False
    CmdImprimir.Enabled = False
    CmdAnular.Visible = False
    LBPerCerrado.Visible = False
    PeriodoCerrado = False
    CalPeriodo.Enabled = True
    
    'muestra los option button para hacer pultiples ordenes
    OptMulti(0).Visible = True
    OptMulti(1).Visible = True
    OptMulti(0).Value = True

End Sub

Private Sub TxtCantOrdenes_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
    End If

End Sub

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call UbicarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtMonto, KeyAscii)
End Sub

Private Sub TxtNroPresupuesto_KeyPress(KeyAscii As Integer)
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

Private Sub TxtNroPresupuesto_LostFocus()
  If Val(TxtNroPresupuesto.Text) <> NroPresupuesto Then
    CmdConfirmar.Visible = TxtNroPresupuesto.Text = ""
    CmdCambiar.Visible = TxtNroPresupuesto.Text <> ""
    Call LimpiarPresupuesto
  End If
End Sub

Private Sub CargarPresupuesto(NroPresupuesto As Integer)
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim ModificarPeriodo As Boolean
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Dim RsValidarPeriodo As ADODB.Recordset
    Set RsValidarPeriodo = New ADODB.Recordset
On Error GoTo ErrorCarga
    LBAnulada.Visible = False
    LBPerCerrado.Visible = False
    Anulado = False
    
    ModificarPeriodo = True
'oculta los option button para hacer pultiples ordenes
    OptMulti(0).Visible = False
    OptMulti(1).Visible = False
    OptMulti(0).Value = True

  With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    Sql = "SpOCPresupuestosCabeceraTraerNro @NroPresupuesto= " & NroPresupuesto & _
            ", @Usuario='" & Usuario & "', @P_CentroDeCostoEmisor ='" & CentroEmisorActual & "'"
            
    .Open Sql, Conec
    
      If .EOF Then
          MsgBox "No existe un Presupuesto con esa numeración", vbInformation
          Call CmdNueva_Click
          Exit Sub
      Else
          Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & !P_Periodo & "'"
          RsValidarPeriodo.Open Sql, Conec
          PeriodoCerrado = RsValidarPeriodo!Cerrado > 0
        If Not IsNull(!P_FechaAnulacion) Or PeriodoCerrado Then
            If Not IsNull(!P_FechaAnulacion) Then
                LBAnulada.Caption = "Anulada " + Mid(CStr(!P_FechaAnulacion), 1, 10)
                LBAnulada.Visible = True
                Anulado = True
            End If
            
            LBPerCerrado.Visible = PeriodoCerrado
            ModificarPeriodo = Not PeriodoCerrado
            
            CmdCambiar.Enabled = False
            CalFecha.Enabled = False
            CalPeriodo.Enabled = False
            CmdAnular.Visible = False
        Else
            LBAnulada.Visible = False
            CmdCambiar.Enabled = True
            CalPeriodo.Enabled = True
            CalFecha.Enabled = True
            CmdAnular.Visible = True
        End If
    End If
    
    TxtNroPresupuesto.Text = Format(!P_NumeroPresupuesto, "0000000000")
    Me.NroPresupuesto = !P_NumeroPresupuesto
    CalPeriodo.Enabled = False

    CmdConfirmar.Visible = False
    CmdCambiar.Visible = True
    CmdImprimir.Enabled = True
        
    CalFecha.Value = !P_FechaEmision
    TxtResp = RsCargar!P_Responsable
    TxtObs.Text = VerificarNulo(RsCargar!P_Observaciones)
    CalPeriodo.Value = !P_Periodo
    Call BuscarCentroEmisor(!P_CentroDeCostoEmisor, CmbCentroDeCostoEmisor)
    
    .Close
    Sql = "SpOCPresupuestosRenglonesTraer2 @NroPresupuesto=" & NroPresupuesto & _
                                        ", @P_CentroDeCostoEmisor ='" & CentroEmisorActual & "'"
    .Open Sql, Conec
    i = 1
    LvListado.ListItems.Clear
    ReDim VecPresupuesto(.RecordCount)
    
      While Not .EOF
          LvListado.ListItems.Add
          VecPresupuesto(i).O_CuentaContable = !P_CuentaContable
          VecPresupuesto(i).Cta_Descripcion = BuscarDescCta(!P_CuentaContable)
          VecPresupuesto(i).Borrable = IsNull(!P_FechaAprobacion)
          VecPresupuesto(i).Monto = !Monto
          VecPresupuesto(i).FechaAprobacion = IIf(IsNull(!P_FechaAprobacion), "NULL", FechaSQL(VerificarNulo(Format(!P_FechaAprobacion, "dd/MM/YYYY")), "SQL"))
          VecPresupuesto(i).P_Observaciones = VerificarNulo(!P_Observaciones)
          VecPresupuesto(i).P_ObservacionesPresupuesto = VerificarNulo(!P_ObservacionesPresupuesto)
          LvListado.ListItems(i).Text = VecPresupuesto(i).Cta_Descripcion & " - Cód. " & VecPresupuesto(i).O_CuentaContable
          LvListado.ListItems(i).SubItems(1) = VecPresupuesto(i).Monto
          LvListado.ListItems(i).SubItems(2) = VecPresupuesto(i).P_ObservacionesPresupuesto
          
          If Not VecPresupuesto(i).Borrable Then
             LvListado.ListItems(i).ForeColor = vbBlue
             LvListado.ListItems(i).ListSubItems(1).ForeColor = vbBlue
          End If
          i = i + 1
          .MoveNext
      Wend
      .Close
      Sql = "SpOcPresupuestosDistrubucionTraer @P_NumeroPresupuesto =" & NroPresupuesto & _
                                            ", @P_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
     .Open Sql, Conec
     ReDim VecDistribucionPresupuesto(.RecordCount)
     For i = 1 To .RecordCount
        VecDistribucionPresupuesto(i).P_CentroDeCostosEmisor = !P_CentroDeCostosEmisor
        VecDistribucionPresupuesto(i).P_CuentaContable = !P_CuentaContable
        VecDistribucionPresupuesto(i).P_Importe = !P_Importe
        VecDistribucionPresupuesto(i).P_NumeroPresupuesto = !P_NumeroPresupuesto
        VecDistribucionPresupuesto(i).P_Periodo = !P_Periodo
        VecDistribucionPresupuesto(i).P_SubCentroDeCosto = !P_SubCentroDeCosto
        .MoveNext
     Next
     
      .Close
      
      Sql = "SpOcPresupuestosServicioBarTraer @D_NumeroPresupuesto =" & NroPresupuesto & _
                                           ", @D_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
     .Open Sql, Conec
     ReDim VecPresServicioDeBar(.RecordCount)
     For i = 1 To .RecordCount
        VecPresServicioDeBar(i).D_Cantidad = !D_Cantidad
        VecPresServicioDeBar(i).D_CentroDeCostosEmisor = !D_CentroDeCostosEmisor
        VecPresServicioDeBar(i).D_Cuenta = !D_Cuenta
        VecPresServicioDeBar(i).D_Linea = !D_Linea
        VecPresServicioDeBar(i).D_Periodo = !D_Periodo
        VecPresServicioDeBar(i).D_NumeroPresupuesto = !D_NumeroPresupuesto
        VecPresServicioDeBar(i).D_PrecioUnitario = !D_PrecioUnitario
        VecPresServicioDeBar(i).D_Proveedor = !D_Proveedor
        .MoveNext
     Next
             
End With

    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
      
  Call CalcularTotal

  CalPeriodo.Enabled = False
  
ErrorCarga:
  Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CopiarPresupuesto(NroPresupuesto As Integer)
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim ModificarPeriodo As Boolean
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Dim RsValidarPeriodo As ADODB.Recordset
    Set RsValidarPeriodo = New ADODB.Recordset

    LBAnulada.Visible = False
    LBPerCerrado.Visible = False
    Anulado = False
    
    ModificarPeriodo = True

  With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    Sql = "SpOCPresupuestosCabeceraTraerNro @NroPresupuesto= " & NroPresupuesto & _
            ", @Usuario='" & Usuario & "', @P_CentroDeCostoEmisor ='" & CentroEmisorActual & "'"
            
    .Open Sql, Conec
    
      If .EOF Then
          MsgBox "No existe un Presupuesto con esa numeración", vbInformation
          Call CmdNueva_Click
          Exit Sub
    End If
    
    TxtNroPresupuesto.Text = Format(!P_NumeroPresupuesto, "0000000000")
    Me.NroPresupuesto = !P_NumeroPresupuesto
    
        
    CalFecha.Value = !P_FechaEmision
    TxtResp = RsCargar!P_Responsable
    TxtObs.Text = VerificarNulo(RsCargar!P_Observaciones)
    CalPeriodo.Value = !P_Periodo
    Call BuscarCentroEmisor(!P_CentroDeCostoEmisor, CmbCentroDeCostoEmisor)
    
    .Close
    Sql = "SpOCPresupuestosRenglonesTraer2 @NroPresupuesto=" & NroPresupuesto & _
                                        ", @P_CentroDeCostoEmisor ='" & CentroEmisorActual & "'"
    .Open Sql, Conec
    i = 1
    LvListado.ListItems.Clear
    ReDim VecPresupuesto(.RecordCount)
    
      While Not .EOF
          LvListado.ListItems.Add
          VecPresupuesto(i).O_CuentaContable = !P_CuentaContable
          VecPresupuesto(i).Cta_Descripcion = BuscarDescCta(!P_CuentaContable)
          VecPresupuesto(i).Borrable = True
          VecPresupuesto(i).Monto = !Monto
          VecPresupuesto(i).FechaAprobacion = "NULL"
          
          LvListado.ListItems(i).Text = VecPresupuesto(i).Cta_Descripcion & " - Cód. " & VecPresupuesto(i).O_CuentaContable
          LvListado.ListItems(i).SubItems(1) = VecPresupuesto(i).Monto
          
          i = i + 1
          .MoveNext
      Wend
      
      'ditribbición de movilidades
      .Close
      Sql = "SpOcPresupuestosDistrubucionTraer @P_NumeroPresupuesto =" & NroPresupuesto & _
                                            ", @P_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
     .Open Sql, Conec
     ReDim VecDistribucionPresupuesto(.RecordCount)
     For i = 1 To .RecordCount
        VecDistribucionPresupuesto(i).P_CentroDeCostosEmisor = !P_CentroDeCostosEmisor
        VecDistribucionPresupuesto(i).P_CuentaContable = !P_CuentaContable
        VecDistribucionPresupuesto(i).P_Importe = !P_Importe
        VecDistribucionPresupuesto(i).P_NumeroPresupuesto = !P_NumeroPresupuesto
        VecDistribucionPresupuesto(i).P_Periodo = !P_Periodo
        VecDistribucionPresupuesto(i).P_SubCentroDeCosto = !P_SubCentroDeCosto
        .MoveNext
     Next
     
      .Close
      'distribucion servicio de Bar
      Sql = "SpOcPresupuestosServicioBarTraer @D_NumeroPresupuesto =" & NroPresupuesto & _
                                           ", @D_CentroDeCostosEmisor ='" & CentroEmisorActual & "'"
     .Open Sql, Conec
     ReDim VecPresServicioDeBar(.RecordCount)
     For i = 1 To .RecordCount
        VecPresServicioDeBar(i).D_Cantidad = !D_Cantidad
        VecPresServicioDeBar(i).D_CentroDeCostosEmisor = !D_CentroDeCostosEmisor
        VecPresServicioDeBar(i).D_Cuenta = !D_Cuenta
        VecPresServicioDeBar(i).D_Linea = !D_Linea
        VecPresServicioDeBar(i).D_Periodo = !D_Periodo
        VecPresServicioDeBar(i).D_NumeroPresupuesto = !D_NumeroPresupuesto
        VecPresServicioDeBar(i).D_PrecioUnitario = !D_PrecioUnitario
        VecPresServicioDeBar(i).D_Proveedor = !D_Proveedor
        .MoveNext
     Next
            
End With

    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
      
  Call CalcularTotal

  'CalPeriodo.Enabled = ModificarPeriodo
  TxtNroPresupuesto.Text = ""
    CalFecha.Value = Date
    CalPeriodo.Value = ValidarPeriodo(Date)

    Timer2.Enabled = False
    CmdCambiar.Visible = False
    CmdCambiar.Enabled = False
    CmdConfirmar.Enabled = True
    CmdConfirmar.Visible = True
    CmdImprimir.Enabled = False
    CmdAnular.Visible = False
    LBAnulada.Visible = False
    LBPerCerrado.Visible = False
    CalPeriodo.Enabled = True
End Sub

'funciones que eran púplical
Private Sub CargarVecCentroEmisor(CentroEmisor As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
      ReDim VecCuentasContables(0)
      'en esta sección carga las cuentas
  With RsCargar
      Sql = "SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CentroEmisor & "'"
      .Open Sql, Conec, adOpenKeyset, adLockOptimistic
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

Public Function BuscarDescArt(A_Codigo As Long) As String
    Dim i As Integer
    For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = A_Codigo Then
            BuscarDescArt = VecArtCompra(i).A_Descripcion
            Exit Function
        End If
    Next
End Function

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
