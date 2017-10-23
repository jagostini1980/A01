VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form BuscarAutorizacionDePagoEspecial 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Autorización de Pago Especial"
   ClientHeight    =   6915
   ClientLeft      =   3225
   ClientTop       =   2250
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   38
      TabIndex        =   5
      Top             =   45
      Width           =   11655
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   10395
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   0
         Top             =   217
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53739521
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   330
         Left            =   3915
         TabIndex        =   1
         Top             =   217
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53739521
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   7380
         TabIndex        =   8
         Top             =   210
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label LbCentroEmisor 
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
         Left            =   5265
         TabIndex        =   9
         Top             =   285
         Width           =   2055
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
         Left            =   135
         TabIndex        =   7
         Top             =   285
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
         Left            =   2745
         TabIndex        =   6
         Top             =   285
         Width           =   1155
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   225
      Top             =   6390
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5138
      TabIndex        =   4
      Top             =   6435
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11385
      Top             =   7605
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   5565
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9816
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
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "BuscarAutorizacionDePagoEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean

Private Sub CargarLV(FechaDesde As Date, FechaHasta As Date)
   Dim Sql As String
   Dim RsListado As ADODB.Recordset
   Set RsListado = New ADODB.Recordset
   Dim i As Integer
   
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   
   Sql = "SpOCAutorizacionesDePagoCabeceraTraer @FechaDesde =" & FechaSQL(CStr(FechaDesde), "SQL") & _
                                             ", @FechaHasta =" & FechaSQL(CStr(FechaHasta), "SQL") & _
                                             ", @CentroDeCosto='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
   
   RsListado.Open Sql, Conec
   With RsListado
   'limpia el LV
   LvListado.ListItems.Clear
   LvListado.Sorted = False
   If .RecordCount > 0 Then
      For i = 1 To .RecordCount
        LvListado.ListItems.Add
                                                          
        LvListado.ListItems(i).Text = Format(!A_NumeroDeAutorizacionDePago, "0000000000")
        LvListado.ListItems(i).SubItems(1) = !A_Fecha
        LvListado.ListItems(i).SubItems(2) = BuscarDescProv(!A_CodigoProveedor)
        LvListado.ListItems(i).SubItems(3) = IIf(IsNull(!A_FechaAnulacion), "", Format(!A_FechaAnulacion, "dd/MMM/yyyy HH:mm"))
        LvListado.ListItems(i).SubItems(4) = !A_Observaciones
        .MoveNext
      Next
    End If
   End With
    LvListado.Sorted = True
   Set RsListado = Nothing
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Nº de Autorización", 1500
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Proveedor", 2750
    LvListado.ColumnHeaders.Add , , "F. de Anulación", 1400
    LvListado.ColumnHeaders.Add , , "Observaciones", LvListado.Width - 7000
End Sub

Private Sub CalFechaDesde_Change()
'con esto controlo que el fecha hasta no sea menos que fecha desde
    CalFechaHasta.MinDate = CalFechaDesde.Value
End Sub

Private Sub CalFechaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    'si se preciona enter se carga el liste view
        Call CmdTraer_Click
    End If

End Sub

Private Sub CalFechaHasta_Change()
'con esto controlo que fecha desde no sea mayor que fecha hasta
    CalFechaDesde.MaxDate = CalFechaHasta.Value
End Sub

Private Sub CalFechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
   'si se preciona enter se carga el liste view
        Call CmdTraer_Click
    End If

End Sub

Private Sub CmbLineas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Beep
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV(CalFechaDesde.Value, CalFechaHasta.Value)
End Sub

Private Sub Form_Load()
    CalFechaDesde.Value = DateAdd("m", -1, Date)
    CalFechaDesde.MaxDate = Date
    CalFechaHasta.Value = Date
    CalFechaHasta.MinDate = Date
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    CmbCentroDeCostoEmisor.ListIndex = 0
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A014400") = 2

    Call CrearEncabezado
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub LVListado_DblClick()
    A01_4400.NroAutorizacion = Val(LvListado.SelectedItem.Text)
    Unload Me
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub


