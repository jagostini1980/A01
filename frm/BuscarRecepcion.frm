VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarRecepcion 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Recepciones"
   ClientHeight    =   6915
   ClientLeft      =   3225
   ClientTop       =   2250
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   9765
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   5985
         TabIndex        =   2
         Top             =   263
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   330
         Left            =   1665
         TabIndex        =   0
         Top             =   255
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   330
         Left            =   4500
         TabIndex        =   1
         Top             =   255
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   38940
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
         Left            =   360
         TabIndex        =   7
         Top             =   323
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
         Left            =   3240
         TabIndex        =   6
         Top             =   323
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
      Left            =   3780
      TabIndex        =   4
      Top             =   6390
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11385
      Top             =   7605
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   5385
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
Attribute VB_Name = "BuscarRecepcion"
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
   LvListado.Sorted = False
   Sql = "SpOCRecepcionOrdenesDeCompraCabeceraTraer @FechaDesde =" + FechaSQL(CStr(FechaDesde), "SQL") + ",  @FechaHasta =" + FechaSQL(CStr(FechaHasta), "SQL") + _
            ", @Usuario='" + Usuario + "'"
   
   RsListado.Open Sql, Conec
   With RsListado
   'limpia el LV
   LvListado.ListItems.Clear
   If .RecordCount > 0 Then
      For i = 1 To .RecordCount
        LvListado.ListItems.Add
                                                    
        LvListado.ListItems(i).Text = Format(!R_NumeroRecepcion, "0000000000")
        LvListado.ListItems(i).SubItems(1) = !R_Fecha
        LvListado.ListItems(i).SubItems(2) = BuscarDescProv(!R_CodigoProveedor)
        LvListado.ListItems(i).SubItems(3) = !R_LugarDeRecepcion
        LvListado.ListItems(i).SubItems(4) = IIf(IsNull(!R_FechaAnulacion), "", !R_FechaAnulacion)
        LvListado.ListItems(i).SubItems(5) = !R_CodigoProveedor
        LvListado.ListItems(i).SubItems(6) = VerificarNulo(!R_Factura)
        .MoveNext
      Next
    End If
   End With
   LvListado.Sorted = True
   Set RsListado = Nothing
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


Private Sub CmdTraer_Click()
    Call CargarLV(CalFechaDesde.Value, CalFechaHasta.Value)
End Sub

Private Sub Form_Load()
    CalFechaDesde.Value = DateAdd("m", -1, Date)
    CalFechaDesde.MaxDate = Date
    CalFechaHasta.Value = Date
    CalFechaHasta.MinDate = Date
    
    Call CrearEncabezado
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , "Nro", "Nº de Recepcion", 1500
    LvListado.ColumnHeaders.Add , "Fecha", "Fecha", 1100
    LvListado.ColumnHeaders.Add , "Prov", "Proveedor", 2000
    LvListado.ColumnHeaders.Add , "LugarEnt", "Lugar de Entrega", 2000
    LvListado.ColumnHeaders.Add , "FAnulacion", "F. de Anulación", 1400
    LvListado.ColumnHeaders.Add , "CodProv", "C_Prov", 0
    LvListado.ColumnHeaders.Add , "Factura", "Factura", 1500
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   LvListado.SortKey = ColumnHeader.Position - 1
End Sub

Private Sub LVListado_DblClick()
    A01_3200.NroRecepcion = Val(LvListado.ListItems(LvListado.SelectedItem.Index).Text)
    Unload Me
End Sub


