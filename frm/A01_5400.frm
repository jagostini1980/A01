VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5400 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Recepción de mercaderías Pendientes"
   ClientHeight    =   7485
   ClientLeft      =   4590
   ClientTop       =   2220
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   210
      Top             =   6885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Height          =   400
      Left            =   3447
      TabIndex        =   8
      Top             =   6975
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   12135
      Begin VB.CheckBox ChkPEndientes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos los Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2175
         TabIndex        =   7
         Top             =   210
         Width           =   2235
      End
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H80000003&
         Caption         =   "Traer"
         Height          =   315
         Left            =   4455
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   1035
         TabIndex        =   5
         Top             =   172
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   24117251
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   400
      Left            =   5295
      TabIndex        =   2
      Top             =   6975
      Width           =   1635
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerra&r"
      Height          =   400
      Left            =   7143
      TabIndex        =   1
      Top             =   6975
      Width           =   1635
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6165
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10874
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
Attribute VB_Name = "A01_5400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FechaDesde As String
Public FechaHasta As String
Public Parada As String
Public Pregunta As String

Private Sub ChkPEndientes_Click()
    CalPeriodo.Enabled = ChkPEndientes.Value = 0
End Sub

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub CrearEncabezado()

    LvListado.ColumnHeaders.Add , , "Centro de Costos", 100 + (LvListado.Width - 6000) / 3
    LvListado.ColumnHeaders.Add , , "Proveedor", 100 + (LvListado.Width - 6000) / 3
    LvListado.ColumnHeaders.Add , , "Orden Nº", 1000, 1
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Artículo", (LvListado.Width - 6000) / 3
    LvListado.ColumnHeaders.Add , , "Cant. Pedida", 1100, 1
    LvListado.ColumnHeaders.Add , , "Cant. Pendiente", 1350, 1
    LvListado.ColumnHeaders.Add , , "Usuario", 1000
    LvListado.ColumnHeaders.Add , , "", 0
End Sub

Private Sub CmdExp_Click()
    Dialogo.Filename = ""
    Call ArmarExcel(Dialogo)
    If Dialogo.Filename <> "" Then
        MousePointer = vbHourglass
        Call GenerarPlanilla(Dialogo.Filename, Dialogo.FilterIndex)
        MousePointer = vbNormal
    End If

End Sub

Private Sub GenerarPlanilla(NombreArchivo As String, Filtro As Integer)
Dim ex As Excel.Application
Dim col As Integer
Dim ColorFondo As Long

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcel(ex, LvListado, Caption, 6)
        Call DatosExcel(ex, LvListado, 6)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To LvListado.ColumnHeaders.Count
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(CalPeriodo, "MM/yyyy")

        ColorFondo = &HC0E0FF
        Call FormatearExcel(ex, LvListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeConsulta
    ListA01_5400.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "CentroDeCosto", adVarChar, 100
    RsListado.Fields.Append "NroOrden", adVarChar, 10
    RsListado.Fields.Append "Fecha", adVarChar, 10
    RsListado.Fields.Append "Proveedor", adVarChar, 100
    RsListado.Fields.Append "Articulo", adVarChar, 100
    RsListado.Fields.Append "Pedida", adInteger
    RsListado.Fields.Append "Pendiente", adInteger
    
    RsListado.Open
    i = 1
    While i <= LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
            RsListado!CentroDeCosto = .Text
            RsListado!Proveedor = .SubItems(1)
            RsListado!NroOrden = .SubItems(2)
            RsListado!Fecha = .SubItems(3)
            RsListado!Articulo = .SubItems(4)
            RsListado!Pedida = .SubItems(5)
            RsListado!Pendiente = .SubItems(6)
      End With
        i = i + 1
    Wend
    RsListado.MoveFirst
    ListA01_5400.LbPeriodo = "Período: " & Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5400.DataControl1.Recordset = RsListado
    ListA01_5400.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV
End Sub

Private Sub Form_Load()
   Call CrearEncabezado
   CalPeriodo.Value = Date
   Call CargarLV
End Sub

Private Sub CargarLV()
Dim RsCargar As New ADODB.Recordset
Dim Sql As String
Dim CentroOrden As String
Dim i As Integer
    
    With RsCargar
        If ChkPEndientes.Value = 1 Then
            Sql = "SpOcConsultaRecepcionDeMercaderiasPendientePorCentroDeCosto @CentroDeCosto='" & CentroEmisor & "'"
        Else
            Sql = "SpOcConsultaRecepcionDeMercaderiasPendientes @Periodo=" & FechaSQL(CalPeriodo.Value, "SQL")
        End If
        .Open Sql, Conec
        LvListado.ListItems.Clear
        While Not .EOF
            i = i + 1
            LvListado.ListItems.Add
                                                    
            If CentroOrden <> !C_Descripcion & !O_NumeroOrdenDeCompra Then
               LvListado.ListItems(i).Text = !C_Descripcion
               LvListado.ListItems(i).SubItems(1) = BuscarDescProv(!O_CodigoProveedor)
               LvListado.ListItems(i).SubItems(2) = Format(!O_NumeroOrdenDeCompra, "00000000")
               LvListado.ListItems(i).SubItems(3) = !O_Fecha
               CentroOrden = !C_Descripcion & !O_NumeroOrdenDeCompra
            End If
            
            LvListado.ListItems(i).SubItems(4) = BuscarDescArt(!O_CodigoArticulo, !C_TablaArticulos)
            LvListado.ListItems(i).SubItems(5) = !O_CantidadPedida
            LvListado.ListItems(i).SubItems(6) = !O_CantidadPendiente
            LvListado.ListItems(i).SubItems(7) = !U_Usuario
            .MoveNext
        Wend
    End With
End Sub

