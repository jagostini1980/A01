VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_5911 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Contable por Cuenta Contable"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   350
      Left            =   4020
      TabIndex        =   4
      Top             =   8190
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   0
      Top             =   8100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   5370
      TabIndex        =   3
      Top             =   8190
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   6735
      TabIndex        =   0
      Top             =   8190
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   7590
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   13388
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
   Begin VB.Label LbTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   11895
   End
   Begin VB.Label LbTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   11430
      TabIndex        =   2
      Top             =   8235
      Width           =   510
   End
End
Attribute VB_Name = "A01_5911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cuenta As String
Public Periodo As Date

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CargarListado()
Dim Sql As String
Dim i As Integer
Dim Importe As Double
Dim PeriodoSql As String
Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LVListado.ListItems.Clear
    MousePointer = vbHourglass
    'Realiza la consulta
    PeriodoSql = "01/" & Format(Month(Periodo), "00") & "/" & Year(Periodo)
    Sql = "SpOcConsultaPresupuestoFinancieroDetalleContableXCuenta @Periodo = " & FechaSQL(PeriodoSql, "SQL") & _
                                                               ", @Cuenta = '" & Cuenta & "'"
                                        
    With RsCargar
        .Open Sql, Conec
        
      LVListado.Sorted = False
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add

            LVListado.ListItems(i).Text = !C_Empresa
            LVListado.ListItems(i).SubItems(1) = Format(!C_Fecha, "dd/MM/yyyy")
            LVListado.ListItems(i).SubItems(2) = !C_Concepto
            LVListado.ListItems(i).SubItems(3) = BuscarDescCentroPorCodSecundario(VerificarNulo(!CentroDeCosto))
            LVListado.ListItems(i).SubItems(4) = BuscarDescCentroEmisorPorJerarquia(VerificarNulo(!C_Emisor))
            LVListado.ListItems(i).SubItems(5) = Format(!C_Importe, "#,##0.00")

            Importe = Importe + !C_Importe
            .MoveNext
        Wend
    End With
    LVListado.Sorted = True
    LbTotal.Caption = "Total: " & Format(Importe, "#,##0.00")
Error:
    MousePointer = vbNormal
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdExpExcel_Click()
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
        Call EncabezadoExcel(ex, LVListado, Caption & " - " & LbTitulo.Caption, 6)
        Call DatosExcel(ex, LVListado, 6)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To LVListado.ColumnHeaders.Count
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(Periodo, "MMM/yyyy")
        .Range("F" & 7 + LVListado.ListItems.Count).Select
        .ActiveCell.Formula = "=Sum($F7:$F" & 6 + LVListado.ListItems.Count & ")"
        .Range("A" & 7 + LVListado.ListItems.Count).Select
        .ActiveCell.FormulaR1C1 = "Total ==>"

        ColorFondo = &HC0E0FF
        Call FormatearExcelConTotal(ex, LVListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5911.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "Empresa", adVarChar, 5
    RsListado.Fields.Append "Fecha", adVarChar, 20
    RsListado.Fields.Append "Detalle", adVarChar, 500
    RsListado.Fields.Append "CentroEmisor", adVarChar, 100
    RsListado.Fields.Append "SubCentro", adVarChar, 100
    RsListado.Fields.Append "Importe", adVarChar, 100
    RsListado.Open
    i = 1
    While i <= LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
           RsListado!Empresa = .Text
           RsListado!Fecha = .SubItems(1)
           RsListado!Detalle = .SubItems(2)
           RsListado!CentroEmisor = .SubItems(3)
           RsListado!SubCentro = .SubItems(4)
           RsListado!Importe = .SubItems(5)

      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5911.TxtCuentaCostable = LbTitulo.Caption
    ListA01_5911.TxtPeriodo.Text = Format(Periodo, "MMMM/yyyy")
    ListA01_5911.DataControl1.Recordset = RsListado
    ListA01_5911.Zoom = -1
End Sub

Private Sub CmpExpPDF_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5730.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export ListA01_5730.Pages
         Unload ListA01_5730
  End If
Error:
    If Err.Number = 0 Then
        MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarListado
End Sub

Private Sub CrearEncabezado()
    LVListado.ColumnHeaders.Add , , "Empresa", 1000
    LVListado.ColumnHeaders.Add , , "Fecha", 1100
    LVListado.ColumnHeaders.Add , , "Concepto", LVListado.Width - 7350
    LVListado.ColumnHeaders.Add , , "Centro De Costos", 2000
    LVListado.ColumnHeaders.Add , , "Centro Emisor", 2000
    LVListado.ColumnHeaders.Add , , "Importe", 1000, 1
    LVListado.ColumnHeaders.Add , , "CodCentroEmisor", 0, 1
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LVListado.SortKey = ColumnHeader.Index - 1
End Sub

