VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5720 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Contable"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Acumular por Sub-CC"
      Height          =   350
      Left            =   1324
      TabIndex        =   7
      Top             =   6840
      Width           =   1725
   End
   Begin VB.CommandButton CmdAcumularProveedor 
      Caption         =   "Acumular por Prov."
      Height          =   350
      Left            =   3161
      TabIndex        =   6
      Top             =   6840
      Width           =   1590
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   45
      Top             =   6705
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   350
      Left            =   7546
      TabIndex        =   5
      Top             =   6840
      Width           =   1230
   End
   Begin VB.CommandButton CmpExpPDF 
      Caption         =   "Exportar &PDF"
      Height          =   350
      Left            =   6201
      TabIndex        =   4
      Top             =   6840
      Width           =   1230
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   4856
      TabIndex        =   3
      Top             =   6840
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   8891
      TabIndex        =   0
      Top             =   6840
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6420
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11324
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
      Left            =   10845
      TabIndex        =   2
      Top             =   6615
      Width           =   510
   End
End
Attribute VB_Name = "A01_5720"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cuenta As String
Public Periodo As Date
Public CentroEmisor As String

Private Sub CmdAcumularProveedor_Click()
    MousePointer = vbHourglass
    A01_5721.Cuenta = Cuenta
    A01_5721.Periodo = Periodo
    A01_5721.CentroEmisor = CentroEmisor
    A01_5721.Show vbModal
    MousePointer = vbNormal
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CargarListado()
Dim Sql As String
Dim i As Integer
Dim RsCargar As New ADODB.Recordset
Dim Importe As Double

MousePointer = vbHourglass
On Error GoTo Error
    LvListado.ListItems.Clear
    'Realiza la consulta
    Sql = "SpOcConsultaDetallePorCuentaContable @Periodo  ='" & Format(Periodo, "MM/yyyy") & _
                                        "', @CuentaContable ='" & Cuenta & _
                                        "', @Emisor='" & CentroEmisor & "'"
    With RsCargar
        .Open Sql, Conec
        
      LvListado.Sorted = False
        While Not .EOF
            i = i + 1
            LvListado.ListItems.Add
            
            LvListado.ListItems(i).Text = !Fecha
            LvListado.ListItems(i).SubItems(1) = BuscarDescCentroEmisorPorJerarquia((VerificarNulo(!C_Emisor)))
            LvListado.ListItems(i).SubItems(2) = BuscarDescCentroPorCodSecundario(Format(Trim(!CentroDeCosto), "000"))
            LvListado.ListItems(i).SubItems(3) = !C_Proveedor
            LvListado.ListItems(i).SubItems(4) = !CONCEPTO
            LvListado.ListItems(i).SubItems(5) = !C_Empresa
            LvListado.ListItems(i).SubItems(6) = Format(!Importe, "0.00")
            LvListado.ListItems(i).SubItems(7) = VerificarNulo(!C_SGP)
            Importe = Importe + !Importe
            .MoveNext
        Wend
    End With
    LvListado.Sorted = True
    LbTotal.Caption = "Total: " & Format(Importe, "0.00")

Error:
    MousePointer = vbNormal
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmpExpPDF_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5720.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export ListA01_5720.Pages
         Unload ListA01_5720
  End If
Error:
    If Err.Number = 0 Then
        MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Sub Command1_Click()
    MousePointer = vbHourglass
    A01_5722.Cuenta = Cuenta
    A01_5722.Periodo = Periodo
    A01_5722.CentroEmisor = CentroEmisor
    A01_5722.Show vbModal
    MousePointer = vbNormal

End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarListado
End Sub

Private Sub CrearEncabezado()

    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Centro Emisor", 1800
    LvListado.ColumnHeaders.Add , , "Sub-Centro de Costo", 1800
    LvListado.ColumnHeaders.Add , , "Cod. Prov", 900, 1
    LvListado.ColumnHeaders.Add , , "Detalle", (LvListado.Width - 8550)
    LvListado.ColumnHeaders.Add , , "Empresa", 900
    LvListado.ColumnHeaders.Add , , "Importe", 1100, 1
    LvListado.ColumnHeaders.Add , , "SGP", 600
    LvListado.ColumnHeaders.Add , , "CodCentro", 0, 1

End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
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
        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(Periodo, "MMM/yyyy")
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Centro de Costo: " & BuscarDescCentroEmisor(CentroEmisor)
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Cuenta Contable: " & BuscarDescCta(Cuenta)
        .Range("G" & 7 + LvListado.ListItems.Count).Select
        .ActiveCell.Formula = "=Sum($G7:$G" & 6 + LvListado.ListItems.Count & ")"
        .Range("A" & 7 + LvListado.ListItems.Count).Select
        .ActiveCell.FormulaR1C1 = "Total ==>"

        ColorFondo = &HC0E0FF
        Call FormatearExcelConTotal(ex, LvListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5720.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "Fecha", adVarChar, 10
    RsListado.Fields.Append "CentroEmisor", adVarChar, 150
    RsListado.Fields.Append "SubCentro", adVarChar, 150
    RsListado.Fields.Append "CodProv", adVarChar, 20
    RsListado.Fields.Append "Detalle", adVarChar, 500
    RsListado.Fields.Append "Empresa", adVarChar, 15
    RsListado.Fields.Append "Importe", adDouble
       
    RsListado.Open
    i = 1
    While i <= LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
           RsListado!Fecha = .Text
           RsListado!CentroEmisor = .SubItems(1)
           RsListado!SubCentro = .SubItems(2)
           RsListado!CodProv = .SubItems(3)
           RsListado!Detalle = .SubItems(4)
           RsListado!Empresa = .SubItems(5)
           RsListado!Importe = ValN(.SubItems(6))
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5720.TxtCuentaCostable.Text = BuscarDescCta(Cuenta)
    ListA01_5720.TxtCentro.Text = BuscarDescCentroEmisor(CentroEmisor)
    ListA01_5720.TxtPeriodo.Text = Format(Periodo, "MMMM/yyyy")
    ListA01_5720.DataControl1.Recordset = RsListado
    ListA01_5720.Zoom = -1
End Sub

