VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5711 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acumulado Por Proveedor"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3390
      TabIndex        =   5
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmpExpPDF 
      Caption         =   "Exportar &PDF"
      Height          =   350
      Left            =   2055
      TabIndex        =   4
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   705
      TabIndex        =   3
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4740
      TabIndex        =   0
      Top             =   6795
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6420
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   6585
      _ExtentX        =   11615
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
      Left            =   6120
      TabIndex        =   2
      Top             =   6570
      Width           =   510
   End
End
Attribute VB_Name = "A01_5711"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cuenta As String
Public Periodo As Date
Public CentroEmisor As String

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
    Sql = "SpOcConsultaDetallePorCuentaContableSGPAcumularProv @Periodo =" & FechaSQL(CStr(Periodo), "SQL") & _
                                                           " , @CuentaContable ='" & Cuenta & _
                                                           "', @Emisor='" & CentroEmisor & "'"
    With RsCargar
        .Open Sql, Conec
        
      LvListado.Sorted = False
        While Not .EOF
            i = i + 1
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = !R_CodigoProveedor
            LvListado.ListItems(i).SubItems(1) = BuscarDescProv(!R_CodigoProveedor)
            LvListado.ListItems(i).SubItems(2) = Format(!Importe, "0.00")
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
     
  If MenuEmisionOrdenCompra.Cuadros.FileName <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5721.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.FileName = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.FileName)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export ListA01_5721.Pages
         Unload ListA01_5721
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
    LvListado.ColumnHeaders.Add , , "Cod. Prov.", 1000
    LvListado.ColumnHeaders.Add , , "Proveedor", LvListado.Width - 2350
    LvListado.ColumnHeaders.Add , , "Importe", 1100, 1
    LvListado.ColumnHeaders.Add , , "CodProv", 0, 1
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub CmdExpExcel_Click()
    Dialogo.FileName = ""
    Call ArmarExcel(Dialogo)
    If Dialogo.FileName <> "" Then
        MousePointer = vbHourglass
        Call GenerarPlanilla(Dialogo.FileName, Dialogo.FilterIndex)
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
        .Range("C" & 7 + LvListado.ListItems.Count).Select
        .ActiveCell.Formula = "=Sum($C7:$C" & 6 + LvListado.ListItems.Count & ")"
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
   ListA01_5721.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
   
    RsListado.Fields.Append "CodProv", adVarChar, 10
    RsListado.Fields.Append "Proveedor", adVarChar, 150
    RsListado.Fields.Append "Importe", adDouble
       
    RsListado.Open
    i = 1
    While i <= LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
           RsListado!CodProv = .Text
           RsListado!Proveedor = .SubItems(1)
           RsListado!Importe = ValN(.SubItems(2))
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5721.TxtCuentaCostable.Text = BuscarDescCta(Cuenta)
    ListA01_5721.TxtCentro.Text = BuscarDescCentroEmisor(CentroEmisor)
    ListA01_5721.TxtPeriodo.Text = Format(Periodo, "MMMM/yyyy")
    ListA01_5721.DataControl1.Recordset = RsListado
    ListA01_5721.Zoom = -1
End Sub

