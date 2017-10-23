VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5730 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle SGP Presupuestado"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7620
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
      Left            =   3510
      TabIndex        =   5
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmpExpPDF 
      Caption         =   "Exportar &PDF"
      Height          =   350
      Left            =   2190
      TabIndex        =   4
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   810
      TabIndex        =   3
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4860
      TabIndex        =   0
      Top             =   6795
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6600
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11642
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
      Left            =   7020
      TabIndex        =   2
      Top             =   6840
      Width           =   510
   End
End
Attribute VB_Name = "A01_5730"
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
Dim Importe As Double
Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LVListado.ListItems.Clear
    MousePointer = vbHourglass
    'Realiza la consulta
    Sql = "SpOcConsultaDetallePorCuentaPresupuestado @Periodo = '" & Format(Periodo, "MM/yyyy") & _
                                                 "', @CuentaContable = '" & Cuenta & _
                                                 "', @CentroEmisor ='" & CentroEmisor & "'"
    With RsCargar
        .Open Sql, Conec
        
      LVListado.Sorted = False
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add
            
            LVListado.ListItems(i).Text = BuscarDescCentroEmisor(!P_CentroDeCostoEmisor)
            LVListado.ListItems(i).SubItems(1) = Format(!P_NumeroPresupuesto, "00000000")
            LVListado.ListItems(i).SubItems(2) = Format(!Importe, "0.00")
           
            Importe = Importe + !Importe
            .MoveNext
        Wend
    End With
    LVListado.Sorted = True
    LbTotal.Caption = "Total: " & Format(Importe, "0.00")
Error:
    MousePointer = vbNormal
    Call ManipularError(Err.Number, Err.Description)
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
        Call EncabezadoExcel(ex, LVListado, Caption, 6)
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
        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(Periodo, "MMM/yyyy")
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Centro de Costo: " & BuscarDescCentroEmisor(CentroEmisor)
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Cuenta Contable: " & BuscarDescCta(Cuenta)
        .Range("C" & 7 + LVListado.ListItems.Count).Select
        .ActiveCell.Formula = "=Sum($C7:$C" & 6 + LVListado.ListItems.Count & ")"
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
   ListA01_5730.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "CentroEmisor", adVarChar, 100
    RsListado.Fields.Append "NroPres", adVarChar, 15
    RsListado.Fields.Append "Importe", adDouble
   
    
    RsListado.Open
    i = 1
    While i <= LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
           RsListado!CentroEmisor = .Text
           RsListado!NroPres = .SubItems(1)
           RsListado!Importe = ValN(.SubItems(2))
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5730.TxtCuentaCostable.Text = BuscarDescCta(Cuenta)
    ListA01_5730.TxtCentro.Text = BuscarDescCentroEmisor(CentroEmisor)
    ListA01_5730.TxtPeriodo.Text = Format(Periodo, "MMMM/yyyy")
    ListA01_5730.DataControl1.Recordset = RsListado
    ListA01_5730.Zoom = -1
End Sub

Private Sub CmpExpPDF_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.FileName <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5730.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.FileName = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.FileName)
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
    LVListado.ColumnHeaders.Add , , "Centro Emisor", (LVListado.Width - 2950)
    LVListado.ColumnHeaders.Add , , "Nº de Presupuesto", 1500, 1
    LVListado.ColumnHeaders.Add , , "Importe", 1200, 1
    LVListado.ColumnHeaders.Add , , "CodCentro", 0
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LVListado.SortKey = ColumnHeader.Index - 1
End Sub

