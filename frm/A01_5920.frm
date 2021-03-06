VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_5920 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Presupuestado"
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   350
      Left            =   1834
      TabIndex        =   4
      Top             =   6795
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   45
      Top             =   6705
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   3195
      TabIndex        =   3
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4557
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
Attribute VB_Name = "A01_5920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rubro As String
Public FechaDesde As Date
Public FechaHasta As Date
Private VecCodCentro() As String
Private VecCodCuenta() As String

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CargarListado()
Dim Sql As String
Dim i As Integer
Dim Importe As Double
Dim RsCargar As New ADODB.Recordset
Dim PorcDias As Double

On Error GoTo Error
    LVListado.ListItems.Clear
    MousePointer = vbHourglass
    PorcDias = (DateDiff("d", FechaDesde, FechaHasta) + 1) / LenMes(Month(FechaDesde), Year(FechaDesde))

    'Realiza la consulta
    Sql = "SpOcConsultaPresupuestoFinancieroDetallePresupuestado @Periodo = " & FechaSQL(CStr(FechaDesde), "SQL") & _
                                                              ", @Rubro = '" & Rubro & "'"
                                        
    With RsCargar
        .Open Sql, Conec
        ReDim VecCodCentro(.RecordCount)
        ReDim VecCodCuenta(.RecordCount)
      LVListado.Sorted = False
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add
            VecCodCentro(i) = !P_CentroDeCostoEmisor
            VecCodCuenta(i) = !Cuenta
            LVListado.ListItems(i).Text = BuscarDescCentroEmisor(!P_CentroDeCostoEmisor)
            LVListado.ListItems(i).SubItems(1) = BuscarDescCta(!Cuenta)
            LVListado.ListItems(i).SubItems(2) = Format(!Importe * PorcDias, "#,##0")
            LVListado.ListItems(i).SubItems(3) = i
            
            Importe = Importe + !Importe
            .MoveNext
        Wend
    End With
    LVListado.Sorted = True
    LbTotal.Caption = "Total: " & Format(Importe, "#,##0")
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
        Call EncabezadoExcel(ex, LVListado, Caption, 6)
        Call DatosExcel(ex, LVListado, 6)
        
        '--------AJUSTO LOS TAMA�OS DE LAS COLUMNAS
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
        .ActiveCell.FormulaR1C1 = "Fecha desde: " & FechaDesde & " hasta " & FechaHasta
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
   ListA01_5920.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "Centro", adVarChar, 100
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "Importe", adVarChar, 20
    
    RsListado.Open
    i = 1
    While i <= LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
           RsListado!Centro = .Text
           RsListado!Cuenta = .SubItems(1)
           RsListado!Importe = .SubItems(2)
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5920.TxtPeriodo.Text = FechaDesde & " hasta " & FechaHasta
    ListA01_5920.DataControl1.Recordset = RsListado
    ListA01_5920.Zoom = -1
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
        MsgBox "La Exportaci�n se raliz� correctamente", vbInformation, "Exportaci�n"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarListado
End Sub

Private Sub CrearEncabezado()
    LVListado.ColumnHeaders.Add , , "Centro Emisor", (LVListado.Width - 1250) / 2
    LVListado.ColumnHeaders.Add , , "Cuenta Contable", (LVListado.Width - 1250) / 2
    LVListado.ColumnHeaders.Add , , "Importe", 1000, 1
    LVListado.ColumnHeaders.Add , , "Index", 0, 1
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LVListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LVListado_DblClick()
    A01_5921.Cuenta = VecCodCuenta(Val(LVListado.SelectedItem.SubItems(3)))
    A01_5921.Centro = VecCodCentro(Val(LVListado.SelectedItem.SubItems(3)))
    A01_5921.FechaDesde = FechaDesde
    A01_5921.FechaHasta = FechaHasta
    A01_5921.LbTitulo = "Centro Emisor: " & LVListado.SelectedItem.Text
    A01_5921.LbCuenta = "Cuenta: " & LVListado.SelectedItem.SubItems(1)
    
    A01_5921.Show vbModal
End Sub

