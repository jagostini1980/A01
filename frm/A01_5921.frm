VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5921 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Presupuestado por Cuenta Contable - Centro Emisor"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   350
      Left            =   953
      TabIndex        =   4
      Top             =   8145
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
      Left            =   2303
      TabIndex        =   3
      Top             =   8145
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3668
      TabIndex        =   0
      Top             =   8145
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   7005
      Left            =   60
      TabIndex        =   1
      Top             =   855
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   12356
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
   Begin VB.Label LbCuenta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   450
      Width           =   5820
   End
   Begin VB.Label LbTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Centro Emisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   5820
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
      Left            =   5265
      TabIndex        =   2
      Top             =   7920
      Width           =   510
   End
End
Attribute VB_Name = "A01_5921"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cuenta As String
Public Centro As String
Public FechaDesde As Date
Public FechaHasta As Date

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
    Sql = "SpOcConsultaPresupuestoFinancieroDetallePresupuestadoXCuenta @Periodo = '" & Format(FechaDesde, "MM/yyyy") & _
                                                                    "', @Cuenta = '" & Cuenta & _
                                                                    "', @CentroDeCostoEmisor='" & Centro & "'"
    With RsCargar
        .Open Sql, Conec
        
      LVListado.Sorted = False
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add

            LVListado.ListItems(i).Text = Format(!P_FechaAprobacion, "dd/MM/yyyy")
            LVListado.ListItems(i).SubItems(1) = Format(!P_NumeroPresupuesto, "00000000")
            LVListado.ListItems(i).SubItems(2) = Format(!Importe * PorcDias, "#,##0")

            Importe = Importe + !Importe * PorcDias
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
        Call EncabezadoExcel(ex, LVListado, Caption, 8)
        Call DatosExcel(ex, LVListado, 8)
        
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
        .ActiveCell.FormulaR1C1 = "Fecha desde: " & FechaDesde & " hasta " & FechaHasta
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Centro Emisor: " & LbTitulo
        .Range("A6").Select
        .ActiveCell.FormulaR1C1 = "Cuenta Contable: " & LbCuenta
        .Range("C" & 9 + LVListado.ListItems.Count).Select
        .ActiveCell.Formula = "=Sum($C7:$C" & 8 + LVListado.ListItems.Count & ")"
        .Range("A" & 9 + LVListado.ListItems.Count).Select
        .ActiveCell.FormulaR1C1 = "Total ==>"

        ColorFondo = &HC0E0FF
        Call FormatearExcelConTotal(ex, LVListado, 8, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5921.Show vbModal
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "Fecha", adVarChar, 20
    RsListado.Fields.Append "NroPres", adVarChar, 8
    RsListado.Fields.Append "Importe", adVarChar, 100
    RsListado.Open
    i = 1
    While i <= LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
           RsListado!Fecha = .Text
           RsListado!NroPres = .SubItems(1)
           RsListado!Importe = .SubItems(2)
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5921.TxtCuentaCostable = LbCuenta.Caption
    ListA01_5921.TxtCentroEmisor = LbTitulo.Caption
    ListA01_5921.TxtPeriodo.Text = FechaDesde & " hasta " & FechaHasta
    ListA01_5921.DataControl1.Recordset = RsListado
    ListA01_5921.Zoom = -1
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
    LVListado.ColumnHeaders.Add , , "Fecha Aprobación", (LVListado.Width - 250) / 3
    LVListado.ColumnHeaders.Add , , "Presupuesto Nº", (LVListado.Width - 250) / 3, 1
    LVListado.ColumnHeaders.Add , , "Importe", (LVListado.Width - 250) / 3, 1
    LVListado.ColumnHeaders.Add , , "CodCentroEmisor", 0, 1
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LVListado.SortKey = ColumnHeader.Index - 1
End Sub

