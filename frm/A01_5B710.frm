VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B710 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Gastos"
   ClientHeight    =   8010
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   6630
      Left            =   45
      TabIndex        =   5
      Top             =   855
      Width           =   10410
      _cx             =   18362
      _cy             =   11695
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   0
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   7
      HeadingRowCount =   1
      HeadingColCount =   1
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   -2147483633
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   12640511
      UserResizeAnimate=   0
      UserResizing    =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   0
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   1
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   -1  'True
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   1
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   2
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   -1  'True
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"A01_5B710.frx":0000
      ColumnsCollection=   "A01_5B710.frx":1DD9
      ValueItems      =   $"A01_5B710.frx":69FB
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8985
      TabIndex        =   4
      Top             =   7530
      Width           =   1455
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   7515
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   810
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10425
      Begin VB.TextBox TxtCuenta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1815
         TabIndex        =   6
         Top             =   405
         Width           =   3420
      End
      Begin MSComCtl2.DTPicker CalDesde 
         Height          =   315
         Left            =   375
         TabIndex        =   1
         Top             =   405
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MMM/yyyy"
         Format          =   24248323
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta Contable"
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
         Left            =   1830
         TabIndex        =   7
         Top             =   180
         Width           =   1425
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
         Left            =   375
         TabIndex        =   2
         Top             =   165
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   450
      Top             =   7515
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "A01_5B710"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Jerarquia As String
Public Cuenta As String
Public SubCentro As String

Public Sub CargarLV()
  Dim Sql As String
  Dim RsListado As New ADODB.Recordset
  Dim RsPresFinanciero As New ADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim CodRubro As String
  Dim CodCuenta As String
  Dim DescRubro As String
  Dim TotalRubro As Double
  
  RsListado.CursorLocation = adUseClient
  RsListado.CursorType = adOpenKeyset
  MousePointer = vbHourglass

On Error GoTo ErrorTraer:
  With RsListado
     '******** Detalle Gastos Por Cuenta ***********
      Sql = "Admin.ElPulquiAdministracion.dbo.SpOcGastosResultadoEconomicoPorUnidadDeNegocioDetalle @PerGasto='" & Format(CalDesde, "MMyy") & _
                                                                                                       "', @Cuenta='" & Cuenta & "'"
                                                      
     .Open Sql, Conec
   
     If .RecordCount > 0 Then
     
        For i = 1 To .RecordCount
            GridListado.DataRowCount = GridListado.DataRowCount + 1
            GridListado.Array(GridListado.DataRowCount - 1, 0) = !C_Fecha
            GridListado.Array(GridListado.DataRowCount - 1, 1) = Trim(!C_Concepto)
            GridListado.Array(GridListado.DataRowCount - 1, 2) = BuscarDescCentroEmisorPorJerarquia(VerificarNulo(!C_Emisor))
            GridListado.Array(GridListado.DataRowCount - 1, 3) = BuscarDescCentro(BuscarCodigoPorCodSecundario(Format(Trim(!C_Rubro), "000")))
            GridListado.Array(GridListado.DataRowCount - 1, 4) = Format(VerificarNulo(!C_Importe, "N"), "#,##0")
            TotalRubro = TotalRubro + VerificarNulo(!C_Importe, "N")
            .MoveNext
        Next
     End If

    GridListado.DataRowCount = GridListado.DataRowCount + 1
    GridListado.Rows.At(GridListado.DataRowCount).Style.Font.Bold = True
    GridListado.Array(GridListado.DataRowCount - 1, 1) = "Total Cuenta"
    GridListado.Array(GridListado.DataRowCount - 1, 4) = Format(TotalRubro, "#,##0")
     
    CmdExp.Enabled = .RecordCount > 0
    
    .Close
  
 End With
    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal
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
Dim Filas As Long
Dim i As Integer
Dim Columnas As Integer

    'cuanta las columnas visibles
    For i = 1 To GridListado.ColCount - 1
        If Not GridListado.Columns(i).Hidden Then
            Columnas = Columnas + 1
        End If
    Next
    
    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcelGrid(ex, Caption, 8, Columnas)
        Call DatosExcelGrid(ex, GridListado, 8, Filas)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To GridListado.ColCount
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Período: " & Format(CalDesde, "MM/yyyy")
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Cuenta Contable: " & TxtCuenta.Text
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 8, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    'GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    'GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    'GridListado.Columns(6).Style.TextAlignment = sgAlignRightCenter

End Sub

