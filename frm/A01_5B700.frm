VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B700 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Gastos por Unidade de Negocio"
   ClientHeight    =   8010
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDetalle 
      Caption         =   "&Detalle Cta."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2670
      TabIndex        =   7
      Top             =   7560
      Width           =   1455
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   6870
      Left            =   45
      TabIndex        =   6
      Top             =   615
      Width           =   11250
      _cx             =   19844
      _cy             =   12118
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
      DataColCount    =   9
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
      StylesCollection=   $"A01_5B700.frx":0000
      ColumnsCollection=   "A01_5B700.frx":1DD9
      ValueItems      =   $"A01_5B700.frx":7CAB
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5700
      TabIndex        =   5
      Top             =   7545
      Width           =   1455
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4185
      TabIndex        =   4
      Top             =   7545
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   570
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      Begin VB.CheckBox ChkNuevo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consulta Nueva"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   210
         Width           =   1770
      End
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   4770
         TabIndex        =   1
         Top             =   150
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalDesde 
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   108789763
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
         Left            =   705
         TabIndex        =   3
         Top             =   210
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
Attribute VB_Name = "A01_5B700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargarLV()
  Dim Sql As String
  Dim RsListado As New ADODB.Recordset
  Dim RsPresFinanciero As New ADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim CodRubro As String
  Dim CodCuenta As String
  Dim DescRubro As String
  Dim TotalRubro As Double
  Dim TotalPP As Double
  Dim TotalTur As Double
  Dim TotalGral As Double
  Dim TotalOtros As Double
  Dim CodSubCentro As String
  
  RsListado.CursorLocation = adUseClient
  RsListado.CursorType = adOpenKeyset
  MousePointer = vbHourglass

'On Error GoTo ErrorTraer:
  With RsListado
     '******** Total Gastos Por Rubro ***********
     If ChkNuevo.Value = 1 Then
        Sql = "Admin.ElPulquiAdministracion.dbo.SpOcGastosResultadoEconomicoPorUnidadDeNegocioNuevo @PerGasto='" & Format(CalDesde, "MMyy") & "'"
     Else
        Sql = "Admin.ElPulquiAdministracion.dbo.SpOcGastosResultadoEconomicoPorUnidadDeNegocio @PerGasto='" & Format(CalDesde, "MMyy") & "'"
     End If
     .Open Sql, Conec
     GridListado.Groups.RemoveAll
     GridListado.DataRowCount = 0
     If .RecordCount > 0 Then
        CodRubro = !CodRubro
        DescRubro = !RubroContable
        
        'GridListado.DataRowCount = .RecordCount
        For i = 1 To .RecordCount
            GridListado.DataRowCount = GridListado.DataRowCount + 1
            GridListado.Array(GridListado.DataRowCount - 1, 0) = !CodRubro & " - " & Convertir(!RubroContable)
            GridListado.Array(GridListado.DataRowCount - 1, 1) = !C_Cuenta & " - " & BuscarDescCta(!C_Cuenta)
            GridListado.Array(GridListado.DataRowCount - 1, 2) = Format(VerificarNulo(!ImportePax, "N"), "#,##0")
            'GridListado.Array(GridListado.DataRowCount - 1, 4) = !CodRubro
            GridListado.Array(GridListado.DataRowCount - 1, 5) = !C_Cuenta
            GridListado.Array(GridListado.DataRowCount - 1, 3) = Format(VerificarNulo(!ImporteEnc, "N"), "#,##0")
            GridListado.Array(GridListado.DataRowCount - 1, 6) = Format(VerificarNulo(!ImporteTur, "N"), "#,##0")
            GridListado.Array(GridListado.DataRowCount - 1, 7) = Format(VerificarNulo(!ImporteOtros, "N"), "#,##0")
            GridListado.Array(GridListado.DataRowCount - 1, 8) = Format(ValN(!ImporteOtros) + ValN(!ImporteTur) + ValN(!ImporteEnc) + ValN(!ImportePax), "#,##0")
            TotalRubro = TotalRubro + VerificarNulo(!ImportePax, "N")
            TotalPP = TotalPP + ValN(!ImporteEnc)
            TotalTur = TotalTur + ValN(!ImporteTur)
            TotalOtros = TotalOtros + ValN(!ImporteOtros)
            TotalGral = TotalGral + ValN(!ImporteOtros) + ValN(!ImporteTur) + ValN(!ImporteEnc) + ValN(!ImportePax)
            
            .MoveNext
            If Not .EOF Then
                If CodRubro <> !CodRubro Then
                    GridListado.DataRowCount = GridListado.DataRowCount + 1
                    GridListado.Rows.At(GridListado.DataRowCount).Style.Font.Bold = True
                    GridListado.Array(GridListado.DataRowCount - 1, 0) = CodRubro & " - " & Convertir(DescRubro)
                    GridListado.Array(GridListado.DataRowCount - 1, 1) = "Total Rubro"
                    GridListado.Array(GridListado.DataRowCount - 1, 2) = Format(TotalRubro, "#,##0")
                    GridListado.Array(GridListado.DataRowCount - 1, 3) = Format(TotalPP, "#,##0")
                    GridListado.Array(GridListado.DataRowCount - 1, 6) = Format(TotalTur, "#,##0")
                    GridListado.Array(GridListado.DataRowCount - 1, 7) = Format(TotalOtros, "#,##0")
                    GridListado.Array(GridListado.DataRowCount - 1, 8) = Format(TotalGral, "#,##0")
                    TotalRubro = 0
                    TotalPP = 0
                    TotalTur = 0
                    TotalOtros = 0
                    TotalGral = 0
                    DescRubro = !RubroContable
                    CodRubro = !CodRubro
                End If
            End If
        Next
     End If

    GridListado.DataRowCount = GridListado.DataRowCount + 1
    GridListado.Rows.At(GridListado.DataRowCount).Style.Font.Bold = True
    GridListado.Array(GridListado.DataRowCount - 1, 0) = CodRubro & " - " & Convertir(DescRubro)
    GridListado.Array(GridListado.DataRowCount - 1, 1) = "Total Rubro"
    GridListado.Array(GridListado.DataRowCount - 1, 2) = Format(TotalRubro, "#,##0")
    GridListado.Array(GridListado.DataRowCount - 1, 3) = Format(TotalPP, "#,##0")
    GridListado.Array(GridListado.DataRowCount - 1, 6) = Format(TotalTur, "#,##0")
    GridListado.Array(GridListado.DataRowCount - 1, 7) = Format(TotalOtros, "#,##0")
    GridListado.Array(GridListado.DataRowCount - 1, 8) = Format(TotalGral, "#,##0")
    
    CmdExp.Enabled = .RecordCount > 0
    
    .Close
    GridListado.Groups.Add "ColRubro"
    GridListado.Redraw sgRedrawAll
 End With
    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal
End Sub

Private Sub CmdDetalle_Click()
   A01_5B710.TxtCuenta = GridListado.Rows.At(GridListado.Row).Cells(2)
   A01_5B710.Cuenta = GridListado.Rows.At(GridListado.Row).Cells(6)
   A01_5B710.CalDesde.Value = CalDesde.Value
   Call A01_5B710.CargarLV
   A01_5B710.Show
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
Dim InicioTabla2 As Integer
Dim Fila As Long
Dim Filas As Long
Dim i As Integer
Dim j As Integer
Dim Rubro As String
Dim TotalContable As Double
Dim TotalPres As Double
Dim TotalPresF As Double
Dim TotalFinanciero As Double
Dim Arr As SGArray

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        ColorFondo = &HC0E0FF

        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = Caption
        .Range("A1:C1").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge  'COMBINAR CELDAS
        
        With .Selection.Font
            .Name = "Arial"
            .Size = 20
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
        .Selection.Font.Bold = True

        '-------- GENERO LOS DATOS ------------------------------
        '*** carga los ingresos
    
        InicioTabla2 = 6
        
         Fila = InicioTabla2 + 1
         .Range("A" & Trim(Fila - 1) & ":F" & Trim(Fila - 1)).Select
         .Selection.Font.Bold = True
         .Range("A" & Trim(Fila - 1)).Value = "Cuenta Contable"
         .Range("B" & Trim(Fila - 1)).Value = "Lineas"
         .Range("C" & Trim(Fila - 1)).Value = "Pulqui Pack"
         .Range("D" & Trim(Fila - 1)).Value = "Turismo"
         .Range("E" & Trim(Fila - 1)).Value = "Gtos. Indirectos"
         .Range("F" & Trim(Fila - 1)).Value = "Total"
         
         
         Set Arr = GridListado.Array
         .Range("A" & Trim(Fila)).Value = Arr(i, 0)
         Fila = Fila + 1
         Rubro = Arr(i, 0)
         
         For i = 0 To Arr.RowCount - 1
            If Rubro <> Arr(i, 0) Then
                .Range("A" & Trim(Fila)).Value = Arr(i, 0)
                Fila = Fila + 1
                Rubro = Arr(i, 0)
            End If
               .Range("A" & Trim(Fila)).Value = Space(10) & Arr(i, 1)
               .Range("B" & Trim(Fila)).Value = FormatNumber(Arr(i, 2), 0, vbUseDefault, vbUseDefault, vbFalse)
               .Range("C" & Trim(Fila)).Value = FormatNumber(Arr(i, 3), 0, vbUseDefault, vbUseDefault, vbFalse)
               .Range("D" & Trim(Fila)).Value = FormatNumber(Arr(i, 6), 0, vbUseDefault, vbUseDefault, vbFalse)
               .Range("E" & Trim(Fila)).Value = FormatNumber(Arr(i, 7), 0, vbUseDefault, vbUseDefault, vbFalse)
               .Range("F" & Trim(Fila)).Value = FormatNumber(Arr(i, 8), 0, vbUseDefault, vbUseDefault, vbFalse)

                  If Mid(Arr(i, 1), 1, 5) = "Total" Then
                   .Range("A" & Trim(Fila)).Value = Trim(.Range("A" & Trim(Fila)).Value)
                   .Range("A" & Trim(Fila) & ":F" & Trim(Fila)).Select
                   .Selection.Font.Bold = True
                   TotalContable = TotalContable + ValN(FormatNumber(Arr(i, 8), 2, vbUseDefault, vbUseDefault, vbFalse))
               End If
               Fila = Fila + 1
         Next

         Call FormatearExcelGrid(ex, InicioTabla2, Fila - InicioTabla2 - 1, 6, ColorFondo)
         .Range("A" & Fila).Select
         .ActiveCell.FormulaR1C1 = "Totales ==>"
         .Range("F" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalContable
         .Range("B" & Trim(InicioTabla2 + 1) & ":F" & Trim(Fila + 1)).Select
         .Selection.NumberFormat = "#,##0"
         
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To 6
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("E1").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("E2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Período: " & Format(CalDesde, "MM/yyyy")
        .Range("A4").Select
        '.ActiveCell.FormulaR1C1 = "Centro de Costos: " & CmbCentroDeCostoEmisor.Text
      
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTraer_Click()
      Call CargarLV
End Sub

Private Sub Form_Load()
    Dim i As Integer
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(6).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(7).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(8).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(9).Style.TextAlignment = sgAlignRightCenter
    
    CalDesde.Value = Date
   
End Sub

Private Sub GridListado_SelChange(CancelSelect As Boolean)
On Error Resume Next
    CmdDetalle.Enabled = False
    CmdDetalle.Enabled = GridListado.Rows.At(GridListado.Row).Cells(6) <> ""
End Sub
