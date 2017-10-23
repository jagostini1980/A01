VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Presupuesto"
   ClientHeight    =   8055
   ClientLeft      =   -4710
   ClientTop       =   -1755
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmbDetalle 
      Caption         =   "Detalle Gastos"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2130
      TabIndex        =   13
      Top             =   7470
      Width           =   1695
   End
   Begin VB.TextBox TxtMontoMaxSinPres 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7695
      TabIndex        =   11
      Top             =   690
      Width           =   1200
   End
   Begin VB.TextBox TxtMontoActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   690
      Width           =   1200
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   6360
      Left            =   45
      TabIndex        =   6
      Top             =   1035
      Width           =   9465
      _cx             =   16695
      _cy             =   11218
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
      DataColCount    =   6
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
      UserResizeAnimate=   1
      UserResizing    =   3
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
      ColumnClickSort =   -1  'True
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   3
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
      StylesCollection=   $"A01_5B100.frx":0000
      ColumnsCollection=   "A01_5B100.frx":1DD9
      ValueItems      =   $"A01_5B100.frx":6067
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   495
      Left            =   3900
      TabIndex        =   5
      Top             =   7470
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5745
      TabIndex        =   2
      Top             =   7470
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   9465
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   350
         Left            =   6480
         TabIndex        =   0
         Top             =   180
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   1005
         TabIndex        =   3
         Top             =   190
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   110166019
         UpDown          =   -1  'True
         CurrentDate     =   38972
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   3465
         TabIndex        =   7
         Top             =   198
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin VB.Label LbCentroEmisor 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro Emisor:"
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
         Left            =   2205
         TabIndex        =   8
         Top             =   258
         Width           =   1245
      End
      Begin VB.Label LbPeriodo 
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
         Left            =   150
         TabIndex        =   4
         Top             =   258
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   225
      Top             =   7425
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Maximo Mensual Sin Presupuestar $:"
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
      Left            =   4545
      TabIndex        =   12
      Top             =   735
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Monto Actual Sin Presupuestar $:"
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
      TabIndex        =   10
      Top             =   735
      Width           =   2865
   End
End
Attribute VB_Name = "A01_5B100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Linea
    Linea As String
    Total As Double
    Max As Double
    Min As Double
End Type
 
Dim VecLineas() As Linea
Dim PorcTg As Double

Const VgNumero = "#0.00" 'esta constante es el formato de los numeros

Private Sub InicializarTodo()
On Error GoTo Errores
Dim i As Integer
Dim ColumnaActual As Integer
    MousePointer = vbHourglass
    CalPeriodo.Value = Date
    MousePointer = vbNormal
    
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.ListIndex = 0
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A015B100") = 2
    
    GridListado.Columns(1).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(6).Style.TextAlignment = sgAlignRightCenter
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub CmbDetalle_Click()
    Call CargarCmbCentrosDeCostosEmisor(A01_5B120.CmbCentroDeCostoEmisor)
    A01_5B120.CmbCentroDeCostoEmisor.ListIndex = CmbCentroDeCostoEmisor.ListIndex
    A01_5B120.CalPeriodo = CalPeriodo
    A01_5B120.CuentaContable = GridListado.Rows.At(GridListado.Row).Cells(1)
    A01_5B120.TxtCuentaContable = GridListado.Rows.At(GridListado.Row).Cells(2)
    A01_5B120.Traer
    A01_5B120.Show vbModal
End Sub

Private Sub CmdTraer_Click()
On Error GoTo Errores
Dim Sql As String
Dim i As Integer
Dim TbListado As New ADODB.Recordset
Dim IndexCol As Integer
Dim TotPres As Double
Dim TotUsado As Double

    MousePointer = vbHourglass
    TbListado.CursorLocation = adUseClient
    Sql = "SpOcConsultaEstadoDePresuesto @CentroEmisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                                     "', @Periodo='" & Format(CalPeriodo, "MM/yyyy") & "'"
          
    TbListado.Open Sql, Conec, adOpenKeyset
    GridListado.DataRowCount = 0
   
    With TbListado
        GridListado.DataRowCount = .RecordCount + 1
        
        While Not .EOF
            ' ProyeccionRecaudacion   ProyeccionPax    ProyeccionServicios
            GridListado.Array(i, 0) = !P_CuentaContable
            GridListado.Array(i, 1) = BuscarDescCta(!P_CuentaContable)
            GridListado.Array(i, 2) = Format(ValN(!PresMonto), "#,##0.00")
            GridListado.Array(i, 3) = Format(ValN(!UsadoMonto), "#,##0.00")
            GridListado.Array(i, 4) = Format(ValN(!SinPres), "#,##0.00")
            GridListado.Array(i, 5) = Format(ValN(!PresMonto) - ValN(!UsadoMonto) + ValN(!SinPres), "#,##0.00")
            TotPres = TotPres + ValN(!PresMonto)
            TotUsado = TotUsado + ValN(!UsadoMonto)
             i = i + 1
            .MoveNext
        Wend
        .Close
        Sql = "SpOCImporteSinPresupuestarAutorizacionDeCargaContable @CentroDeCosto ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                         "', @NroAutorizacion =0" & _
                         ", @Periodo=" & FechaSQL(CalPeriodo, "SQL")
        .Open Sql, Conec
        TxtMontoMaxSinPres = Format(ValN(!MontoSinPresupuestarMensual), "0.00")
        TxtMontoActual = Format(ValN(!MontoSinPres), "0.00")
        GridListado.Array(i, 1) = "Totales"
        GridListado.Array(i, 2) = Format(TotPres, "#,##0.00")
        GridListado.Array(i, 3) = Format(TotUsado, "#,##0.00")
        GridListado.Rows.At(i + 1).Style.Font.Bold = True
    End With
    
    MousePointer = vbNormal
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdSalir_Click()
        Unload Me
End Sub

Private Sub Form_Load()
    Call InicializarTodo
End Sub

Private Sub CmdExportar_Click()
    ArmarExcel CommonDialog1
    GenerarPlanilla CommonDialog1.Filename, CommonDialog1.FilterIndex
    MousePointer = vbNormal
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
        Call EncabezadoExcelGrid(ex, Caption, 7, Columnas)
        Call DatosExcelGrid(ex, GridListado, 7, Filas)
        
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
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(CalPeriodo.Value, "MM/yyyy")
        .Range("C4").Select
        .ActiveCell.FormulaR1C1 = "Centro De Costo: " & CmbCentroDeCostoEmisor.Text
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Monto Actual Sin Presupuestar $: " & TxtMontoActual.Text
        .Range("C5").Select
        .ActiveCell.FormulaR1C1 = "Maximo Mensual Sin Presupuestar $: " & TxtMontoMaxSinPres.Text
     
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 7, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub GridListado_Click()
    CmbDetalle.Enabled = GridListado.RowCount - 1 > GridListado.Row
End Sub

Private Sub GridListado_DblClick()
    If GridListado.Rows.At(GridListado.Row).Cells(1) = "5121" Then
        Call CargarCmbCentrosDeCostosEmisor(A01_5B110.CmbCentroDeCostoEmisor)
        A01_5B110.CmbCentroDeCostoEmisor.ListIndex = CmbCentroDeCostoEmisor.ListIndex
        A01_5B110.CalPeriodo = CalPeriodo
        A01_5B110.Traer
        A01_5B110.Show vbModal
    End If
End Sub
