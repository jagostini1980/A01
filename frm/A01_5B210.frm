VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B210 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Rubro Contable"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4558
      TabIndex        =   8
      Top             =   6795
      Width           =   1230
   End
   Begin VB.CommandButton CmdExpExcel 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   3196
      TabIndex        =   7
      Top             =   6795
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   8925
      Begin VB.TextBox TxtCuenta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3870
         TabIndex        =   6
         Top             =   180
         Width           =   4965
      End
      Begin MSComCtl2.DTPicker CalDesde 
         Height          =   315
         Left            =   945
         TabIndex        =   3
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MMM/yyyy"
         Format          =   59637763
         CurrentDate     =   38940
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta Contable:"
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
         Left            =   2340
         TabIndex        =   5
         Top             =   240
         Width           =   1485
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
         Left            =   105
         TabIndex        =   4
         Top             =   240
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   45
      Top             =   6705
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   6045
      Left            =   45
      TabIndex        =   1
      Top             =   675
      Width           =   8925
      _cx             =   15743
      _cy             =   10663
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
      DataColCount    =   4
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
      StylesCollection=   $"A01_5B210.frx":0000
      ColumnsCollection=   $"A01_5B210.frx":1DD9
      ValueItems      =   $"A01_5B210.frx":35A3
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
      Left            =   8415
      TabIndex        =   0
      Top             =   6885
      Width           =   510
   End
End
Attribute VB_Name = "A01_5B210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cuenta As String

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Public Sub CargarListado()
Dim Sql As String
Dim i As Integer
Dim Importe As Double
Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    GridListado.DataRowCount = 0
    MousePointer = vbHourglass
    'Realiza la consulta
    Sql = "SpOcConsultaRubrosContableDetalle @Periodo = '" & Format(CalDesde, "MMyy") & _
                                         "', @Cuenta = '" & Cuenta & "'"
                                        
    With RsCargar
        .Open Sql, Conec
        
        GridListado.DataRowCount = .RecordCount
        While Not .EOF
            GridListado.Array(i, 0) = !C_Fecha
            GridListado.Array(i, 1) = !C_Empresa
            GridListado.Array(i, 2) = !C_Concepto
            GridListado.Array(i, 3) = Format(!C_Importe, "#,##0.00")
            Importe = Importe + Format(!C_Importe, "#,##0.00")
            i = i + 1
            .MoveNext
        Wend
    End With
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
Dim Filas As Long
Dim ColorFondo As Long

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcelGrid(ex, Caption & " - " & TxtCuenta, 6, 4)
        Call DatosExcelGrid(ex, GridListado, 6, Filas)
        
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
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(CalDesde, "MMM/yyyy")
        .Range("D" & 7 + GridListado.DataRowCount).Select
        .ActiveCell.Formula = "=Sum($D7:$D" & 6 + GridListado.DataRowCount & ")"
        .Range("A" & 7 + GridListado.DataRowCount).Select
        .ActiveCell.FormulaR1C1 = "Total ==>"

        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 6, Filas + 1, 4, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub Form_Load()
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
End Sub
