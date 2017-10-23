VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B400 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Totales  por Empresa"
   ClientHeight    =   6015
   ClientLeft      =   -4710
   ClientTop       =   -1755
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   4560
      Left            =   45
      TabIndex        =   6
      Top             =   810
      Width           =   5550
      _cx             =   9790
      _cy             =   8043
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
      DataColCount    =   3
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
      ColumnClickSort =   0   'False
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
      StylesCollection=   $"A01_5B400.frx":0000
      ColumnsCollection=   $"A01_5B400.frx":1DD9
      ValueItems      =   $"A01_5B400.frx":30F1
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   495
      Left            =   1057
      TabIndex        =   5
      Top             =   5445
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2902
      TabIndex        =   2
      Top             =   5445
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Left            =   52
      TabIndex        =   1
      Top             =   0
      Width           =   5550
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   350
         Left            =   4185
         TabIndex        =   0
         Top             =   360
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   375
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   49545219
         UpDown          =   -1  'True
         CurrentDate     =   38972
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   390
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
         Left            =   1260
         TabIndex        =   8
         Top             =   150
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
         Top             =   165
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "A01_5B400"
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
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A015B400") = 2
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
Errores:
    ManipularError Err.Number, Err.Description
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
    Sql = "SpOcConsultaTotalesPorEmpresa @CentroDeCostoEmisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                                     "', @Año=" & CalPeriodo.Year & _
                                     " , @Mes=" & CalPeriodo.Month
          
    TbListado.Open Sql, Conec, adOpenKeyset
    GridListado.DataRowCount = 0
   
    With TbListado
        GridListado.DataRowCount = .RecordCount + 1
        
        While Not .EOF
            GridListado.Array(i, 0) = !O_EmpresaFacturaANombreDe
            GridListado.Array(i, 1) = BuscarDescEmpresa(!O_EmpresaFacturaANombreDe)
            GridListado.Array(i, 2) = Format(ValN(!Importe), "#,##0.00")
            TotUsado = TotUsado + ValN(!Importe)
             i = i + 1
            .MoveNext
        Wend
        .Close
        GridListado.Array(i, 1) = "Totales"
        GridListado.Array(i, 2) = Format(TotUsado, "#,##0.00")
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
        Call EncabezadoExcelGrid(ex, Caption, 6, Columnas)
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
        .ActiveCell.FormulaR1C1 = "Periodo: " & Format(CalPeriodo.Value, "MM/yyyy")
        .Range("C4").Select
        .ActiveCell.FormulaR1C1 = "Centro De Costo: " & CmbCentroDeCostoEmisor.Text
     
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 6, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

