VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_7400 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OP Pendientes de Devolución"
   ClientHeight    =   8055
   ClientLeft      =   -4710
   ClientTop       =   -1755
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   7350
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   12060
      _cx             =   21272
      _cy             =   12965
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
      DataColCount    =   8
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
      AllowEdit       =   0   'False
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
      StylesCollection=   $"A01_7400.frx":0000
      ColumnsCollection=   "A01_7400.frx":1DD9
      ValueItems      =   $"A01_7400.frx":7335
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   495
      Left            =   4297
      TabIndex        =   1
      Top             =   7470
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6142
      TabIndex        =   0
      Top             =   7470
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   7515
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "A01_7400"
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
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(8).Style.TextAlignment = sgAlignRightCenter
    
MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub Traer()
On Error GoTo Errores
Dim Sql As String
Dim i As Integer
Dim TbListado As New ADODB.Recordset
Dim OP As Long
Dim TotPres As Double
Dim TotUsado As Double

    MousePointer = vbHourglass
    TbListado.CursorLocation = adUseClient
     Sql = "Admin.ElPulquiAdministracion..SpOrdenesPendientesDeDevolucion"
           
     TbListado.Open Sql, Conec, adOpenKeyset
     GridListado.DataRowCount = 0
    
     With TbListado
         GridListado.DataRowCount = .RecordCount
         
         While Not .EOF
            If OP <> !NROOP Then
                OP = !NROOP
                GridListado.Array(i, 0) = !EMPRESA
                GridListado.Array(i, 1) = Format(!Fecha, "dd/MM/yyyy")
                GridListado.Array(i, 2) = Replace(!NROOP, " ", "0")
                GridListado.Array(i, 3) = !Descripcion
                GridListado.Array(i, 4) = Format(ValN(!Importe), "#,##0.00")
                GridListado.Array(i, 5) = !CONCEPTORESUMEN
             Else
                GridListado.Array(i, 0) = ""
                GridListado.Array(i, 1) = ""
                GridListado.Array(i, 2) = ""
                GridListado.Array(i, 3) = ""
                GridListado.Array(i, 4) = ""
                GridListado.Array(i, 5) = ""
             End If
             
             If ValN(!IMPORTECHEQUE) > 0 Then
                GridListado.Array(i, 7) = Format(ValN(!IMPORTECHEQUE), "#,##0.00")
                GridListado.Array(i, 6) = Format(!FECHACHEQUE, "dd/MM/yyyy")
             Else
                GridListado.Array(i, 6) = ""
                GridListado.Array(i, 7) = ""
             End If
             
              i = i + 1
             .MoveNext
         Wend
         .Close
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
    Call Traer
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
        Call EncabezadoExcelGrid(ex, Caption, 4, Columnas)
        Call DatosExcelGrid(ex, GridListado, 4, Filas)
        
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
       
        
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 4, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

