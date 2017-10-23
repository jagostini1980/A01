VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5900vieja 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Presupuesto/Financiero"
   ClientHeight    =   8010
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDetalleFinanciero 
      Caption         =   "Detalle &Financiero"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5556
      TabIndex        =   19
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdDetallePres 
      Caption         =   "Detalle &Pres."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3948
      TabIndex        =   18
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdDetalleCont 
      Caption         =   "Detalle &Contable"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2340
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10380
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7164
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8772
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   2565
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   990
         TabIndex        =   2
         Top             =   210
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   58589187
         CurrentDate     =   38940
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo:"
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
         Left            =   165
         TabIndex        =   3
         Top             =   285
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   450
      Top             =   7515
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   2490
      Left            =   45
      TabIndex        =   7
      Top             =   675
      Width           =   11805
      _cx             =   20823
      _cy             =   4392
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
      DataColCount    =   5
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
      StylesCollection=   $"A01_5900 Vieja.frx":0000
      ColumnsCollection=   $"A01_5900 Vieja.frx":1DD9
      ValueItems      =   $"A01_5900 Vieja.frx":3A73
   End
   Begin MSComctlLib.TreeView TvContable 
      Height          =   3480
      Left            =   45
      TabIndex        =   8
      Top             =   3735
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   6138
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contable/Pres."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8730
      TabIndex        =   16
      Top             =   3465
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Financiero/Pres."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   10305
      TabIndex        =   15
      Top             =   3465
      Width           =   1500
   End
   Begin VB.Label LbDesvio 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8730
      TabIndex        =   14
      Top             =   3195
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4005
      TabIndex        =   13
      Top             =   3195
      Width           =   4650
   End
   Begin VB.Label LbImpFinanciero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Financiero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7155
      TabIndex        =   12
      Top             =   3465
      Width           =   1500
   End
   Begin VB.Label LbImpPres 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Presupuestado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   11
      Top             =   3465
      Width           =   1500
   End
   Begin VB.Label LbImpContable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4005
      TabIndex        =   10
      Top             =   3465
      Width           =   1500
   End
   Begin VB.Label LbGOF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   7245
      Width           =   11790
   End
End
Attribute VB_Name = "A01_5900vieja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IvaCredito As Double
Private VecAgrupacionRubros() As TipoAgrupacionRubroContrable

Private Sub CargarLV()
  Dim Sql As String
  Dim RsListado As New ADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim TotalIvaDebito As Double
  Dim TotalIngresos As Double
  Dim TotalReal As Double
  Dim TotalIvaReal As Double
  Dim TotalEgresos As Double
  Dim TotalPres As Double
  Dim TotalFinanciero As Double
  
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   MousePointer = vbHourglass

On Error GoTo ErrorTraer:
   Call CargarTv
   Sql = "SpOcConsultaPresupuestoFinancieroIngresos @Periodo='" & Format(CalPeriodo, "MM/yyyy") & "'"
  
  '******** Total Ingresos ***********
   RsListado.Open Sql, Conec
   With RsListado
        'limpia el LV
        CmdImprimir.Enabled = True
        CmdExp.Enabled = True
        
        If .RecordCount > 0 Then
           For i = 1 To .RecordCount
               For j = 1 To GridListado.DataRowCount - 2
                    If Trim(GridListado.Rows.At(j).Cells(5)) = Trim(!T_Negocio) Then
                        GridListado.Rows.At(j).Cells(2) = Format(!T_ProyeccionIngresosFinancieros, "#,##0.00")
                        Exit For
                    End If
               Next
               TotalIvaDebito = TotalIvaDebito + !T_ProyeccionIngresosFinancieros * (!U_PorcenjateDeIva / 100)
               TotalIngresos = TotalIngresos + !T_ProyeccionIngresosFinancieros
               .MoveNext
           Next
        End If
        GridListado.Rows.At(GridListado.Rows.Count - 2).Cells(2) = Format(TotalIvaDebito, "#,##0.00")
        GridListado.Rows.At(GridListado.Rows.Count - 1).Cells(2) = Format(TotalIngresos + TotalIvaDebito, "#,##0.00")
        .Close
        
        '******** Total Real ***********
         Sql = "SpOcConsultaPresupuestoFinancieroIngresosReal @Periodo='" & Format(CalPeriodo, "MM/yyyy") & "'"
        .Open Sql, Conec
        If .RecordCount > 0 Then
           For i = 1 To .RecordCount
               For j = 1 To GridListado.DataRowCount - 2
                    If Trim(GridListado.Rows.At(j).Cells(5)) = Trim(!T_Negocio) Then
                        GridListado.Rows.At(j).Cells(3) = Format(VerificarNulo(!T_RealIngresosFinancieros, "N"), "#,##0.00")
                        Exit For
                    End If
               Next
               TotalIvaReal = TotalIvaReal + VerificarNulo(!T_RealIngresosFinancieros, "N") * (!U_PorcenjateDeIva / 100)
               TotalReal = TotalReal + VerificarNulo(!T_RealIngresosFinancieros, "N")
               .MoveNext
           Next
        End If
        GridListado.Rows.At(GridListado.Rows.Count - 2).Cells(3) = Format(TotalIvaReal, "#,##0.00")
        GridListado.Rows.At(GridListado.Rows.Count - 1).Cells(3) = Format(TotalReal + TotalIvaReal, "#,##0.00")
       .Close
       
        '******** Total Contable ***********
        Sql = "SpOcConsultaPresupuestoFinancieroEgresos @Periodo=" & FechaSQL("01/" & Format(CalPeriodo, "MM/yyyy"), "SQL")
        .Open Sql, Conec
        If .RecordCount > 0 Then
           For i = 1 To .RecordCount
               For k = 1 To TvContable.Nodes.Count
                   If TvContable.Nodes(k).Tag = !CodRubro Then
                        TvContable.Nodes(k).Text = TvContable.Nodes(k).Text & Space(48 - Len(TvContable.Nodes(k).Text) - Len(Format(VerificarNulo(!Importe, "N"), "#,##0.00"))) & Format(VerificarNulo(!Importe, "N"), "#,##0.00")
                        VecAgrupacionRubros(k).ImpContable = VerificarNulo(!Importe, "N")
                        TotalEgresos = TotalEgresos + !Importe
                        Exit For
                   End If
               Next
               .MoveNext
           Next
        End If
        
        .Close
        '******** Total Presupuestado ***********
        Sql = "SpOcConsultaPresupuestoFinancieroPresupuestado @Periodo=" & FechaSQL("01/" & Format(CalPeriodo, "MM/yyyy"), "SQL")
        .Open Sql, Conec
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
               For k = 1 To TvContable.Nodes.Count
                   If TvContable.Nodes(k).Tag = !CodRubro Then
                        TvContable.Nodes(k).Text = TvContable.Nodes(k).Text & Rellenar(Format(VerificarNulo(!Importe, "N"), "#,##0.00"), 17)
                        VecAgrupacionRubros(k).ImpPresupuestado = VerificarNulo(!Importe, "N")
                        Exit For
                   End If
               Next
               TotalPres = TotalPres + !Importe
               .MoveNext
           Next
        End If
        
        .Close
        '******** Total Financiero ***********
        Sql = "SpOcConsultaPresupuestoFinanciero @Periodo=" & FechaSQL("01/" & Format(CalPeriodo, "MM/yyyy"), "SQL")
        .Open Sql, Conec
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
               For k = 1 To TvContable.Nodes.Count
                   If TvContable.Nodes(k).Tag = !CodRubro Then
                        TvContable.Nodes(k).Text = TvContable.Nodes(k).Text & Rellenar(Format(VerificarNulo(!Importe, "N"), "#,##0.00"), 18)
                        VecAgrupacionRubros(k).ImpFinanciero = VerificarNulo(!Importe, "N")
                        Exit For
                   End If
               Next
               TotalFinanciero = TotalFinanciero + !Importe
               .MoveNext
           Next
        End If
   End With
    LbGOF = "Generación Operativa de Fondos: " & Format((GridListado.Rows.At(GridListado.Rows.Count - 1).Cells(2)) - (TotalEgresos * (1 + IvaCredito)), "#,##0.00")

    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal

End Sub

Private Sub CmdDetalleCont_Click()
    If Not TvContable.SelectedItem Is Nothing Then
        If TvContable.SelectedItem.Tag <> "  " Then
          '  A01_5910.Periodo = CalPeriodo.Value
            A01_5910.Rubro = TvContable.SelectedItem.Tag
            A01_5910.Show vbModal
        End If
    End If
End Sub

Private Sub CmdDetalleFinanciero_Click()
    If Not TvContable.SelectedItem Is Nothing Then
        If TvContable.SelectedItem.Tag <> "  " Then
         '   A01_5930.Periodo = CalPeriodo.Value
            A01_5930.Rubro = TvContable.SelectedItem.Tag
            A01_5930.Show vbModal
        End If
    End If
End Sub

Private Sub CmdDetallePres_Click()
    If Not TvContable.SelectedItem Is Nothing Then
        If TvContable.SelectedItem.Tag <> "  " Then
           ' A01_5920.Periodo = CalPeriodo.Value
            A01_5920.Rubro = TvContable.SelectedItem.Tag
            A01_5920.Show vbModal
        End If
    End If
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
Dim Fila As Integer
Dim Filas As Long
Dim i As Integer
Dim j As Integer
Dim Total As Double
Dim TotalPres As Double
Dim TotalFinanciero As Double

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
                ColorFondo = &HC0E0FF

        '-------- GENERO LOS DATOS ------------------------------
        '*** carga los ingresos
        
        .Range("A" & 5).Select
        .ActiveCell.FormulaR1C1 = "Ingresos"
        
        .Range("A" & 5).Font.Bold = True
        .Range("A" & 5).Select
        .Selection.HorizontalAlignment = xlCenter
      
        InicioTabla2 = GridListado.Rows.Count + 8
        Call EncabezadoExcelGrid(ex, Caption, 6, 4)
        Call DatosExcelGrid(ex, GridListado, 6, Filas)
        Call FormatearExcelGrid(ex, 6, Filas, 4, ColorFondo)
        
        '*** carga los Egresos
        .Range("A" & InicioTabla2 - 1).Font.Bold = True
        .Range("A" & InicioTabla2 - 1).Select
        .Selection.HorizontalAlignment = xlCenter
        .ActiveCell.FormulaR1C1 = "Egresos"

         Fila = InicioTabla2 + 1
         .Range("A" & Trim(Fila - 1) & ":F" & Trim(Fila - 1)).Select
         .Selection.Font.Bold = True
         .Range("A" & Trim(Fila - 1)).Value = "Rubro"
         .Range("B" & Trim(Fila - 1)).Value = "Imp. Contable"
         .Range("C" & Trim(Fila - 1)).Value = "Imp. Presupuestado"
         .Range("D" & Trim(Fila - 1)).Value = "Imp. Financiero"
         .Range("E" & Trim(Fila - 1)).Value = "Desvio Contable/Pres"
         .Range("F" & Trim(Fila - 1)).Value = "Desvio Financiero/Pres"

         Call AgregarNodoExcel(TvContable.Nodes(1), Fila, ex, Total, TotalPres, TotalFinanciero)
         Call FormatearExcelGrid(ex, InicioTabla2, Fila - InicioTabla2 - 1, 6, ColorFondo)
         .Range("A" & Fila).Select
         .ActiveCell.FormulaR1C1 = "Totales ==>"
         .Range("B" & Fila).Select
         .ActiveCell.FormulaR1C1 = Total
         .Range("C" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalPres
         .Range("D" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalFinanciero
         
         .Range("A" & Fila + 2).Select
         .ActiveCell.FormulaR1C1 = "Generación Operativa de Fondos:"
         .Range("B" & Fila + 2).Select
         .ActiveCell.FormulaR1C1 = GridListado.Rows.At(GridListado.Rows.Count - 1).Cells(2) - Total

        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To 6
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Periodo : " & Format(CalPeriodo, "MMMM/yyyy")
      
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub AgregarNodoExcel(Nodo As Node, Fila As Integer, ex As Excel.Application, Total As Double, TotalPres As Double, TotalFinanciero As Double)
Dim j As Integer
    
    If Not Nodo Is Nothing Then
        j = Nodo.Index
        If Nodo.Child Is Nothing Then
            ex.Range("A" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
            ex.Range("B" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpContable, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("B" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00"
            ex.Range("C" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpPresupuestado, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("C" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00"
            ex.Range("D" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpFinanciero, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("D" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00"
            ex.Range("E" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioContable, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("E" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00%"
            ex.Range("F" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioFinanciero, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("F" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00%"
            Total = Total + VecAgrupacionRubros(j).ImpContable
            TotalPres = TotalPres + VecAgrupacionRubros(j).ImpPresupuestado
            TotalFinanciero = TotalFinanciero + VecAgrupacionRubros(j).ImpFinanciero
            Fila = Fila + 1
            
        Else
            ex.Range("A" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
            Fila = Fila + 1
            Call AgregarNodoExcel(Nodo.Child, Fila, ex, Total, TotalPres, TotalFinanciero)

            ex.Range("A" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & "TOTAL " & VecAgrupacionRubros(j).A_Descripcion & " ==>"
            ex.Range("B" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpContable, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("C" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpPresupuestado, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("D" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpFinanciero, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("A" & Trim(Fila), "F" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00"
            ex.Selection.Font.Bold = True
            ex.Range("E" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioContable, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("E" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00%"
            ex.Range("F" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioFinanciero, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("F" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0.00%"

            
            Fila = Fila + 1
        End If
        
        Call AgregarNodoExcel(Nodo.Next, Fila, ex, Total, TotalPres, TotalFinanciero)
    End If
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim Total As Double
  Dim TotalPres As Double
  Dim TotalFinanciero As Double

  Dim RsListado As New ADODB.Recordset
    RsListado.Fields.Append "Tipo", adVarChar, 35
    RsListado.Fields.Append "Descripcion", adVarChar, 150
    RsListado.Fields.Append "Importe", adVarChar, 25
    RsListado.Fields.Append "Importe2", adVarChar, 25
    RsListado.Fields.Append "Importe3", adVarChar, 25
    RsListado.Fields.Append "Desvio", adVarChar, 25
    RsListado.Fields.Append "Desvio2", adVarChar, 25
    RsListado.Open
    i = 1
    '**** INGRESOS *******
    While i < GridListado.Rows.Count
      RsListado.AddNew
      With GridListado.Rows.At(i)
           RsListado!Tipo = "Ingresos"
           RsListado!Descripcion = .Cells(1)
           RsListado!Importe2 = .Cells(2)
           RsListado!Importe3 = .Cells(3)
           RsListado!Desvio = .Cells(4)
      End With
        i = i + 1
    Wend

   '******* EGRESOS ********
    Call AgregarNodoImpresion(TvContable.Nodes(1), RsListado, Total, TotalPres, TotalFinanciero)
    RsListado.AddNew
    RsListado!Tipo = "Egresos"
    RsListado!Descripcion = "TOTALES"
    RsListado!Importe = Format(Total, "#,##0.00")
    RsListado!Importe2 = Format(TotalPres, "#,##0.00")
    RsListado!Importe3 = Format(TotalFinanciero, "#,##0.00")
    
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    ListA01_5900.TxtFondo = LbGOF
    ListA01_5900.TxtPeriodo = Format(CalPeriodo, "MMMM/yyyy")
    ListA01_5900.DataControl1.Recordset = RsListado
    ListA01_5900.Zoom = -1
End Sub


Private Sub CmdImprimir_Click()
    Call ConfImpresionDeConsulta
    ListA01_5900.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV
    Call CalcularTotalesArboles
    Call CalcularDesvios
End Sub

Private Sub Form_Load()
    Dim i As Integer
    GridListado.DataRowCount = UBound(VecUnidadesDeNegocio) + 2
    For i = 1 To UBound(VecUnidadesDeNegocio)
        GridListado.Rows.At(i).Cells(1) = VecUnidadesDeNegocio(i).U_Descripcion
        GridListado.Rows.At(i).Cells(5) = VecUnidadesDeNegocio(i).U_Codigo
    Next
    GridListado.Rows.At(i).Cells(1) = "IVA Debito"
    GridListado.Rows.At(i + 1).Cells(1) = "Totales"
    GridListado.Rows.At(i + 1).Style.Font.Bold = True
    GridListado.Columns(2).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    
    Call CargarTv
    CalPeriodo.Value = Date
    Dim RsParametros As New ADODB.Recordset
    RsParametros.Open "SpCosTaParametros", Conec
    IvaCredito = RsParametros!P_PorcentajeIvaCredito / 100
End Sub

Private Sub CalcularDesvios()
    Dim i As Integer
    Dim j As Integer
    Dim Desvio As Double
    Dim Pres As Double
    Dim DesvioFinanciero As Double
    
    '********* Desvio Egresos **************
    For i = 1 To UBound(VecAgrupacionRubros)
        With VecAgrupacionRubros(i)
            If .ImpPresupuestado > 0 Then
                Pres = .ImpPresupuestado
                Desvio = (.ImpContable - Pres) / Pres
                DesvioFinanciero = (.ImpFinanciero - Pres) / Pres
            Else
                Desvio = 1
                DesvioFinanciero = 1
            End If
            .DesvioContable = Desvio
            .DesvioFinanciero = DesvioFinanciero
             '********* Contable Presupuestado **************
             TvContable.Nodes(i).Text = TvContable.Nodes(i).Text & Rellenar(Format(Desvio, "0.00%"), 17)
             '********* Finaciero Presupuestado **************
             TvContable.Nodes(i).Text = TvContable.Nodes(i).Text & Rellenar(Format(DesvioFinanciero, "0.00%"), 15)

        End With
    Next
        
    '********* Desvio Imgresos **************
    For i = 1 To GridListado.DataRowCount
        With GridListado.Rows.At(i)
            If ValN(FormatNumber(.Cells(2), 2, vbUseDefault, vbUseDefault, vbFalse)) > 0 Then
                Pres = ValN(FormatNumber(.Cells(2), 2, vbUseDefault, vbUseDefault, vbFalse))
                Desvio = (ValN(FormatNumber(VerificarNulo(.Cells(3), "N"), 2, vbUseDefault, vbUseDefault, vbFalse)) - Pres) / Pres
            Else
                Desvio = 1
            End If
            .Cells(4) = Format(Desvio, "0.00%")
        End With
    Next
    
End Sub

Private Sub CalcularTotalesArboles()
    Dim i As Integer
    Dim j As Integer
    Dim Total As Double
    Dim TotPres As Double
    Dim TotFinanciero As Double
 
        For i = TvContable.Nodes.Count To 1 Step -1
            Total = 0
            TotPres = 0
            TotFinanciero = 0
            If TvContable.Nodes(i).Children > 0 Then
                For j = 1 To UBound(VecAgrupacionRubros)
                    If VecAgrupacionRubros(i).A_Codigo = VecAgrupacionRubros(j).A_Padre Then
                        Total = Total + VecAgrupacionRubros(j).ImpContable
                        TotPres = TotPres + VecAgrupacionRubros(j).ImpPresupuestado
                        TotFinanciero = TotFinanciero + VecAgrupacionRubros(j).ImpFinanciero
                    End If
                Next
                TvContable.Nodes(i).Text = TvContable.Nodes(i).Text & Space(48 + (12 - 4 * VecAgrupacionRubros(i).A_Nivel) - Len(TvContable.Nodes(i).Text) - Len(Format(Total, "#,##0.00"))) & Format(Total, "#,##0.00")
                TvContable.Nodes(i).Text = TvContable.Nodes(i).Text & Rellenar(Format(TotPres, "#,##0.00"), 17)
                TvContable.Nodes(i).Text = TvContable.Nodes(i).Text & Rellenar(Format(TotFinanciero, "#,##0.00"), 18)
               
                VecAgrupacionRubros(i).ImpContable = Total
                TvContable.Nodes(i).Expanded = True
                TvContable.Nodes(i).BackColor = &HFFC0C0
                VecAgrupacionRubros(i).ImpPresupuestado = TotPres
                VecAgrupacionRubros(i).ImpFinanciero = TotFinanciero
            End If
        Next
End Sub

Private Sub CargarTv()
On Error GoTo ErrorCarga
Dim RsCargar As New ADODB.Recordset
Dim Sql As String
    Dim i As Integer
    With RsCargar
        ReDim VecAgrupacionRubros(0)
        Sql = "SpTaAgrupacionRubrosContables"
        .Open Sql, Conec
        TvContable.Nodes.Clear
        ReDim VecAgrupacionRubros(.RecordCount)
        For i = 1 To UBound(VecAgrupacionRubros)
            VecAgrupacionRubros(i).A_Codigo = !A_Codigo
            VecAgrupacionRubros(i).A_Descripcion = !A_Descripcion
            VecAgrupacionRubros(i).A_Nivel = !A_Nivel
            VecAgrupacionRubros(i).A_Padre = !A_Padre
            VecAgrupacionRubros(i).A_Rubro = !A_Rubro
            If !A_Padre <> 0 Then
                TvContable.Nodes.Add !A_Padre & "R", tvwChild, !A_Codigo & "R", !A_Descripcion
            Else
                TvContable.Nodes.Add , , !A_Codigo & "R", !A_Descripcion
            End If
            TvContable.Nodes(i).Tag = !A_Rubro
          
            .MoveNext
        Next
        CmdDetalleCont.Enabled = False
        CmdDetallePres.Enabled = False
        CmdDetalleFinanciero.Enabled = False
       
    End With
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub TvContable_Click()
    If Not TvContable.SelectedItem Is Nothing Then
        If TvContable.SelectedItem.Tag <> "  " Then
            CmdDetalleCont.Enabled = True
            CmdDetallePres.Enabled = True
            CmdDetalleFinanciero.Enabled = True
        Else
            CmdDetalleCont.Enabled = False
            CmdDetallePres.Enabled = False
            CmdDetalleFinanciero.Enabled = False
        End If
    End If
End Sub

Private Sub AgregarNodoImpresion(Nodo As Node, Rs As ADODB.Recordset, Total As Double, TotalPres As Double, TotalFinanciero As Double)
Dim j As Integer
    
    If Not Nodo Is Nothing Then
        j = Nodo.Index
        If Nodo.Child Is Nothing Then
            Rs.AddNew
            Rs!Tipo = "Egresos"
            Rs!Descripcion = VecAgrupacionRubros(j).A_Descripcion
            Rs!Importe = Format(VecAgrupacionRubros(j).ImpContable, "#,##0.00")
            Rs!Importe2 = Format(VecAgrupacionRubros(j).ImpPresupuestado, "#,##0.00")
            Rs!Importe3 = Format(VecAgrupacionRubros(j).ImpFinanciero, "#,##0.00")
            Rs!Desvio = Format(VecAgrupacionRubros(j).DesvioContable, "0.00%")
            Rs!Desvio2 = Format(VecAgrupacionRubros(j).DesvioFinanciero, "0.00%")
            Total = Total + VecAgrupacionRubros(j).ImpContable
            TotalPres = TotalPres + VecAgrupacionRubros(j).ImpPresupuestado
            TotalFinanciero = TotalFinanciero + VecAgrupacionRubros(j).ImpFinanciero
        Else
            Rs.AddNew
            Rs!Tipo = "Egresos"
            Rs!Descripcion = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
            Call AgregarNodoImpresion(Nodo.Child, Rs, Total, TotalPres, TotalFinanciero)
            Rs.AddNew
            Rs!Descripcion = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & "TOTAL " & VecAgrupacionRubros(j).A_Descripcion & " ==>"
            Rs!Importe = Format(VecAgrupacionRubros(j).ImpContable, "#,##0.00")
            Rs!Importe2 = Format(VecAgrupacionRubros(j).ImpPresupuestado, "#,##0.00")
            Rs!Importe3 = Format(VecAgrupacionRubros(j).ImpFinanciero, "#,##0.00")
            Rs!Desvio = Format(VecAgrupacionRubros(j).DesvioContable, "0.00%")
            Rs!Desvio2 = Format(VecAgrupacionRubros(j).DesvioFinanciero, "0.00%")
        End If
        
        Call AgregarNodoImpresion(Nodo.Next, Rs, Total, TotalPres, TotalFinanciero)
    End If
End Sub

