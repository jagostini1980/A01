VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_Requerimientos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos de Compra"
   ClientHeight    =   8010
   ClientLeft      =   -4710
   ClientTop       =   -1755
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkRechazados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir Rechazador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1665
      TabIndex        =   18
      Top             =   30
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mis Requeriminetos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   11
      Top             =   3945
      Width           =   11910
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   6060
         TabIndex        =   12
         Top             =   165
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   315
         Left            =   1860
         TabIndex        =   13
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   22675457
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   315
         Left            =   4650
         TabIndex        =   14
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   22675457
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin VB.Label LBFechaHasta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3435
         TabIndex        =   16
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Desde:"
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
         Left            =   555
         TabIndex        =   15
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Height          =   810
      Left            =   60
      TabIndex        =   4
      Top             =   3120
      Width           =   11910
      Begin VB.CommandButton CmdCrearOc 
         Caption         =   "Crear OC"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9870
         TabIndex        =   17
         Top             =   225
         Width           =   1150
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Motivo Rechazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2265
         TabIndex        =   9
         Top             =   150
         Width           =   1755
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F. Probable Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Value           =   -1  'True
         Width           =   2100
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8625
         TabIndex        =   6
         Top             =   225
         Width           =   1150
      End
      Begin VB.TextBox TxtMotivoRechazo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2265
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   380
         Width           =   6285
      End
      Begin MSComCtl2.DTPicker CalFechaProbEntrega 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   380
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   38993
      End
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   2835
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   11910
      _cx             =   21008
      _cy             =   5001
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
      StylesCollection=   $"A01_Requrimientos.frx":0000
      ColumnsCollection=   "A01_Requrimientos.frx":1DD9
      ValueItems      =   $"A01_Requrimientos.frx":7D2B
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   495
      Left            =   4237
      TabIndex        =   1
      Top             =   7455
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6082
      TabIndex        =   0
      Top             =   7455
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   7485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDSharpGrid2.SGGrid GridMisRequerimientos 
      Height          =   2835
      Left            =   60
      TabIndex        =   10
      Top             =   4560
      Width           =   11910
      _cx             =   21008
      _cy             =   5001
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
      StylesCollection=   $"A01_Requrimientos.frx":8110
      ColumnsCollection=   "A01_Requrimientos.frx":9EE9
      ValueItems      =   $"A01_Requrimientos.frx":F483
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Me Requieren"
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
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   1200
   End
End
Attribute VB_Name = "A01_Requerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const VgNumero = "#0.00" 'esta constante es el formato de los numeros

Private Sub InicializarTodo()
'On Error GoTo Errores
Dim i As Integer
Dim ColumnaActual As Integer
    MousePointer = vbHourglass
    CalFechaHasta.Value = Date
    CalFechaDesde.Value = DateAdd("M", -1, Date)
    CalFechaProbEntrega.Value = Date
    Call CmdTraer_Click
    MousePointer = vbNormal
    
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(6).Style.TextAlignment = sgAlignRightCenter
    
    GridMisRequerimientos.Columns(2).Style.TextAlignment = sgAlignRightCenter
    GridMisRequerimientos.Columns(5).Style.TextAlignment = sgAlignRightCenter
    GridMisRequerimientos.Columns(6).Style.TextAlignment = sgAlignRightCenter
    Call TraerPendientes
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub TraerPendientes()
On Error GoTo Errores
Dim RsCargar As New ADODB.Recordset
Dim Sql As String
Dim i As Integer
    With RsCargar
        Sql = "SpOcRequerimientosDeCompraTraerPendientes @CentroDeCostos='" & CentroEmisor & _
                                                     "', @Rechazados=" & IIf(ChkRechazados.Value, 1, 0)
        .Open Sql, Conec
        GridListado.DataRowCount = .RecordCount
        For i = 1 To .RecordCount
            GridListado.Rows.At(i).Cells(1) = False
            GridListado.Rows.At(i).Cells(2) = BuscarDescCentroEmisor(!R_CentroEmisor)
            GridListado.Rows.At(i).Cells(3) = !R_Numero
            GridListado.Rows.At(i).Cells(4) = IIf(!R_Prioridad = 0, "Normal", "Urgente")
            GridListado.Rows.At(i).Cells(5) = BuscarDescArt(!R_Articulo, BuscarTablaCentroEmisor(!R_CentroDestino))
            GridListado.Rows.At(i).Cells(6) = !R_CantidadPendiente
            GridListado.Rows.At(i).Cells(7) = VerificarNulo(!R_FechaProbableDeEntrega)
            GridListado.Rows.At(i).Cells(8) = VerificarNulo(!R_MotivoRechazo)
            GridListado.Rows.At(i).Cells(9) = ValN(!R_Articulo)
            .MoveNext
        Next
    End With
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub ChkRechazados_Click()
    Call TraerPendientes
End Sub

Private Sub CmdCrearOc_Click()
Dim i As Integer
    ReDim VecRequerimientoCompra(0)
    For i = 1 To GridListado.Rows.Count - 1
        If GridListado.Rows.At(i).Cells(1) Then
            ReDim Preserve VecRequerimientoCompra(UBound(VecRequerimientoCompra) + 1)
            VecRequerimientoCompra(UBound(VecRequerimientoCompra)).Cantidad = GridListado.Rows.At(i).Cells(6)
            VecRequerimientoCompra(UBound(VecRequerimientoCompra)).CodArticulo = GridListado.Rows.At(i).Cells(9)
            VecRequerimientoCompra(UBound(VecRequerimientoCompra)).DescArticulo = GridListado.Rows.At(i).Cells(5)
            VecRequerimientoCompra(UBound(VecRequerimientoCompra)).Numero = GridListado.Rows.At(i).Cells(3)
        End If
    Next
    Unload A01_2100
    A01_2100.Show
    A01_2100.CargarRequerimineto
End Sub

Private Sub CmdModif_Click()
Dim Rta As Integer
Dim i As Integer
Dim Sql As String

    Rta = MsgBox("¿Confirma Fecha de Entrega/Rechazo?", vbYesNo)
    If Rta = vbYes Then
        If OptTipo(0).Value Then
            GridListado.Rows.At(GridListado.Row).Cells(7) = CalFechaProbEntrega.Value
            GridListado.Rows.At(GridListado.Row).Cells(8) = ""
        Else
            GridListado.Rows.At(GridListado.Row).Cells(8) = TxtMotivoRechazo.Text
            GridListado.Rows.At(GridListado.Row).Cells(7) = ""
        End If
    
        i = GridListado.Row
        Sql = "SpOcRequerimientosDeCompraRenglones @R_Numero =" & GridListado.Rows.At(i).Cells(3) & _
                                                ", @R_Articulo =" & GridListado.Rows.At(i).Cells(9) & _
                                                ", @R_Rechazado =" & IIf(OptTipo(1).Value, 1, 0) & _
                                                ", @R_MotivoRechazo ='" & GridListado.Rows.At(i).Cells(8) & _
                                               "', @R_FechaProbableDeEntrega =" & IIf(OptTipo(0).Value, FechaSQL(CalFechaProbEntrega.Value, "SQL"), "Null")
        Conec.Execute Sql
        
        If GridListado.Row < GridListado.Rows.Count Then
            GridListado.Row = GridListado.Row + 1
        End If
    End If
    
End Sub

Private Sub CmdTraer_Click()
On Error GoTo Errores
Dim Sql As String
Dim i As Integer
Dim TbListado As New ADODB.Recordset

    MousePointer = vbHourglass
    TbListado.CursorLocation = adUseClient
    Sql = "SpOcRequerimientosDeCompraTraerDelCentroDeCosto @CentroDeCostos ='" & CentroEmisor & _
                                                       "', @FechaDesde =" & FechaSQL(CalFechaDesde, "SQL") & _
                                                       ",  @FechaHasta =" & FechaSQL(CalFechaHasta, "SQL")

    TbListado.Open Sql, Conec, adOpenKeyset
   
    With TbListado
        GridMisRequerimientos.DataRowCount = .RecordCount
        For i = 1 To .RecordCount
            GridMisRequerimientos.Rows.At(i).Cells(1) = BuscarDescCentroEmisor(!R_CentroDestino)
            GridMisRequerimientos.Rows.At(i).Cells(2) = !R_Numero
            GridMisRequerimientos.Rows.At(i).Cells(3) = IIf(!R_Prioridad = 0, "Normal", "Urgente")
            GridMisRequerimientos.Rows.At(i).Cells(4) = BuscarDescArt(!R_Articulo, BuscarTablaCentroEmisor(!R_CentroDestino))
            GridMisRequerimientos.Rows.At(i).Cells(5) = !R_Cantidad
            GridMisRequerimientos.Rows.At(i).Cells(6) = VerificarNulo(!R_NumeroOc)
            GridMisRequerimientos.Rows.At(i).Cells(7) = VerificarNulo(!R_FechaProbableDeEntrega)
            GridMisRequerimientos.Rows.At(i).Cells(8) = VerificarNulo(!R_MotivoRechazo)
            .MoveNext
        Next
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
        '.ActiveCell.FormulaR1C1 = "Periodo: " & Format(CalPeriodo.Value, "MM/yyyy")
        '.Range("C4").Select
        .ActiveCell.FormulaR1C1 = "Centro De Costo: " & BuscarDescCentroEmisor(CentroEmisor)
     
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 6, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub GridListado_Change()
Dim i As Integer
Dim CantArticulo As Integer
    For i = 1 To GridListado.Rows.Count - 1
        If GridListado.Rows.At(i).Cells(1) Then
            CantArticulo = CantArticulo + 1
        End If
    Next
    CmdCrearOc.Enabled = CantArticulo > 0
End Sub

Private Sub GridListado_SelChange(CancelSelect As Boolean)
   CmdModif.Enabled = True
   If GridListado.Rows.At(GridListado.Row).Cells(8) <> "" Then
        OptTipo(1).Value = True
        CalFechaProbEntrega.Value = Date
        TxtMotivoRechazo.Text = GridListado.Rows.At(GridListado.Row).Cells(8)
   Else
        OptTipo(0).Value = True
        If GridListado.Rows.At(GridListado.Row).Cells(7) = "" Then
           CalFechaProbEntrega.Value = Date
        Else
           CalFechaProbEntrega.Value = GridListado.Rows.At(GridListado.Row).Cells(7)
        End If
        TxtMotivoRechazo.Text = ""
   End If
End Sub

Private Sub OptTipo_Click(Index As Integer)
    CalFechaProbEntrega.Enabled = Index = 0
    TxtMotivoRechazo.Enabled = Index = 1
End Sub
