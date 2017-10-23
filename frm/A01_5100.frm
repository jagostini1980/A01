VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por Cuenta Contable - Artículo"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   585
      Left            =   105
      TabIndex        =   16
      Top             =   6345
      Visible         =   0   'False
      Width           =   11850
      _cx             =   20902
      _cy             =   1032
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
      DataColCount    =   11
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
      StylesCollection=   $"A01_5100.frx":0000
      ColumnsCollection=   "A01_5100.frx":1DD9
      ValueItems      =   $"A01_5100.frx":8F7D
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   885
      Top             =   6990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   350
      Left            =   7935
      TabIndex        =   17
      Top             =   7440
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   7020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5100.frx":9362
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5100.frx":967C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5100.frx":BE2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   11865
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Traer"
         Height          =   315
         Left            =   6480
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   405
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   2430
         TabIndex        =   0
         Top             =   255
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM/yyyy"
         Format          =   109379587
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   2430
         TabIndex        =   1
         Top             =   675
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo Emisor:"
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
         Left            =   270
         TabIndex        =   7
         Top             =   750
         Width           =   2055
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
         Left            =   1575
         TabIndex        =   6
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10665
      TabIndex        =   4
      Top             =   7425
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   9315
      TabIndex        =   3
      Top             =   7425
      Width           =   1230
   End
   Begin MSComctlLib.TreeView TVListado 
      Height          =   5325
      Left            =   105
      TabIndex        =   8
      Top             =   1620
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9393
      _Version        =   393217
      Indentation     =   794
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desvío"
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
      Left            =   9720
      TabIndex        =   15
      Top             =   1305
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desvío %"
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
      Left            =   10890
      TabIndex        =   14
      Top             =   1305
      Width           =   825
   End
   Begin VB.Label LBReal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Real: $"
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
      Height          =   195
      Left            =   6345
      TabIndex        =   13
      Top             =   7020
      Width           =   1125
   End
   Begin VB.Label LBPres 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Presupuestado: $"
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
      Left            =   9000
      TabIndex        =   12
      Top             =   7020
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Presup."
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
      Left            =   8190
      TabIndex        =   11
      Top             =   1305
      Width           =   1470
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Real"
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
      Height          =   195
      Left            =   6705
      TabIndex        =   10
      Top             =   1305
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuenta Contable - Artículo"
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
      Left            =   270
      TabIndex        =   9
      Top             =   1305
      Width           =   6375
   End
End
Attribute VB_Name = "A01_5100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoCuenta
    CuentaContable As String
    DescCta As String
    R_Total As Double
    P_Total As Double
End Type

Private Type TipoArticulo
    Articulo As String
    CodigoArt As String
    CodCuenta As String
    CentroEmisor As String
    SubCentro As String
    OrdenNro As String
    Precio As Double
    Cantidad As Double
    Usuario As String
End Type

Private Cuenta() As TipoCuenta
Private Articulos() As TipoArticulo
Private Nivel As Integer
Private AlterColor As Boolean

Private Sub CmdCerra_Click()
    Unload Me
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
     
        ColorFondo = &HC0E0FF
        Call FormatearExcelGrid(ex, 7, GridListado.DataRowCount, Columnas, ColorFondo)
    End With
    Call GuardarPlanillaGrid(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5200.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
  
    RsListado.Fields.Append "CodCuenta", adVarChar, 6
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "Total", adDouble
    RsListado.Fields.Append "TotalPres", adDouble

    RsListado.Open
    i = 1
    While i <= UBound(Cuenta)
        RsListado.AddNew
        With Cuenta(i)
          RsListado!CodCuenta = .CuentaContable
          RsListado!Cuenta = .DescCta
          RsListado!Total = .R_Total
          RsListado!TotalPres = .P_Total
        End With
        i = i + 1
    Wend
    
    For i = 1 To UBound(Articulos)
        RsListado.AddNew
        With Articulos(i)
            RsListado!CodCuenta = .CodCuenta
            RsListado!Cuenta = "        " & .Articulo
            RsListado!Total = .Precio
            'RsListado!TotalPres = .P_Total
        End With
    Next
    
    RsListado.MoveFirst
    RsListado.Sort = "CodCuenta"
    
    ListA01_5200.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5200.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5200.DataControl1.Recordset = RsListado
    ListA01_5200.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
    Call CargarTreeView(CalPeriodo.Value)
    Call CalcularTotales
    CmdImprimir.Enabled = TVListado.Nodes.Count > 0
End Sub

Private Sub CargarTreeView(Periodo As Date)
Dim Item As String
Dim Sql As String
Dim i As Integer
Dim j As Integer
Dim CodCuenta As String
Dim Artuculo As String
Dim Desvio As Double
Dim Total As String
Dim TotalPres As String

Dim RsCargar As ADODB.Recordset
Set RsCargar = New ADODB.Recordset

'On Error GoTo Error
    
    TVListado.Nodes.Clear
    'pone el príodo en el primer día del mes
    Periodo = "01/" + CStr(Format(Periodo, "MM/yyyy"))
    
    Sql = "SpOCConsultaPorCuentaTraerCuentas " + _
                   "@PeriodoCta = " + FechaSQL(CStr(Periodo), "SQL") + _
                 ", @PeriodoPres = '" + CStr(Format(Periodo, "MM/yyyy")) + _
                "', @CentroDeCostoEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
    'trae las cuantas de ese período
    RsCargar.Open Sql, Conec
  With RsCargar
      ReDim Cuenta(.RecordCount)
      
        i = 1
    While Not .EOF
    'carga las cuentas en un vertor para una eventual impresión
        Cuenta(i).CuentaContable = !CuentaContable
        Cuenta(i).DescCta = BuscarDescCta(!CuentaContable)
        Cuenta(i).R_Total = VerificarNulo(!O_Total, "N")
        Cuenta(i).P_Total = VerificarNulo(!P_Total, "N")
        
     'agrega el artículo en el tv sin totales
        Desvio = Cuenta(i).R_Total - Cuenta(i).P_Total
        Item = Cuenta(i).DescCta + Space(85 - Len(Cuenta(i).DescCta) - Len(Format(Cuenta(i).R_Total, "0.00"))) + Format(Cuenta(i).R_Total, "0.00")
        Item = Item + Space(102 - Len(Item) - Len(Format(Cuenta(i).P_Total, "0.00"))) + Format(Cuenta(i).P_Total, "0.00")
        Item = Item + Space(115 - Len(Item) - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00")
        If Cuenta(i).P_Total = 0 Then
           Desvio = 1
        Else
            Desvio = Desvio / Cuenta(i).P_Total
        End If
        Item = Item + Space(124 - Len(Item) - Len(Format(Desvio, "0.00%"))) + Format(Desvio, "0.00%")

        TVListado.Nodes.Add , , CStr(!CuentaContable) + "C", Item, 1
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = &HFFC0C0
        i = i + 1
        .MoveNext
    Wend
        .Close
        
    Sql = "SpOCConsultaPorCuentaTraerArticulo " + _
            "@Periodo =" + FechaSQL(CStr(Periodo), "SQL") + "," + _
            "@CentroDeCostoEmisor = '" + CStr(VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo) & "'"
    'trae los artículos que participan del período
     
     .Open Sql, Conec
     i = 1
     ReDim Articulos(.RecordCount)
     
     If Not .EOF Then
        CodCuenta = !R_CuentaContable
     End If
    While Not .EOF
        
        If !R_CodigoArticulo = 0 Then
            Artuculo = "Servicio"
        Else
            Artuculo = BuscarDescArt(!R_CodigoArticulo, BuscarTablaCentroEmisor(!O_CentroDeCostoEmisor))
        End If
        Articulos(i).Cantidad = ValN(!CantidadRecibida)
        Articulos(i).CodigoArt = ValN(!R_CodigoArticulo)
        Articulos(i).Articulo = Artuculo
        Articulos(i).CodCuenta = !R_CuentaContable
        Articulos(i).Precio = VerificarNulo(!Total, "N")
        Articulos(i).SubCentro = !R_CentroDeCosto
        Articulos(i).OrdenNro = !Numero
        Articulos(i).Usuario = !Usuario
        
        Total = VerificarNulo(!Total, "N")
        Desvio = Total - Val(TotalPres)
        Item = Artuculo + Space(80 - Len(Artuculo) - Len(Format(Total, "0.00"))) + Format(Total, "0.00")
      'agraga el Tv el nodo con la Cta - centro que componen el artículo
        TVListado.Nodes.Add CStr(!R_CuentaContable) + "C", tvwChild, , Item, 2
        'TVListado.Nodes(TVListado.Nodes.Count).BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
        'AlterColor = Not AlterColor
        i = i + 1
      .MoveNext
    Wend

        .Close
  End With
   ' Para Exportar a Excel
   GridListado.DataRowCount = 0
    For j = 1 To UBound(Cuenta)
        GridListado.DataRowCount = GridListado.DataRowCount + 1
        GridListado.Rows.At(GridListado.DataRowCount).Cells(1).Value = Cuenta(j).DescCta & " (Cod. " & Cuenta(j).CuentaContable & ")"
        GridListado.Rows.At(GridListado.DataRowCount).Cells(6) = Cuenta(j).P_Total
        GridListado.Rows.At(GridListado.DataRowCount).Cells(7) = Cuenta(j).R_Total - Cuenta(j).P_Total
        If Cuenta(j).P_Total <> 0 Then
            GridListado.Rows.At(GridListado.DataRowCount).Cells(8) = Format((Cuenta(j).R_Total - Cuenta(j).P_Total) / Cuenta(j).P_Total, "0.00%")
        Else
            GridListado.Rows.At(GridListado.DataRowCount).Cells(8) = Format(1, "0.00%")
        End If
        For i = 1 To UBound(Articulos)
            If Articulos(i).CodCuenta = Cuenta(j).CuentaContable Then
                GridListado.DataRowCount = GridListado.DataRowCount + 1
                GridListado.Rows.At(GridListado.DataRowCount).Cells(2) = Articulos(i).CodigoArt
                GridListado.Rows.At(GridListado.DataRowCount).Cells(3) = Articulos(i).Articulo
                GridListado.Rows.At(GridListado.DataRowCount).Cells(4) = Articulos(i).Cantidad
                GridListado.Rows.At(GridListado.DataRowCount).Cells(5) = Articulos(i).Precio
                GridListado.Rows.At(GridListado.DataRowCount).Cells(9) = BuscarDescCentro(Articulos(i).SubCentro)
                GridListado.Rows.At(GridListado.DataRowCount).Cells(10) = Articulos(i).OrdenNro
                GridListado.Rows.At(GridListado.DataRowCount).Cells(11) = Articulos(i).Usuario
            End If
        Next
    Next
  
    ' Fin Exportar

   TVListado.Sorted = True
    
   For i = 1 To TVListado.Nodes.Count
        TVListado.Nodes(i).Expanded = True
        TVListado.Nodes(i).Sorted = True
   Next
   
   Dim Nodo
   i = 1
   If TVListado.Nodes.Count > 0 Then
    While i < TVListado.Nodes.Count
         Set Nodo = TVListado.Nodes(i).Child
         AlterColor = True
         While Not Nodo Is Nothing
         'pone la alteración de colores
            Nodo.BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
            AlterColor = Not AlterColor
            Set Nodo = Nodo.Next
         Wend
         
         i = i + 1
     Wend
    End If
Error:
    Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub Form_Load()
   
    CalPeriodo.Value = Date
        
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    
    Nivel = TraerNivel("A015100")
    If Nivel = 2 Then
        CmbCentroDeCostoEmisor.ListIndex = 0
    Else
        Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
        CmbCentroDeCostoEmisor.Enabled = False
        CmdTraer.Enabled = CentroEmisor <> ""
        If CentroEmisor = "" Then
            MsgBox "Ud. No tiene un Centro de Costos Asociado", vbInformation
        End If
    End If

End Sub

Private Sub CalcularTotales()
  Dim i As Integer
  Dim Total As Double
  Dim TotalPres As Double
  
    For i = 1 To UBound(Cuenta)
        Total = Total + Cuenta(i).R_Total
        TotalPres = TotalPres + Cuenta(i).P_Total
    Next
        LBReal.Caption = "Total Real: $" + Format(Total, "0.00")
        LBPres.Caption = "Total Presupuestado: $" + Format(TotalPres, "0.00")
End Sub

