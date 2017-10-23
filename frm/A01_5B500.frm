VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_5B500 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Egresos Finacieros por Empresa"
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
   Begin DDSharpGrid2.SGGrid GridEgresos 
      Height          =   6795
      Left            =   45
      TabIndex        =   9
      Top             =   645
      Width           =   11805
      _cx             =   20823
      _cy             =   11986
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
      DataColCount    =   10
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
      StylesCollection=   $"A01_5B500.frx":0000
      ColumnsCollection=   "A01_5B500.frx":1DD9
      ValueItems      =   $"A01_5B500.frx":85B3
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6825
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5217
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   5310
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   51183617
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalHasta 
         Height          =   315
         Left            =   3675
         TabIndex        =   7
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   51183617
         CurrentDate     =   38940
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta:"
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
         Left            =   3015
         TabIndex        =   8
         Top             =   240
         Width           =   570
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
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   1200
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
Attribute VB_Name = "A01_5B500"
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

  Dim TotalFinanciero As Double
  
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   MousePointer = vbHourglass

On Error GoTo ErrorTraer:
    For k = 0 To GridEgresos.DataRowCount - 1
        GridEgresos.Array(k, 2) = ""
        GridEgresos.Array(k, 3) = ""
        GridEgresos.Array(k, 4) = ""
        GridEgresos.Array(k, 5) = ""
        GridEgresos.Array(k, 6) = ""
    Next
        '******** Total Egresos ***********
        '******** Total Real Financiero ***********
        Sql = "SpOcConsultaFinancieroPorEmpresa @FechaDesde=" & FechaSQL(CalDesde, "SQL") & _
                                              ", @FechaHasta=" & FechaSQL(CalHasta, "SQL")
    With RsListado
        .Open Sql, Conec
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
               For k = 0 To GridEgresos.DataRowCount - 2
                   If GridEgresos.Array(k, 9) = !C_Cuenta Then
                         Select Case !C_Empresa
                         Case "ELP"
                             GridEgresos.Array(k, 2) = Format(!Importe, "#,##0")
                         Case "LTO"
                              GridEgresos.Array(k, 3) = Format(!Importe, "#,##0")
                         Case "PPK"
                              GridEgresos.Array(k, 4) = Format(!Importe, "#,##0")
                         Case "VYM"
                             GridEgresos.Array(k, 5) = Format(!Importe, "#,##0")
                         End Select
                         'TotalFinanciero = TotalFinanciero + !Importe
                         GridEgresos.Array(k, 6) = Format(ValN(Replace(GridEgresos.Array(k, 6), ".", "")) + !Importe, "#,##0")
                         Exit For
                   End If
               Next
               
               .MoveNext
           Next
        End If
   End With
    CmdExp.Enabled = True
    GridEgresos.Redraw sgRedrawAll
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
Dim InicioTabla2 As Integer
Dim Fila As Integer
Dim Filas As Long
Dim i As Integer
Dim j As Integer
Dim Rubro As String
Dim TotalELP As Double
Dim TotalLTO As Double
Dim TotalPPK As Double
Dim TotalVyM As Double
Dim TotalGral As Double

Dim Arr As SGArray

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        ColorFondo = &HC0E0FF

        '-------- GENERO LOS DATOS ------------------------------
        
        .Range("A" & 6).Font.Bold = True
        .Range("A" & 6).Select
        .Selection.HorizontalAlignment = xlCenter
         InicioTabla2 = 6
        'InicioTabla2 = GridListado.Rows.Count + 8
        'Call EncabezadoExcelGrid(ex, Caption, 6, 4)
        'Call DatosExcelGrid(ex, GridListado, 6, Filas)
        'Call FormatearExcelGrid(ex, 6, Filas, 4, ColorFondo)
        
        '*** carga los Egresos
        .Range("A" & InicioTabla2 - 1).Font.Bold = True
        .Range("A" & InicioTabla2 - 1).Select
        .Selection.HorizontalAlignment = xlCenter
        .ActiveCell.FormulaR1C1 = "Egresos"

         Fila = InicioTabla2 + 1
         .Range("A" & Trim(Fila - 1) & ":F" & Trim(Fila - 1)).Select
         .Selection.Font.Bold = True
         .Range("A" & Trim(Fila - 1)).Value = "Rubro"
         .Range("B" & Trim(Fila - 1)).Value = "El Pulqui"
         .Range("C" & Trim(Fila - 1)).Value = "La Tosta"
         .Range("D" & Trim(Fila - 1)).Value = "P. Pack"
         .Range("E" & Trim(Fila - 1)).Value = "V y M"
         .Range("F" & Trim(Fila - 1)).Value = "Total"
        
         Set Arr = GridEgresos.Array

        For j = 1 To UBound(VecAgrupacionRubros)
            If VecAgrupacionRubros(j).A_Nivel = 1 Then
                Rubro = VecAgrupacionRubros(j).A_Descripcion
                .Range("A" & Trim(Fila)).Value = Rubro
                Fila = Fila + 1
                For i = 0 To Arr.RowCount - 1
                   If Rubro = Arr(i, 0) Then
                       .Range("A" & Trim(Fila)).Value = Space(10) & Arr(i, 1)
                       .Range("B" & Trim(Fila)).Value = ValN(Replace(Arr(i, 2), ".", ""))
                       .Range("C" & Trim(Fila)).Value = ValN(Replace(Arr(i, 3), ".", ""))
                       .Range("D" & Trim(Fila)).Value = ValN(Replace(Arr(i, 4), ".", ""))
                       .Range("E" & Trim(Fila)).Value = ValN(Replace(Arr(i, 5), ".", ""))
                       .Range("F" & Trim(Fila)).Value = ValN(Replace(Arr(i, 6), ".", ""))
                       If (Trim(Arr(i, 7)) <> "" And Trim(Arr(i, 9)) = "") Or Trim(Arr(i, 8)) <> "" Then
                          If Trim(Arr(i, 7)) <> "" Then
                             .Range("A" & Trim(Fila) & ":F" & Trim(Fila)).Font.Color = &H800000    'Azul
                          Else
                             .Range("A" & Trim(Fila) & ":F" & Trim(Fila)).Font.Color = &H4000&       'vende
                          End If
                          .Range("A" & Trim(Fila) & ":F" & Trim(Fila)).Font.Bold = True 'GridEgresos.Rows.At(i).Style.Font.Bold
                       End If
                       If Mid(Arr(i, 1), 1, 5) = "Total" Then
                           .Range("A" & Trim(Fila)).Value = Trim(.Range("A" & Trim(Fila)).Value)
                           .Range("A" & Trim(Fila) & ":F" & Trim(Fila)).Font.Bold = True
                           TotalELP = TotalELP + ValN(FormatNumber(Arr(i, 2), 2, vbUseDefault, vbUseDefault, vbFalse))
                           TotalLTO = TotalLTO + ValN(FormatNumber(Arr(i, 3), 2, vbUseDefault, vbUseDefault, vbFalse))
                           TotalPPK = TotalPPK + ValN(FormatNumber(Arr(i, 4), 2, vbUseDefault, vbUseDefault, vbFalse))
                           TotalVyM = TotalVyM + ValN(FormatNumber(Arr(i, 5), 2, vbUseDefault, vbUseDefault, vbFalse))
                           TotalGral = TotalGral + ValN(FormatNumber(Arr(i, 6), 2, vbUseDefault, vbUseDefault, vbFalse))
                        End If

                       Fila = Fila + 1
                   End If
                Next
            End If
        Next

         Call FormatearExcelGrid(ex, InicioTabla2, Fila - InicioTabla2 - 1, 6, ColorFondo)
         .Range("A" & Fila).Select
         .ActiveCell.FormulaR1C1 = "Totales ==>"
         .Range("B" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalELP
         .Range("C" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalLTO
         .Range("D" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalPPK
         .Range("E" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalVyM
         .Range("F" & Fila).Select
         .ActiveCell.FormulaR1C1 = TotalGral
         
         .Range("A" & Fila + 1).Select
         .ActiveCell.FormulaR1C1 = "Porcentajes ==>"
         .Range("B" & Fila + 1).Select
         .ActiveCell.FormulaR1C1 = TotalELP / TotalGral
         .Range("C" & Fila + 1).Select
         .ActiveCell.FormulaR1C1 = TotalLTO / TotalGral
         .Range("D" & Fila + 1).Select
         .ActiveCell.FormulaR1C1 = TotalPPK / TotalGral
         .Range("E" & Fila + 1).Select
         .ActiveCell.FormulaR1C1 = TotalVyM / TotalGral
         .Range("B" & Trim(Fila + 1) & ":E" & Trim(Fila + 1)).Select
         .Selection.NumberFormat = "0.0%"
         '.Range("F" & Fila).Select
         '.ActiveCell.FormulaR1C1 = TotalGral
         
         .Range("A" & Fila + 2).Select
         
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
        .ActiveCell.FormulaR1C1 = "Fecha Desde " & CalDesde & " Hasta " & CalHasta.Value
      
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
            ex.Selection.NumberFormat = "0"
            ex.Range("C" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpPresupuestado, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("C" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0"
            ex.Range("D" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).ImpFinanciero, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("D" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0"
            ex.Range("E" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioContable, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("E" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0%"
            ex.Range("F" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioFinanciero, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("F" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0%"
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
            ex.Selection.NumberFormat = "0"
            ex.Selection.Font.Bold = True
            ex.Range("E" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioContable, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("E" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0%"
            ex.Range("F" & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(VecAgrupacionRubros(j).DesvioFinanciero, 4, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
            ex.Range("F" & Trim(Fila)).Select
            ex.Selection.NumberFormat = "0%"
            
            Fila = Fila + 1
        End If
        
        Call AgregarNodoExcel(Nodo.Next, Fila, ex, Total, TotalPres, TotalFinanciero)
    End If
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim j As Integer
  Dim Total As Double
  Dim TotalPres As Double
  Dim TotalFinanciero As Double
  Dim Arr As SGArray
  Dim Rubro As String
  
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

   '******* EGRESOS ********

    Set Arr = GridEgresos.Array

    For j = 1 To UBound(VecAgrupacionRubros)
        If VecAgrupacionRubros(j).A_Nivel = 1 Then
            Rubro = VecAgrupacionRubros(j).A_Descripcion
            RsListado.AddNew
            RsListado!Tipo = "Egresos"
            RsListado!Descripcion = Rubro
            
            For i = 0 To Arr.RowCount - 1
               If Rubro = Arr(i, 0) Then
                   RsListado.AddNew
                   RsListado!Tipo = "Egresos"
                   RsListado!Descripcion = Space(10) & Arr(i, 1)
                   RsListado!Importe = Format(Arr(i, 2), "#,##0")
                   RsListado!Importe2 = Format(Arr(i, 3), "#,##0")
                   RsListado!Importe3 = Format(Arr(i, 4), "#,##0")
                   RsListado!Desvio = Arr(i, 5)
                   RsListado!Desvio2 = Arr(i, 6)
                   If Mid(Arr(i, 1), 1, 5) = "Total" Then
                      RsListado!Descripcion = Trim(RsListado!Descripcion)
                      Total = Total + ValN(FormatNumber(Arr(i, 2), 2, vbUseDefault, vbUseDefault, vbFalse))
                      TotalPres = TotalPres + ValN(FormatNumber(Arr(i, 3), 2, vbUseDefault, vbUseDefault, vbFalse))
                      TotalFinanciero = TotalFinanciero + ValN(FormatNumber(Arr(i, 4), 2, vbUseDefault, vbUseDefault, vbFalse))
                   End If
               End If
            Next
           
        End If
    Next
    
    RsListado.AddNew
    RsListado!Tipo = "Egresos"
    RsListado!Descripcion = "TOTALES"
    RsListado!Importe = Format(Total, "#,##0")
    RsListado!Importe2 = Format(TotalPres, "#,##0")
    RsListado!Importe3 = Format(TotalFinanciero, "#,##0")
    
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    'ListA01_5900.TxtFondo = LbGOF
    ListA01_5900.TxtPeriodo = "Fecha Desde " & CalDesde.Value & " Hasta " & CalHasta
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
    If CalDesde.Year <> CalHasta.Year Or CalDesde.Month <> CalHasta.Month Then
        MsgBox "El Rango está en diferentes meses", vbInformation
        Exit Sub
    End If
    
    Call CargarLV
    Call CalcularTotalesArboles
    Call CalcularTotalesRubro
    'Call CalcularDesvios
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    GridEgresos.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridEgresos.Columns(4).Style.TextAlignment = sgAlignRightCenter
    GridEgresos.Columns(5).Style.TextAlignment = sgAlignRightCenter
    GridEgresos.Columns(6).Style.TextAlignment = sgAlignRightCenter
    GridEgresos.Columns(7).Style.TextAlignment = sgAlignRightCenter
    'GridEgresos.Groups.Add "ColNivel1"
    Call CargarTv
    CalDesde.Value = Date
    CalDesde.Day = 1
    CalHasta.Value = Date
    'Dim RsParametros As New ADODB.Recordset
    'RsParametros.Open "SpCosTaParametros", Conec
    'IvaCredito = RsParametros!P_PorcentajeIvaCredito / 100
End Sub

Private Sub CalcularTotalesArboles()
Dim i As Integer
Dim j As Integer
Dim TotalELP As Double
Dim TotLTO As Double
Dim TotPPK As Double
Dim TotVyM As Double
Dim TotGeneral As Double
Dim Rubro As String
Dim Arr As SGArray
Dim IndexTotal As Integer

    Set Arr = GridEgresos.Array
    Rubro = Arr(0, 0)
    For j = 1 To UBound(VecAgrupacionRubros)
        If VecAgrupacionRubros(j).A_Nivel = 1 Then
            TotalELP = 0
            TotLTO = 0
            TotPPK = 0
            TotGeneral = 0
            TotVyM = 0
            Rubro = VecAgrupacionRubros(j).A_Descripcion
            For i = 0 To Arr.RowCount - 1
                If Rubro = Arr(i, 0) Then
                    TotalELP = TotalELP + ValN(Replace(Arr(i, 2), ".", ""))
                    TotLTO = TotLTO + ValN(Replace(Arr(i, 3), ".", ""))
                    TotPPK = TotPPK + ValN(Replace(Arr(i, 4), ".", ""))
                    TotVyM = TotVyM + ValN(Replace(Arr(i, 5), ".", ""))
                    TotGeneral = TotGeneral + ValN(Replace(Arr(i, 6), ".", ""))
                    If Arr(i, 1) = "Total " & Rubro Then
                       IndexTotal = i
                    End If
               End If
            Next
            Arr(IndexTotal, 2) = Format(TotalELP, "#,##0")
            Arr(IndexTotal, 3) = Format(TotLTO, "#,##0")
            Arr(IndexTotal, 4) = Format(TotPPK, "#,##0")
            Arr(IndexTotal, 5) = Format(TotVyM, "#,##0")
            Arr(IndexTotal, 6) = Format(TotGeneral, "#,##0")
        End If
    Next
End Sub

Private Sub CalcularTotalesRubro()
Dim i As Integer
Dim j As Integer
Dim TotalELP As Double
Dim TotLTO As Double
Dim TotPPK As Double
Dim TotVyM As Double
Dim TotGeneral As Double
Dim RubroTotalELP As Double
Dim RubroTotLTO As Double
Dim RubroTotPPK As Double
Dim RubroTotVyM As Double
Dim Rubro As String
Dim Arr As SGArray
Dim IndexTotal As Integer

    Set Arr = GridEgresos.Array
    Rubro = Arr(0, 0)
    For j = 1 To UBound(VecAgrupacionRubros)
        VecAgrupacionRubros(j).TotELP = 0
        VecAgrupacionRubros(j).TotLTO = 0
        VecAgrupacionRubros(j).TotPPK = 0
        VecAgrupacionRubros(j).TotVyM = 0
        'Debug.Print VecAgrupacionRubros(j).A_Nivel & " - " & VecAgrupacionRubros(j).A_Descripcion & " - R" & VecAgrupacionRubros(j).A_Codigo & " - P " & VecAgrupacionRubros(j).A_Padre
        If VecAgrupacionRubros(j).A_Nivel = 3 Then
            TotalELP = 0
            TotLTO = 0
            TotPPK = 0
            TotGeneral = 0
            TotVyM = 0
            Rubro = VecAgrupacionRubros(j).A_Codigo
            If VecAgrupacionRubros(j).A_Rubro <> "  " Then
                For i = 0 To Arr.RowCount - 1
                    If Arr(i, 7) = VecAgrupacionRubros(j).A_Rubro And _
                       Arr(i, 9) <> "" Then
                        TotalELP = TotalELP + ValN(Replace(Arr(i, 2), ".", ""))
                        TotLTO = TotLTO + ValN(Replace(Arr(i, 3), ".", ""))
                        TotPPK = TotPPK + ValN(Replace(Arr(i, 4), ".", ""))
                        TotVyM = TotVyM + ValN(Replace(Arr(i, 5), ".", ""))
                        TotGeneral = TotGeneral + ValN(Replace(Arr(i, 6), ".", ""))
                        'If Arr(i, 9) = "" Then
                         '  IndexTotal = i
                        'End If
                    End If
                    
                    If Arr(i, 7) = VecAgrupacionRubros(j).A_Rubro And Arr(i, 9) = "" Then
                       IndexTotal = i
                    End If
                Next
                
                VecAgrupacionRubros(j).TotELP = TotalELP
                VecAgrupacionRubros(j).TotLTO = TotLTO
                VecAgrupacionRubros(j).TotPPK = TotPPK
                VecAgrupacionRubros(j).TotVyM = TotVyM
                Arr(IndexTotal, 2) = Format(TotalELP, "#,##0")
                Arr(IndexTotal, 3) = Format(TotLTO, "#,##0")
                Arr(IndexTotal, 4) = Format(TotPPK, "#,##0")
                Arr(IndexTotal, 5) = Format(TotVyM, "#,##0")
                Arr(IndexTotal, 6) = Format(TotGeneral, "#,##0")
            End If
        End If
    Next
    '######## Totales en Verde ##########
    For j = 1 To UBound(VecAgrupacionRubros)
        RubroTotalELP = 0
        RubroTotLTO = 0
        RubroTotPPK = 0
        RubroTotVyM = 0
        If VecAgrupacionRubros(j).A_Nivel = 2 Then
            For i = 1 To UBound(VecAgrupacionRubros)
                If Val(VecAgrupacionRubros(j).A_Codigo) = VecAgrupacionRubros(i).A_Padre Then
                    RubroTotalELP = RubroTotalELP + VecAgrupacionRubros(i).TotELP
                    RubroTotLTO = RubroTotLTO + VecAgrupacionRubros(i).TotLTO
                    RubroTotPPK = RubroTotPPK + VecAgrupacionRubros(i).TotPPK
                    RubroTotVyM = RubroTotVyM + VecAgrupacionRubros(i).TotVyM
                End If
            Next
            
            For i = 0 To Arr.RowCount - 1
                If Arr(i, 8) = VecAgrupacionRubros(j).A_Codigo Then
                    Arr(i, 2) = Format(RubroTotalELP, "#,##0")
                    Arr(i, 3) = Format(RubroTotLTO, "#,##0")
                    Arr(i, 4) = Format(RubroTotPPK, "#,##0")
                    Arr(i, 5) = Format(RubroTotVyM, "#,##0")
                    Arr(i, 6) = Format(RubroTotalELP + RubroTotLTO + RubroTotPPK + RubroTotVyM, "#,##0")
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub CargarTv()
On Error GoTo ErrorCarga
Dim RsCargar As New ADODB.Recordset
Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    With RsCargar
        ReDim VecAgrupacionRubros(0)
        Sql = "SpTaAgrupacionRubrosContables"
        .Open Sql, Conec

        GridEgresos.DataRowCount = 0
        ReDim VecAgrupacionRubros(.RecordCount)
        For i = 1 To UBound(VecAgrupacionRubros)
            VecAgrupacionRubros(i).A_Codigo = !A_Codigo
            VecAgrupacionRubros(i).A_Descripcion = !A_Descripcion
            VecAgrupacionRubros(i).A_Nivel = !A_Nivel
            VecAgrupacionRubros(i).A_Padre = !A_Padre
            VecAgrupacionRubros(i).A_Rubro = !A_Rubro
           .MoveNext
        Next
    End With

    For i = 1 To UBound(VecAgrupacionRubros)
         With VecAgrupacionRubros(i)
            If Not Existe(.A_Codigo) Then
             If .A_Padre > 0 Then
                 GridEgresos.DataRowCount = GridEgresos.DataRowCount + 1
                 If .A_Rubro = "  " Then
                     GridEgresos.Array(j, 0) = BuscarDescRubroContable(.A_Padre)
                     GridEgresos.Array(j, 8) = ""
                     GridEgresos.Rows.At(j + 1).Style.Font.Bold = True
                     GridEgresos.Rows.At(j + 1).Style.ForeColor = &H4000&
                 Else
                     If BuscarNivelPadre(.A_Padre) = 1 Then
                         GridEgresos.Array(j, 0) = BuscarDescRubroContable(.A_Padre)
                         GridEgresos.Array(j, 8) = ""
                         GridEgresos.Rows.At(j + 1).Style.Font.Bold = True
                         GridEgresos.Rows.At(j + 1).Style.ForeColor = &H4000&
                     'Else
                         'GridEgresos.Array(j, 0) = BuscarDescRubroContable(VecAgrupacionRubros(1).A_Codigo)
                     End If
                 End If
                 GridEgresos.Array(j, 1) = Space((.A_Nivel - 2) * 8) & .A_Descripcion
                 GridEgresos.Array(j, 7) = .A_Rubro
                 GridEgresos.Array(j, 8) = .A_Codigo
                  j = j + 1
                 For k = i To UBound(VecAgrupacionRubros)
                    If .A_Codigo = VecAgrupacionRubros(k).A_Padre Then
                        GridEgresos.DataRowCount = GridEgresos.DataRowCount + 1
                        GridEgresos.Array(j, 0) = BuscarDescRubroContable(.A_Padre)
                        GridEgresos.Array(j, 1) = Space((VecAgrupacionRubros(k).A_Nivel - 2) * 8) & VecAgrupacionRubros(k).A_Descripcion
                        GridEgresos.Array(j, 7) = VecAgrupacionRubros(k).A_Rubro
                        GridEgresos.Array(j, 8) = VecAgrupacionRubros(k).A_Codigo
                        GridEgresos.Rows.At(j + 1).Style.Font.Bold = True
                        GridEgresos.Rows.At(j + 1).Style.ForeColor = &H800000
                        'j = j + 1
                        Call CargarCuentasRubro(VecAgrupacionRubros(k).A_Rubro, .A_Padre)
                        j = GridEgresos.DataRowCount
                    End If
                 Next
                
             End If
            End If
         End With
     Next
     
     For i = 1 To UBound(VecAgrupacionRubros)
        If VecAgrupacionRubros(i).A_Nivel = 1 Then
            GridEgresos.DataRowCount = GridEgresos.DataRowCount + 1
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 0) = BuscarDescRubroContable(VecAgrupacionRubros(i).A_Codigo)
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 1) = "Total " & BuscarDescRubroContable(VecAgrupacionRubros(i).A_Codigo)
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 7) = "  "
            GridEgresos.Rows.At(GridEgresos.DataRowCount).Style.Font.Bold = True
        End If
     Next
     
     'CmdDetalleCont.Enabled = False
     'CmdDetallePres.Enabled = False
     'CmdDetalleFinanciero.Enabled = False
     
     GridEgresos.Groups.Add "ColNivel0"

ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarCuentasRubro(Rubro As String, Padre As Integer)
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    Dim i As Integer
    
    With RsCargar
        Sql = "SpOcCuentasPorRubroTraer @Rubro='" & Rubro & "'"
        .Open Sql, Conec
        For i = 1 To .RecordCount
            GridEgresos.DataRowCount = GridEgresos.DataRowCount + 1
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 0) = BuscarDescRubroContable(Padre)
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 1) = Space(16) & Convertir(!DescCuenta) & " - Cod. " & !CodCuenta
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 7) = Rubro
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 8) = "" 'VecAgrupacionRubros(k).A_Codigo
            GridEgresos.Array(GridEgresos.DataRowCount - 1, 9) = !CodCuenta
           .MoveNext
        Next
    End With
End Sub

Private Function Existe(Codigo As Integer) As Boolean
Dim i As Integer
    Existe = False
    For i = 0 To GridEgresos.DataRowCount - 1
        If Codigo = Val(GridEgresos.Array(i, 8)) Then
            Existe = True
            Exit Function
        End If
    Next
End Function

Private Function BuscarDescRubroContable(Codigo As Integer) As String
Dim i As Integer
    For i = 1 To UBound(VecAgrupacionRubros)
        If Codigo = VecAgrupacionRubros(i).A_Codigo Then
            BuscarDescRubroContable = VecAgrupacionRubros(i).A_Descripcion
            Exit Function
        End If
    Next
End Function

Private Function BuscarNivelPadre(CodPadre As Integer) As Integer
Dim i As Integer
    For i = 1 To UBound(VecAgrupacionRubros)
        If CodPadre = VecAgrupacionRubros(i).A_Codigo Then
            BuscarNivelPadre = VecAgrupacionRubros(i).A_Nivel
            Exit Function
        End If
    Next
End Function

Private Sub AgregarNodoImpresion(Nodo As Node, Rs As ADODB.Recordset, Total As Double, TotalPres As Double, TotalFinanciero As Double)
Dim j As Integer
    
    If Not Nodo Is Nothing Then
        j = Nodo.Index
        If Nodo.Child Is Nothing Then
            Rs.AddNew
            Rs!Tipo = "Egresos"
            Rs!Descripcion = VecAgrupacionRubros(j).A_Descripcion
            Rs!Importe = Format(VecAgrupacionRubros(j).ImpContable, "#,##0")
            Rs!Importe2 = Format(VecAgrupacionRubros(j).ImpPresupuestado, "#,##0")
            Rs!Importe3 = Format(VecAgrupacionRubros(j).ImpFinanciero, "#,##0")
            Rs!Desvio = Format(VecAgrupacionRubros(j).DesvioContable, "0%")
            Rs!Desvio2 = Format(VecAgrupacionRubros(j).DesvioFinanciero, "0%")
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
            Rs!Importe = Format(VecAgrupacionRubros(j).ImpContable, "#,##0")
            Rs!Importe2 = Format(VecAgrupacionRubros(j).ImpPresupuestado, "#,##0")
            Rs!Importe3 = Format(VecAgrupacionRubros(j).ImpFinanciero, "#,##0")
            Rs!Desvio = Format(VecAgrupacionRubros(j).DesvioContable, "0%")
            Rs!Desvio2 = Format(VecAgrupacionRubros(j).DesvioFinanciero, "0%")
        End If
        
        Call AgregarNodoImpresion(Nodo.Next, Rs, Total, TotalPres, TotalFinanciero)
    End If
End Sub
