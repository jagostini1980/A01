VERSION 5.00
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_4610 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anticipos"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   2565
      TabIndex        =   1
      Top             =   6855
      Width           =   1230
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   350
      Left            =   3960
      TabIndex        =   0
      Top             =   6855
      Width           =   1230
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   6705
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7650
      _cx             =   13494
      _cy             =   11827
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
      AutoResize      =   0
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
      StylesCollection=   $"A01_4610.frx":0000
      ColumnsCollection=   $"A01_4610.frx":1DD9
      ValueItems      =   $"A01_4610.frx":3A93
   End
End
Attribute VB_Name = "A01_4610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Aceptar As Boolean
Public Destino As Integer
Public Proveedor As Integer
Public FileNro As Long
Public TotalAnticipo As Double

Private Sub CmdAceptar_Click()
Dim i As Integer
    Aceptar = True
    TotalAnticipo = 0
    ReDim VecAutorizacionesAnticiposApli(0)
    For i = 1 To GridListado.Rows.Count - 1
        If GridListado.Rows.At(i).Cells(1) Then
            ReDim Preserve VecAutorizacionesAnticiposApli(UBound(VecAutorizacionesAnticiposApli) + 1)
            VecAutorizacionesAnticiposApli(UBound(VecAutorizacionesAnticiposApli)) = Val(GridListado.Rows.At(i).Cells(2))
            TotalAnticipo = TotalAnticipo + ValN(GridListado.Rows.At(i).Cells(4))
        End If
    Next
    Visible = False
End Sub

Private Sub CmdCancelar_Click()
    Aceptar = False
    Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim RsCargar As New ADODB.Recordset
Dim i As Integer
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    If FileNro = 0 Then
        Sql = "SpOcAutorizacionesDePagoCabeceraAnticiposSinAplicarDestino @Destino = " & Destino & _
                                                                       ", @CodProveedor =" & Proveedor
    Else
        Sql = "SpOcAutorizacionesDePagoCabeceraAnticiposSinAplicarFile @FileNro = " & FileNro & _
                                                                    ", @CodProveedor =" & Proveedor
    End If
    
    With RsCargar
        .Open Sql, Conec
        GridListado.DataRowCount = .RecordCount
        For i = 1 To .RecordCount
            GridListado.Rows.At(i).Cells(1) = False
            GridListado.Rows.At(i).Cells(2) = !A_NumeroDeAutorizacionDePago
            GridListado.Rows.At(i).Cells(3) = !A_Fecha
            GridListado.Rows.At(i).Cells(4) = !Importe
            GridListado.Rows.At(i).Cells(5) = !A_Observaciones
            .MoveNext
        Next
    End With
End Sub

