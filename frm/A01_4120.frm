VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_4120 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Mini Cenas"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmpAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   7620
      TabIndex        =   16
      Top             =   7620
      Width           =   1150
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación"
      Height          =   1320
      Left            =   60
      TabIndex        =   8
      Top             =   6660
      Width           =   7530
      Begin VB.TextBox TxtPrecioU 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3450
         TabIndex        =   26
         Top             =   915
         Width           =   870
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   315
         Left            =   1110
         TabIndex        =   24
         Top             =   915
         Width           =   870
      End
      Begin Controles.ComboEsp CmbLineas 
         Height          =   315
         Left            =   1110
         TabIndex        =   20
         Top             =   195
         Width           =   3900
         _ExtentX        =   6879
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
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6180
         TabIndex        =   19
         Top             =   915
         Width           =   1150
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6195
         TabIndex        =   18
         Top             =   540
         Width           =   1150
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   300
         Left            =   6195
         TabIndex        =   17
         Top             =   150
         Width           =   1150
      End
      Begin VB.TextBox TxtMonto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5235
         TabIndex        =   2
         Top             =   930
         Width           =   870
      End
      Begin Controles.ComboEsp CmbProveedoresMc 
         Height          =   315
         Left            =   1110
         TabIndex        =   22
         Top             =   562
         Width           =   3900
         _ExtentX        =   6879
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio Unitario:"
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
         Left            =   2070
         TabIndex        =   27
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad:"
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
         Left            =   255
         TabIndex        =   25
         Top             =   975
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor:"
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
         Left            =   135
         TabIndex        =   23
         Top             =   615
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Linea:"
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
         Left            =   540
         TabIndex        =   21
         Top             =   255
         Width           =   540
      End
      Begin VB.Label LbCant 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   5235
         TabIndex        =   9
         Top             =   705
         Width           =   645
      End
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   8730
      TabIndex        =   7
      Text            =   "0"
      Top             =   7080
      Width           =   1275
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   8865
      TabIndex        =   3
      Top             =   7605
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Presupuesta"
      Height          =   1200
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Width           =   9990
      Begin VB.TextBox TxtCuentaContable 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   14
         Top             =   795
         Width           =   4755
      End
      Begin VB.TextBox TxtCentroDeCostoEmisor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   4590
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   4005
         TabIndex        =   1
         Top             =   210
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MMMM/yyyy"
         Format          =   53477379
         UpDown          =   -1  'True
         CurrentDate     =   38980
      End
      Begin VB.TextBox TxtNroPresupuesto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1890
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta Contable"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   570
         Width           =   1425
      End
      Begin VB.Label Label2 
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
         Left            =   3210
         TabIndex        =   11
         Top             =   285
         Width           =   750
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
         Left            =   135
         TabIndex        =   10
         Top             =   555
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Presupuesto:"
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
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1665
      End
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   5370
      Left            =   15
      TabIndex        =   15
      Top             =   1275
      Width           =   10005
      _cx             =   17648
      _cy             =   9472
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
      DataColCount    =   7
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
      MultiSelect     =   1
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
      StylesCollection=   $"A01_4120.frx":0000
      ColumnsCollection=   "A01_4120.frx":1DD9
      ValueItems      =   $"A01_4120.frx":69AF
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   8145
      TabIndex        =   6
      Top             =   7125
      Width           =   510
   End
End
Attribute VB_Name = "A01_4120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JerarquiaCentro As String
Private Modificado As Boolean
Public Ok As Boolean
Public NroPresupuesto As Integer


Private Sub CmbProveedoresMc_Click()
    TxtPrecioU.Text = VecProvMinicenas(CmbProveedoresMc.ListIndex).P_Precio
End Sub

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
    If Validar Then
        Dim Index As Integer
        Index = GridListado.Row - 1
        GridListado.DataRowCount = GridListado.DataRowCount + 1
        GridListado.Array(Index, 0) = CmbLineas.Text
        GridListado.Array(Index, 1) = CmbProveedoresMc.Text
        GridListado.Array(Index, 2) = ValN(TxtCantidad)
        GridListado.Array(Index, 3) = ValN(TxtPrecioU)
        GridListado.Array(Index, 4) = ValN(TxtCantidad) * ValN(Me.TxtPrecioU)
        GridListado.Array(Index, 5) = VecLineas(CmbLineas.ListIndex).Codigo
        GridListado.Array(Index, 6) = VecProvMinicenas(CmbProveedoresMc.ListIndex).P_Codigo
        Call CalcularTotal
        GridListado.Row = GridListado.DataRowCount
    End If
    
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub


Private Function Validar() As Boolean
    Validar = True
    If CmbLineas.ListIndex = 0 Then
        Validar = False
        CmbLineas.SetFocus
        Call MsgBox("Debe Selecccionar una línea", vbInformation)
        Exit Function
    End If
    
    If CmbProveedoresMc.ListIndex = 0 Then
        Validar = False
        CmbProveedoresMc.SetFocus
        Call MsgBox("Debe Selecccionar un Proveedor", vbInformation)
        Exit Function
    End If
    
    If Val(TxtCantidad.Text) = 0 Then
        Validar = False
        TxtCantidad.SetFocus
        Call MsgBox("Debe Ingresar una Cantidad", vbInformation)
        Exit Function
    End If
    
End Function

Private Sub CmdCerra_Click()
    Ok = False
    Unload Me
End Sub

Private Sub CmdEliminar_Click()
   Call GridListado.Rows.RemoveAt(GridListado.Row)
   Call CalcularTotal
End Sub

Private Sub CmdModif_Click()
On Error GoTo Errores
    If Validar Then
        Dim Index As Integer
        Index = GridListado.Row - 1
        'GridListado.DataRowCount = GridListado.DataRowCount + 1
        GridListado.Array(Index, 0) = CmbLineas.Text
        GridListado.Array(Index, 1) = CmbProveedoresMc.Text
        GridListado.Array(Index, 2) = ValN(TxtCantidad)
        GridListado.Array(Index, 3) = ValN(TxtPrecioU)
        GridListado.Array(Index, 4) = ValN(TxtCantidad) * ValN(Me.TxtPrecioU)
        GridListado.Array(Index, 5) = VecLineas(CmbLineas.ListIndex).Codigo
        GridListado.Array(Index, 6) = VecProvMinicenas(CmbProveedoresMc.ListIndex).P_Codigo

        GridListado.Row = GridListado.Row + 1
        Call CalcularTotal
    End If
  
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmpAceptar_Click()
    Dim i As Integer
    Dim j As Integer
       
    ReDim VecPresServicioDeBar(GridListado.DataRowCount - 1)
    For i = 1 To UBound(VecPresServicioDeBar)
        With VecPresServicioDeBar(i)
           .D_Cuenta = "5023"
           .D_NumeroPresupuesto = Val(TxtNroPresupuesto)
           .D_Cantidad = Val(GridListado.Rows.At(i).Cells(3))
           .D_PrecioUnitario = ValN(GridListado.Rows.At(i).Cells(4))
           .D_Linea = GridListado.Rows.At(i).Cells(6)
           .D_Proveedor = GridListado.Rows.At(i).Cells(7)
           .D_Periodo = Format(CalPeriodo, "MM/yyyy")
        End With
    Next
    A01_4100.TxtMonto = Replace(ValN(txtTotal), ",", ".")
    Ok = True
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    Call CargarComboLineas(CmbLineas)
    Call CargarCmbProv(CmbProveedoresMc)
    GridListado.DataRowCount = UBound(VecPresServicioDeBar) + 1
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    
    For i = 1 To UBound(VecPresServicioDeBar)
        GridListado.Rows.At(i).Cells(1) = BuscarDescLineas(VecPresServicioDeBar(i).D_Linea)
        GridListado.Rows.At(i).Cells(2) = BuscarDescProvMinicena(VecPresServicioDeBar(i).D_Proveedor)
        GridListado.Rows.At(i).Cells(3) = VecPresServicioDeBar(i).D_Cantidad
        GridListado.Rows.At(i).Cells(4) = VecPresServicioDeBar(i).D_PrecioUnitario
        GridListado.Rows.At(i).Cells(5) = VecPresServicioDeBar(i).D_Cantidad * VecPresServicioDeBar(i).D_PrecioUnitario
        GridListado.Rows.At(i).Cells(6) = VecPresServicioDeBar(i).D_Linea
        GridListado.Rows.At(i).Cells(7) = VecPresServicioDeBar(i).D_Proveedor
    Next
    Call CalcularTotal
End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To GridListado.DataRowCount
        Total = Total + Val(Replace(GridListado.Rows.At(i).Cells(5), ",", "."))
    Next
       
    txtTotal.Text = Replace(Format(Total, "0.00"), ",", ".")

End Sub

Private Sub GridListado_SelChange(CancelSelect As Boolean)
    Call CargarEnModificar(GridListado.Row)
End Sub

Private Sub CargarEnModificar(Indice As Integer)
On Error GoTo Errores
    If Indice < GridListado.DataRowCount Then
        TxtCantidad = ValN(GridListado.Rows.At(Indice).Cells(3))
        TxtPrecioU = ValN(GridListado.Rows.At(Indice).Cells(4))
        TxtMonto = ValN(GridListado.Rows.At(Indice).Cells(5))
        Call UbicarCmbLineas(CmbLineas, GridListado.Rows.At(Indice).Cells(6))
        Call UbicarCmbProvMinicenas(CmbProveedoresMc, GridListado.Rows.At(Indice).Cells(7))
        CmdAgregar.Enabled = False
        CmdEliminar.Enabled = True
        CmdModif.Enabled = True
    Else
        CmbLineas.ListIndex = 0
        CmbProveedoresMc.ListIndex = 0
        TxtCantidad = ""
        TxtPrecioU = ""
        TxtMonto = ""
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
        CmdModif.Enabled = False
       
    End If
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub TxtCantidad_Change()
    TxtMonto = ValN(TxtCantidad.Text) * VecProvMinicenas(CmbProveedoresMc.ListIndex).P_Precio
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtCantidad, KeyAscii)
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtMonto, KeyAscii)
End Sub
