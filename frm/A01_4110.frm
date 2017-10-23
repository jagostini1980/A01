VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_4110 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución de Presupuestos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmpAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   7350
      TabIndex        =   19
      Top             =   5235
      Width           =   1150
   End
   Begin VB.TextBox TxtTotalPres 
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
      Left            =   8445
      TabIndex        =   17
      Text            =   "0"
      Top             =   4485
      Width           =   1275
   End
   Begin VB.Frame FrameAsig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación"
      Height          =   885
      Left            =   7380
      TabIndex        =   9
      Top             =   1260
      Width           =   2355
      Begin VB.TextBox TxtMonto 
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   465
         Width           =   870
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1005
         TabIndex        =   3
         Top             =   430
         Width           =   1150
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
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   8460
      TabIndex        =   8
      Text            =   "0"
      Top             =   4890
      Width           =   1275
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   8580
      TabIndex        =   4
      Top             =   5235
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Presupuesta"
      Height          =   1200
      Left            =   45
      TabIndex        =   5
      Top             =   30
      Width           =   9675
      Begin VB.TextBox TxtCuentaContable 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   15
         Top             =   795
         Width           =   4755
      End
      Begin VB.TextBox TxtCentroDeCostoEmisor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   14
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
         Format          =   22675459
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   6
         Top             =   270
         Width           =   1665
      End
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   4305
      Left            =   30
      TabIndex        =   16
      Top             =   1275
      Width           =   7290
      _cx             =   12859
      _cy             =   7594
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
      StylesCollection=   $"A01_4110.frx":0000
      ColumnsCollection=   $"A01_4110.frx":1DD9
      ValueItems      =   $"A01_4110.frx":310F
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total Pres.:"
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
      Left            =   7380
      TabIndex        =   18
      Top             =   4530
      Width           =   1005
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
      Left            =   7875
      TabIndex        =   7
      Top             =   4935
      Width           =   510
   End
End
Attribute VB_Name = "A01_4110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JerarquiaCentro As String
Private Modificado As Boolean
Public Ok As Boolean
Public NroPresupuesto As Integer

Private Sub CmdCerra_Click()
    Ok = False
    Unload Me
End Sub

Private Sub CmdModif_Click()
'On Error GoTo errores
Dim i As Integer
    
 'If ValidarCargaPresupuesto Then
       i = GridListado.Row
       'VecDistribucionPresupuesto(i).P_Importe = Val(TxtMonto)
       'VecDistribucionPresupuesto(i).P_SubCentroDeCosto = GridListado.Rows.At(i).Cells(3)
       GridListado.Rows.At(i).Cells(2) = Format(Val(TxtMonto), "0.00")
 'End If
  
  'calcula el total de la orden
   Call CalcularTotal
  
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmpAceptar_Click()
    Dim i As Integer
    Dim j As Integer
    
    If ValN(txtTotal) <> ValN(TxtTotalPres) Then
        MsgBox "El Importe presupuestado no está asignado correctamente", vbInformation
        Exit Sub
    End If
    
    ReDim VecDistribucionPresupuesto(GridListado.DataRowCount)
    For i = 1 To UBound(VecDistribucionPresupuesto)
        With VecDistribucionPresupuesto(i)
            .P_Importe = ValN(GridListado.Rows.At(i).Cells(2))
            .P_CuentaContable = "5121"
            .P_SubCentroDeCosto = GridListado.Rows.At(i).Cells(3)
            .P_NumeroPresupuesto = Val(TxtNroPresupuesto)
        End With
    Next
    Ok = True
    Unload Me
End Sub

Private Sub Form_Load()
   ' CentroEmisorActual = CentroEmisor
    Call CargarCentrosDeCostos
    Call CargarImportes
    GridListado.Columns(2).Style.TextAlignment = sgAlignRightCenter
End Sub

Private Sub CargarImportes()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To UBound(VecDistribucionPresupuesto)
        For j = 1 To GridListado.DataRowCount
            With GridListado.Rows.At(j)
                If .Cells(3) = VecDistribucionPresupuesto(i).P_SubCentroDeCosto Then
                    .Cells(2) = VecDistribucionPresupuesto(i).P_Importe
                    Exit For
                End If
            End With
        Next
    Next
    Call CalcularTotal
End Sub

Private Function ValidarCargaPresupuesto() As Boolean
    ValidarCargaPresupuesto = True
    Dim i As Integer

    If Val(TxtMonto.Text) = 0 Then
        MsgBox "Debe ingresar un Monto mayor que 0"
        TxtMonto.SetFocus
        ValidarCargaPresupuesto = False
    End If
End Function

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To GridListado.DataRowCount
        Total = Total + Val(Replace(GridListado.Rows.At(i).Cells(2), ",", "."))
    Next
       
    txtTotal.Text = Format(Total, "0.00")

End Sub

Private Sub GridListado_SelChange(CancelSelect As Boolean)
    TxtMonto.Text = Replace(GridListado.Rows.At(GridListado.Row).Cells(2), ",", ".")
    CmdModif.Enabled = True
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtMonto, KeyAscii)
End Sub

Private Sub CargarCentrosDeCostos()
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim ModificarPeriodo As Boolean
    Dim RsCargar As New ADODB.Recordset
  
On Error GoTo ErrorCarga
    
  ModificarPeriodo = True
  With RsCargar
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    
    Sql = "SpOcPresupuestosCentrosMovilidad @CentroDeCostoEmisor ='" & JerarquiaCentro & "'"
    .Open Sql, Conec
    
    If .EOF Then
        MsgBox "No existen Centor de Costos para " & TxtCuentaContable, vbInformation
        Exit Sub
    End If
   
    GridListado.DataRowCount = .RecordCount
    For i = 1 To .RecordCount
        GridListado.Rows.At(i).Cells(1) = !C_Descripcion
        GridListado.Rows.At(i).Cells(3) = !C_Codigo
        .MoveNext
   Next
End With
  Call CalcularTotal

ErrorCarga:
  Call ManipularError(Err.Number, Err.Description)
End Sub
