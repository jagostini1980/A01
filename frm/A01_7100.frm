VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{4313501F-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20.ocx"
Begin VB.Form A01_7100 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de Proveedores"
   ClientHeight    =   4515
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4456.259
   ScaleMode       =   0  'User
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2655
      TabIndex        =   14
      Top             =   4065
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1230
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   7215
      Begin VB.TextBox TxtProductos 
         Height          =   315
         Left            =   1110
         TabIndex        =   15
         Top             =   855
         Width           =   6030
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Calificar De Nuevo"
         Height          =   315
         Left            =   3420
         TabIndex        =   10
         Top             =   495
         Width           =   1740
      End
      Begin VB.TextBox TxtCalificacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1845
         TabIndex        =   9
         Top             =   495
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   315
         Left            =   5850
         TabIndex        =   0
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52232193
         CurrentDate     =   39865
      End
      Begin Controles.ComboEsp CmbProv 
         Height          =   330
         Left            =   1110
         TabIndex        =   6
         Top             =   157
         Width           =   4080
         _ExtentX        =   7197
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Productos:"
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
         TabIndex        =   16
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Última Calificación:"
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
         TabIndex        =   8
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   5310
         TabIndex        =   5
         Top             =   225
         Width           =   495
      End
   End
   Begin VB.CommandButton CMDGuardar 
      Appearance      =   0  'Flat
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4065
      Width           =   1455
   End
   Begin VB.CommandButton CMDSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5745
      TabIndex        =   3
      Top             =   4065
      Width           =   1455
   End
   Begin DDSharpGrid2.SGGrid GridListado 
      Height          =   1845
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   7215
      _cx             =   12726
      _cy             =   3254
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   6
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
      CellPadding     =   14.74
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
      TabKeyBehavior  =   2
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200.189
      DefaultRowHeight=   251.799
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
      GroupIndentation=   225.071
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
      GroupByBoxText  =   "Matriz de Calificación"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"A01_7100.frx":0000
      ColumnsCollection=   "A01_7100.frx":1DD9
      ValueItems      =   $"A01_7100.frx":69CD
   End
   Begin VB.Line Line3 
      X1              =   30
      X2              =   7260
      Y1              =   1243.607
      Y2              =   1243.607
   End
   Begin VB.Line Line2 
      X1              =   30
      X2              =   7245
      Y1              =   1510.094
      Y2              =   1510.094
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   7260
      Y1              =   1820.996
      Y2              =   1820.996
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "|    Muy Malo - 2 ptos.  |  Malo - 4 ptos.  |  Bueno - 6 ptos.  |  Muy Bueno - 8 Ptos.  |  Exelente 10 ptos.  |"
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   1590
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "TABLA DE CALIFIACIÓN DE FACTOSRES DE EVALUACIÓN"
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
      Left            =   30
      TabIndex        =   12
      Top             =   1305
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Matriz de Calificación"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1905
      Width           =   7215
   End
End
Attribute VB_Name = "A01_7100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NumeroEvaluacion As Integer
Private Nuevo As Boolean

Private Sub CmbProv_Click()
    Dim Sql As String
    Dim i As Integer
    Dim RsTraer As New ADODB.Recordset
    Dim Total As Double
    
      CmdGuardar.Enabled = CmbProv.ListIndex > 0
      GridListado.Enabled = CmbProv.ListIndex > 0
      CalFecha.Enabled = CmbProv.ListIndex > 0
      CmdNuevo.Enabled = CmbProv.ListIndex > 0
      TxtProductos.Enabled = CmbProv.ListIndex > 0
      
      Sql = "SpTaProveedoresEvaluacionCabeceraTraerUltimo @E_Proveedor=" & VecProveedores(CmbProv.ListIndex).Codigo
      With RsTraer
          .Open Sql, Conec
          If Not .EOF Then
              NumeroEvaluacion = !E_NumeroEvaluacion
              Nuevo = False
              CalFecha.Value = !E_Fecha
              TxtProductos.Text = !E_Productos
              
              Select Case !E_Calificacion
              Case 60 To 100
                  TxtCalificacion = "Aprobado"
              Case 40 To 59
                  TxtCalificacion = "Condicional"
              Case 1 To 39
                  TxtCalificacion = "No Comprar"
              Case 0
                  TxtCalificacion = "Nuevo"
              End Select
              .Close
              Sql = "SpTaProveedoresEvaluacionRenglonesTraer @E_NumeroEvaluacion=" & NumeroEvaluacion
              .Open Sql, Conec
              For i = 1 To .RecordCount
                  '!E_Criterio
                  'GridListado.Rows.At(i).Cells(3) = !E_Importacia
                  GridListado.Rows.At(i).Cells(4) = !E_Calificacion
                  GridListado.Rows.At(i).Cells(5) = !E_Calificacion * !E_Importacia / 10
                  Total = Total + !E_Calificacion * !E_Importacia / 10
                  .MoveNext
              Next
              GridListado.Rows.At(GridListado.Rows.Count - 1).Cells(5) = Total
              CmdImprimir.Enabled = True
          Else
              Call CmdNuevo_Click
          End If
      End With
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresion
    ListA01_7100.Show

End Sub

Private Sub ConfImpresion()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    RsListado.Fields.Append "Item", adInteger
    RsListado.Fields.Append "Factor", adVarChar, 100
    RsListado.Fields.Append "Importancia", adInteger
    RsListado.Fields.Append "Calificacion", adDouble
    RsListado.Fields.Append "Resultado", adDouble

    RsListado.Open
    i = 1
    For i = 1 To GridListado.Rows.Count - 2
        RsListado.AddNew
      With GridListado.Rows.At(i)
        RsListado!Item = i
        RsListado!Factor = .Cells(2)
        RsListado!Importancia = .Cells(3)
        RsListado!Calificacion = ValN(.Cells(4))
        RsListado!Resultado = ValN(.Cells(5))
      End With
    Next
    
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    ListA01_7100.TxtFecha = CalFecha.Value
    ListA01_7100.TxtProductos.Text = TxtProductos.Text
    ListA01_7100.TxtProveedor.Text = CmbProv.Text & " (Cod. " & VecProveedores(CmbProv.ListIndex).Codigo & ")"
    ListA01_7100.LbClase = TxtCalificacion
    'ListA01_7100.TxtAnulada.Text = LBAnulada.Caption
    'ListA01_7100.TxtObservaciones.Text = TxtObs.Text
    ListA01_7100.Zoom = -1
    ListA01_7100.DataControl1.Recordset = RsListado
End Sub

Private Sub CmdNuevo_Click()
    Dim i As Integer
    TxtCalificacion = ""
    CalFecha.Value = Date
    TxtProductos.Text = ""
    For i = 1 To GridListado.Rows.Count - 1
        GridListado.Rows.At(i).Cells(4) = ""
        GridListado.Rows.At(i).Cells(5) = ""
    Next
    Nuevo = True
    CmdImprimir.Enabled = False
End Sub

Private Sub Form_Load()
    Call CargarComboProveedores(CmbProv)
    CalFecha.Value = Date
   
    GridListado.Columns(1).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(3).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(4).Style.TextAlignment = sgAlignRightCenter
    GridListado.Columns(5).Style.TextAlignment = sgAlignRightCenter
    
    GridListado.Rows.At(1).Cells(1) = 1
    GridListado.Rows.At(1).Cells(2) = "Calidad de Producto"
    GridListado.Rows.At(1).Cells(3) = 30
    GridListado.Rows.At(2).Cells(1) = 2
    GridListado.Rows.At(2).Cells(2) = "Servicio Técnico y/o Posventa"
    GridListado.Rows.At(2).Cells(3) = 10
    GridListado.Rows.At(3).Cells(1) = 3
    GridListado.Rows.At(3).Cells(2) = "Plazo de Entrega"
    GridListado.Rows.At(3).Cells(3) = 20
    GridListado.Rows.At(4).Cells(1) = 4
    GridListado.Rows.At(4).Cells(2) = "Precio"
    GridListado.Rows.At(4).Cells(3) = 25
    GridListado.Rows.At(5).Cells(1) = 5
    GridListado.Rows.At(5).Cells(2) = "Condiciones de Pago"
    GridListado.Rows.At(5).Cells(3) = 15
    GridListado.Rows.At(6).Cells(2) = "Grado de Aceptabilidad"
    GridListado.Rows.At(6).Style.Font.Bold = True
    GridListado.Rows.At(6).Style.ForeColor = &HC00000
End Sub

Private Sub CmdSalir_Click()
Dim Pregunta As Integer
     Pregunta = MsgBox("¿Esta seguro que desea Salir?", vbQuestion + vbOKCancel, "El Pulqui")
     If Pregunta = vbOK Then
        Unload Me
     End If
End Sub

Private Sub CMDGuardar_Click()
    Call GuardarCambios
End Sub

Private Sub GuardarCambios()
On Error GoTo Errores
    Dim TbGuardar As New ADODB.Recordset
    Dim Sql As String
    Dim Numero As Integer
    Dim Pregunta As Integer
        
    Dim i As Integer
    
    Pregunta = MsgBox("¿Desea grabar los datos?", vbQuestion + vbOKCancel, "El Pulqui")
    If Pregunta = vbOK Then
            If GridListado.DataRowCount = 1 Then
                MsgBox "Debe Imgresar artículos", vbInformation
                Exit Sub
            End If
            
            Conec.BeginTrans
            If Nuevo Then
                
                Sql = "SpTaProveedoresEvaluacionCabeceraAgregar @E_Proveedor =" & VecProveedores(Me.CmbProv.ListIndex).Codigo & _
                                                             ", @E_Fecha =" & FechaSQL(CalFecha, "SQL") & _
                                                             ", @E_Calificacion =" & Val(GridListado.Rows.At(GridListado.DataRowCount).Cells(5)) & _
                                                             ", @E_Productos ='" & TxtProductos.Text & _
                                                             "',@E_Usuario='" & Usuario & "'"
                TbGuardar.Open Sql, Conec
                Numero = TbGuardar!E_NumeroEvaluacion
            Else
                'modifica la cabecera y borra los renglones
                Sql = "SpTaProveedoresEvaluacionCabeceraModificar @E_NumeroEvaluacion =" & NumeroEvaluacion & _
                                                               ", @E_Fecha =" & FechaSQL(CalFecha, "SQL") & _
                                                               ", @E_Calificacion =" & Val(GridListado.Rows.At(GridListado.DataRowCount).Cells(5)) & _
                                                               ", @E_Productos ='" & TxtProductos.Text & _
                                                               "',@E_Usuario='" & Usuario & "'"
                Conec.Execute Sql
                Numero = NumeroEvaluacion
            End If
                        
            For i = 1 To GridListado.DataRowCount - 1 'recorro todos los renglones
                
                With GridListado.Rows.At(i)
                   Sql = "SpTaProveedoresEvaluacionRenglonesAgregar @E_NumeroEvaluacion =" & Numero & _
                                                                 ", @E_Criterio =" & Val(.Cells(1)) & _
                                                                 ", @E_Importacia =" & Val(.Cells(3)) & _
                                                                 ", @E_Calificacion =" & Val(.Cells(4))
                   Conec.Execute Sql
                   
                End With
             Next
             
            Conec.CommitTrans
            MsgBox "Los datos fueron guardaron correctamente", vbInformation
            Call CmbProv_Click
        End If
     
Errores:
    If Err.Number <> 0 Then
        Conec.RollbackTrans
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub

Private Sub GridListado_KeyPressEdit(ByVal RowKey As Long, ByVal ColIndex As Long, KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub GridListado_ValidateEdit(ByVal RowKey As Long, ByVal ColIndex As Long, ByVal OldValue As Variant, Cancel As Boolean)
    Dim i As Integer
    Dim Total As Integer
    If Val(GridListado.Rows(RowKey).Cells(ColIndex)) > 10 Then
        GridListado.Rows(RowKey).Cells(ColIndex) = ""
        MsgBox "El valor debe estar en entre 0 y 10", vbInformation
       
    End If
    GridListado.Rows(RowKey).Cells(ColIndex + 1) = Replace(ValN(GridListado.Rows(RowKey).Cells(ColIndex)) * ValN(GridListado.Rows(RowKey).Cells(ColIndex - 1)) / 10, ",", ".")
    For i = 1 To GridListado.Rows.Count - 2
        Total = Total + ValN(GridListado.Rows.At(i).Cells(ColIndex + 1))
    Next
    GridListado.Rows.At(i).Cells(ColIndex + 1) = Total
End Sub
