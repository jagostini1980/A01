VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_1B200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrupación Rubros Contables"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   5940
      Top             =   7335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   1215
      TabIndex        =   11
      Top             =   7470
      Width           =   1300
   End
   Begin MSComctlLib.TreeView TreeRubros 
      Height          =   5910
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   10425
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.Frame FrameServ 
      BackColor       =   &H00E0E0E0&
      Height          =   1410
      Left            =   37
      TabIndex        =   8
      Top             =   5985
      Width           =   6540
      Begin Controles.ComboEsp CmbRubros 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   540
         Width           =   4800
         _ExtentX        =   8467
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
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1260
         TabIndex        =   3
         Top             =   945
         Width           =   1300
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1575
         MaxLength       =   30
         TabIndex        =   1
         Top             =   180
         Width           =   4800
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2655
         TabIndex        =   4
         Top             =   945
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4050
         TabIndex        =   5
         Top             =   945
         Width           =   1300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rubro Contable:"
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
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label LbCta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción:"
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
         Left            =   405
         TabIndex        =   9
         Top             =   225
         Width           =   1080
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Height          =   350
      Left            =   2617
      TabIndex        =   6
      Top             =   7470
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4057
      TabIndex        =   7
      Top             =   7470
      Width           =   1320
   End
End
Attribute VB_Name = "A01_1B200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private VecAgrupacionRubros() As TipoAgrupacionRubroContrable
Private xNodo As Node

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = TreeRubros.SelectedItem.Index
      ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
      Modificado = True
     'lo pone en el LV
      VecAgrupacionRubros(i).A_Descripcion = TxtDescripcion.Text
      VecAgrupacionRubros(i).A_Rubro = VecRubrosContables(CmbRubros.ListIndex).Codigo
      VecAgrupacionRubros(i).Estado = "A"
      VecAgrupacionRubros(i).A_Nivel = NivelNodo(TreeRubros.Nodes(i))
      If Not TreeRubros.SelectedItem.Parent Is Nothing Then
        VecAgrupacionRubros(i).A_Padre = VecAgrupacionRubros(TreeRubros.SelectedItem.Parent.Index).A_Codigo
      End If
      
      TreeRubros.Nodes(i).Text = TxtDescripcion.Text
      TreeRubros.Nodes(i).Checked = True
      
      If TreeRubros.SelectedItem.Parent Is Nothing Then
         TreeRubros.Nodes.Add , , , "Nuevo"
      Else
        ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
        TreeRubros.Nodes.Add TreeRubros.SelectedItem.Parent, 4, , "Nuevo"
      End If
      
      If CmbRubros.ListIndex = 0 Then
        ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
        TreeRubros.Nodes.Add TreeRubros.SelectedItem, 4, , "Nuevo"
      End If
      
      TreeRubros.Nodes(i + 1).Selected = True
      Call TreeRubros_NodeClick(TreeRubros.SelectedItem)
      TreeRubros.Nodes(i).Expanded = True
      Modificado = True
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Function ValidarCarga() As Boolean
Dim i As Integer
Dim NodoPadre As Node
Dim NodoHijo As Node

        ValidarCarga = True
        If TxtDescripcion.Text = "" Then
            ValidarCarga = False
            MsgBox "Debe Ingresar La Descripcion", vbInformation
            TxtDescripcion.SetFocus
            Exit Function
        End If
        
        Set NodoPadre = TreeRubros.SelectedItem.Parent
        If Not NodoPadre Is Nothing Then
            If CmbRubros.ListIndex = 0 Then
               For i = 1 To TreeRubros.Nodes.Count
                    If Not TreeRubros.Nodes(i).Parent Is Nothing Then
                        If i <> TreeRubros.Nodes(TreeRubros.SelectedItem.Index).Index And _
                           TreeRubros.Nodes(i).Text <> "Nuevo" And _
                           TreeRubros.Nodes(i).Parent = NodoPadre And _
                           Trim(VecAgrupacionRubros(TreeRubros.Nodes(i).Index).A_Rubro) <> "" Then
                                ValidarCarga = False
                                MsgBox "Debe Seleccionar un Rubro", vbInformation
                                CmbRubros.SetFocus
                                Exit Function
                        End If
                    End If
               Next
            Else
               For i = 1 To TreeRubros.Nodes.Count
                    If Not TreeRubros.Nodes(i).Parent Is Nothing Then
                        If i <> TreeRubros.Nodes(TreeRubros.SelectedItem.Index).Index And _
                           TreeRubros.Nodes(i).Text <> "Nuevo" And _
                           TreeRubros.Nodes(i).Parent = NodoPadre And _
                           Trim(VecAgrupacionRubros(TreeRubros.Nodes(i).Index).A_Rubro) = "" Then
                                ValidarCarga = False
                                MsgBox "No Debe Seleccionar un Rubro", vbInformation
                                CmbRubros.ListIndex = 0
                                Exit Function
                        End If
                    End If
               Next
            End If
        End If
        
        For i = 1 To UBound(VecAgrupacionRubros)
          If TreeRubros.SelectedItem.Index <> i Then
            'If VecAgrupacionRubros(i).A_Descripcion = TxtDescripcion Then
            '    ValidarCarga = False
            '    MsgBox "La Descripción ya fue Cargarda", vbInformation
            '    TxtDescripcion.SetFocus
            '    Exit Function
            'End If
            If CmbRubros.ListIndex > 0 Then
                If VecAgrupacionRubros(i).A_Rubro = VecRubrosContables(CmbRubros.ListIndex).Codigo Then
                    ValidarCarga = False
                    MsgBox "El Rubro Contable ya está en uso", vbInformation
                    CmbRubros.SetFocus
                    Exit Function
                End If
            End If
          End If
        Next
End Function

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirnar_Click()
    Dim Rta As Integer
    Rta = MsgBox("¿Desea Guardar los Datos?", vbYesNo)
    If Rta = vbYes Then
        Call Grabar
    End If
End Sub

Private Sub Grabar()
'On Error GoTo ErrorGrabar
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer
    Dim NodoHijo As Node
    Dim RsAgregar As New ADODB.Recordset
    Conec.BeginTrans

    For i = 1 To TreeRubros.Nodes.Count - 1
        With VecAgrupacionRubros(i)
            If Not TreeRubros.Nodes(i).Checked Then
                .Estado = "B"
            End If
            
            Select Case .Estado
            Case "A"
                Sql = "SpTaAgrupacionRubrosContablesAgregar @A_Descripcion ='" & .A_Descripcion & _
                                                        "', @A_Padre =" & .A_Padre & _
                                                        " , @A_Nivel =" & .A_Nivel & _
                                                        " , @A_Rubro ='" & .A_Rubro & "'"
                RsAgregar.Open Sql, Conec
                If TreeRubros.Nodes(i).Children > 0 Then
                   For j = 1 To TreeRubros.Nodes.Count
                       If Not TreeRubros.Nodes(j).Parent Is Nothing Then
                            If TreeRubros.Nodes(i) = TreeRubros.Nodes(j).Parent Then
                              VecAgrupacionRubros(j).A_Padre = RsAgregar!A_Codigo
                            End If
                       End If
                   Next
                End If
                RsAgregar.Close
            Case "M"
                Sql = "SpTaAgrupacionRubrosContablesModificar @A_Codigo =" & .A_Codigo & _
                                                          " , @A_Descripcion='" & .A_Descripcion & _
                                                          "', @A_Rubro='" & .A_Rubro & "'"
                Conec.Execute Sql
            Case "B"
                Sql = "SpTaAgrupacionRubrosContablesBorrar @A_Codigo =" & .A_Codigo
                Conec.Execute Sql
            End Select
        End With
    Next
    Conec.CommitTrans
    Modificado = False
ErrorGrabar:
    If Err.Number <> 0 Then
        Conec.RollbackTrans
        
        Call ManipularError(Err.Number, Err.Description)
    Else
        Call CargarVecRubrosContables
        Call CargarTv
        MsgBox "La Agrupación de Rubros se grabaron correctamente", vbInformation, "Grabado"
    End If
End Sub

Private Sub CmdEliminar_Click()
    If TreeRubros.SelectedItem.Children > 1 Then
        If TreeRubros.SelectedItem.Child.Text <> "Nuevo" Then
           TreeRubros.Nodes(TreeRubros.SelectedItem.Index).Checked = True
           MsgBox "El Rubro no se puede eliminar por que tiene Hijos", vbInformation
           Exit Sub
        Else
          TreeRubros.SelectedItem.Checked = False
        End If
    Else
      TreeRubros.SelectedItem.Checked = False
    End If
    Modificado = True
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
        TreeRubros.Nodes.Clear
        ReDim VecAgrupacionRubros(.RecordCount)
        For i = 1 To UBound(VecAgrupacionRubros)
            VecAgrupacionRubros(i).A_Codigo = !A_Codigo
            VecAgrupacionRubros(i).A_Descripcion = !A_Descripcion
            VecAgrupacionRubros(i).A_Nivel = !A_Nivel
            VecAgrupacionRubros(i).A_Padre = !A_Padre
            VecAgrupacionRubros(i).A_Rubro = !A_Rubro
            If !A_Padre <> 0 Then
                TreeRubros.Nodes.Add !A_Padre & "R", tvwChild, !A_Codigo & "R", !A_Descripcion
            Else
                TreeRubros.Nodes.Add , , !A_Codigo & "R", !A_Descripcion
            End If
            TreeRubros.Nodes(i).Checked = True
            .MoveNext
        Next
        For i = 1 To TreeRubros.Nodes.Count
            If VecAgrupacionRubros(TreeRubros.Nodes(i).Index).A_Rubro = "  " Then
                ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
                TreeRubros.Nodes.Add i, tvwChild, , "Nuevo"
            End If
        Next
        TreeRubros.Nodes.Add , , , "Nuevo"
        ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)

        TreeRubros.Nodes(TreeRubros.Nodes.Count).Selected = True
        Call TreeRubros_NodeClick(TreeRubros.SelectedItem)
    End With
    CmdConfirnar.Enabled = True
    Modificado = False
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdExportar_Click()
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
Dim Fila As Integer

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        ColorFondo = &HC0E0FF
        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = Caption
        .Range("A1:D1").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        
        .Selection.Merge  'COMBINAR CELDAS
        With .Selection.Font
            .Name = "Arial"
            .Size = 20
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        .Selection.Font.Bold = True
        .Range("C2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        
          Fila = 3
         .Range("A4" & ":F4").Select
         .Selection.Font.Bold = True
         .Range("B4").Value = "Rubro"
         .Range("A4").Value = "Codigo"
         Fila = 5
         Call AgregarNodoExcel(TreeRubros.Nodes(1), Fila, ex)
         '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
         .Columns("A:B").EntireColumn.AutoFit

         Call FormatearExcelGrid(ex, 4, CLng(Fila - 5), 2, ColorFondo)
    End With
    
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub AgregarNodoExcel(Nodo As Node, Fila As Integer, ex As Excel.Application)
Dim j As Integer
    
    If Not Nodo Is Nothing Then
        j = Nodo.Index
        If VecAgrupacionRubros(j).A_Descripcion <> "" Then
            If Nodo.Child Is Nothing Then
                ex.Range("A" & Trim(Fila)).FormulaR1C1 = VecAgrupacionRubros(j).A_Codigo
                ex.Range("B" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
                Fila = Fila + 1
            Else
                ex.Range("A" & Trim(Fila)).FormulaR1C1 = VecAgrupacionRubros(j).A_Codigo
                ex.Range("B" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
                Fila = Fila + 1
                Call AgregarNodoExcel(Nodo.Child, Fila, ex)
    
                'ex.Range("A" & Trim(Fila)).FormulaR1C1 = VecAgrupacionRubros(j).A_Codigo
                'ex.Range("B" & Trim(Fila)).Value = Space((VecAgrupacionRubros(j).A_Nivel - 1) * 4) & VecAgrupacionRubros(j).A_Descripcion
                'Fila = Fila + 1
            End If
        End If
        Call AgregarNodoExcel(Nodo.Next, Fila, ex)
    End If
End Sub

Private Sub CmdModificar_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = TreeRubros.SelectedItem.Index
      
      Modificado = True
     'lo pone en el LV
      VecAgrupacionRubros(i).A_Descripcion = TxtDescripcion.Text
      VecAgrupacionRubros(i).A_Rubro = VecRubrosContables(CmbRubros.ListIndex).Codigo
      VecAgrupacionRubros(i).Estado = IIf(VecAgrupacionRubros(i).A_Codigo = 0, "A", "M")
      VecAgrupacionRubros(i).A_Nivel = NivelNodo(TreeRubros.Nodes(i))
      If Not TreeRubros.SelectedItem.Parent Is Nothing Then
        VecAgrupacionRubros(i).A_Padre = VecAgrupacionRubros(TreeRubros.SelectedItem.Parent.Index).A_Codigo
      End If
      
      TreeRubros.Nodes(i).Text = TxtDescripcion.Text
      TreeRubros.Nodes(i).Checked = True
      
      'If TreeRubros.SelectedItem.Parent Is Nothing Then
      '   TreeRubros.Nodes.Add , , , "Nuevo"
      'Else
      '  ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
      '  TreeRubros.Nodes.Add TreeRubros.SelectedItem.Parent, 4, , "Nuevo"
      'End If
      
      If CmbRubros.ListIndex = 0 Then
        ReDim Preserve VecAgrupacionRubros(UBound(VecAgrupacionRubros) + 1)
        TreeRubros.Nodes.Add TreeRubros.SelectedItem, 4, , "Nuevo"
      End If
      
      TreeRubros.Nodes(i + 1).Selected = True
      Call TreeRubros_NodeClick(TreeRubros.SelectedItem)
      TreeRubros.Nodes(i).Expanded = True
      Modificado = True
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Rta As Integer
    If Modificado Then
        Rta = MsgBox("¿Desea guardar los cambios?", vbYesNoCancel)
        If Rta = vbYes Then
           Call Grabar
        Else
            If Rta = vbCancel Then
                Cancel = 1
            End If
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Call CargarCmbRubrosContables(CmbRubros)
    Call CargarTv
End Sub

Private Sub TreeRubros_KeyUp(KeyCode As Integer, Shift As Integer)
    Call TreeRubros_MouseUp(0, 0, 0, 0)
End Sub

Private Sub TreeRubros_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not xNodo Is Nothing Then
        xNodo.Checked = True
        Modificado = True
        Set xNodo = Nothing
        Exit Sub
    End If
End Sub

Private Sub TreeRubros_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Node.Checked = False Then
        If Node.Children > 1 Then
            If Node.Child.Text <> "Nuevo" Then
               Set xNodo = Node
               Node.Checked = True
               Modificado = True
               MsgBox "El Rubro no se puede eliminar por que tiene Hijos", vbInformation
               Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TreeRubros_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo Errores
'NO SE TOCA
   If Node.Text <> "Nuevo" Then
       TxtDescripcion.Text = VecAgrupacionRubros(Node.Index).A_Descripcion
       Call UbicarCmbRubrosContables(CmbRubros, VecAgrupacionRubros(Node.Index).A_Rubro)
       CmdEliminar.Enabled = True
       CmdAgregar.Enabled = False
       CmdModificar.Enabled = True
       If Node.Children > 0 Then
          CmbRubros.Enabled = False
       Else
          CmbRubros.Enabled = True
       End If
    Else
        TxtDescripcion = ""
        CmbRubros.ListIndex = 0
        CmbRubros.Enabled = True
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
        CmdModificar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub
