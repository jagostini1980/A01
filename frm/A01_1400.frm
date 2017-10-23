VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_1400 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros de costo"
   ClientHeight    =   8325
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   10380
   Icon            =   "A01_1400.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtEmailAutorizar 
      Height          =   285
      Left            =   1755
      TabIndex        =   7
      Top             =   1350
      Width           =   4590
   End
   Begin VB.TextBox TxtDescripcion 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1755
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   18
      Top             =   75
      Width           =   3315
   End
   Begin VB.TextBox TxtMonto 
      Height          =   285
      Left            =   1770
      TabIndex        =   3
      Top             =   765
      Width           =   1005
   End
   Begin VB.CommandButton CmdAuxiliares 
      Caption         =   "Au&xiliares"
      Height          =   300
      Left            =   5190
      TabIndex        =   2
      Top             =   435
      Width           =   960
   End
   Begin VB.TextBox TxtMaxSinAutorizar 
      Height          =   285
      Left            =   4740
      TabIndex        =   4
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imp. Centros de Costos"
      Height          =   375
      Left            =   45
      TabIndex        =   17
      Top             =   7875
      Width           =   1785
   End
   Begin MSComctlLib.TreeView TvCuentas 
      Height          =   2985
      Left            =   45
      TabIndex        =   8
      Top             =   4860
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5265
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton CmdImpArt 
      Caption         =   "Imp. &Articulos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3150
      TabIndex        =   15
      Top             =   7875
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   6930
      Top             =   7830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExportarCta 
      Caption         =   "&Exporta Cuentas a Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4410
      TabIndex        =   14
      Top             =   7875
      Width           =   1920
   End
   Begin VB.CommandButton CmdImpCuentas 
      Caption         =   "I&mp. Cuentas"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1890
      TabIndex        =   13
      Top             =   7875
      Width           =   1200
   End
   Begin VB.Frame FrameTipoArt 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   12
      Top             =   1095
      Width           =   2085
      Begin VB.OptionButton OptArt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "De Taller"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   1005
      End
      Begin VB.OptionButton OptArt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Otros"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   6
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   7785
      TabIndex        =   10
      Top             =   7875
      Width           =   1200
   End
   Begin VB.CommandButton CMDSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9090
      TabIndex        =   11
      Top             =   7875
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_1400.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_1400.frx":4F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_1400.frx":7706
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_1400.frx":7A20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvArticulos 
      Height          =   7755
      Left            =   6480
      TabIndex        =   9
      Top             =   60
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   13679
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin Controles.ComboEsp CmbUsuarios 
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   428
      Width           =   3315
      _ExtentX        =   5847
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
   Begin MSComctlLib.TreeView TvCentros 
      Height          =   2955
      Left            =   45
      TabIndex        =   0
      Top             =   1665
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5212
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "E-Mail Autorizacion:"
      Height          =   195
      Left            =   255
      TabIndex        =   24
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label LbDescripcion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Centro de Costo:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   510
      TabIndex        =   23
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label LBUsuario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ususario Responsable:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   488
      Width           =   1635
   End
   Begin VB.Label LbTipoArt 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo de Artículo:"
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Monto Max. Sin Presupuestar $:"
      Height          =   375
      Left            =   465
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Monto Max. Sin Autorizar:"
      Height          =   195
      Left            =   2850
      TabIndex        =   19
      Top             =   810
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuentas Contables:"
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
      Left            =   90
      TabIndex        =   16
      Top             =   4635
      Width           =   1665
   End
End
Attribute VB_Name = "A01_1400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private C_Codigo As String
Private xNodo As Node

Private Sub CmbUsuarios_Click()
        Modificado = True
End Sub

Private Function AgregarCentro() As Boolean
 On Error GoTo Error
    Dim Sql As String
    AgregarCentro = True
     
     If CmbUsuarios.ListIndex > 0 Then
        Dim RsGuardar As ADODB.Recordset
        Set RsGuardar = New ADODB.Recordset
        
        Sql = "SpTA_CentrosDeCostosAgregar @C_Descripcion='" + TxtDescripcion.Text + _
              "', @C_Padre = Null , @@C_UsuarioResponsable ='" & CmbUsuarios.Text & "'" & _
              ", @@C_TablaArticulos = " & IIf(OptArt(1).Value, "'TA_Articulos'", "NULL") & _
              ", @@C_MontoSinPresupuestar = " & Val(TxtMonto)

          RsGuardar.Open Sql, Conec
          
          If RsGuardar!Ok = "OK" Then
                'MsgBox RsGuardar!Mensaje, , Me.Caption
                C_Codigo = RsGuardar!C_Codigo
          Else
                MsgBox RsGuardar!Mensaje, vbInformation, Me.Caption
                AgregarCentro = False
                TxtDescripcion.SetFocus
                TxtDescripcion.SelStart = 0
                TxtDescripcion.SelLength = Len(TxtDescripcion.Text)
                Exit Function
          End If
       Else
           MsgBox "Debe seleccionar un Usuario Responsable"
           AgregarCentro = False
           CmbUsuarios.SetFocus
       End If
Error:
    If Err.Number <> 0 Then
        AgregarCentro = False
        Call ManipularError(Err.Number, Err.Description)
    End If

End Function

Private Sub CmdAuxiliares_Click()
    A01_1430.CentroDeCastoActual = C_Codigo
    A01_1430.Show vbModal
End Sub

Private Sub CmdConfirmar_Click()
 Dim Sql As String
 Dim Pregunta As Integer
 Dim i As Integer
  Pregunta = MsgBox("¿Desea Modificar el Centro de Costo?", vbQuestion + vbOKCancel, "Pulqui")
  
  If Pregunta = vbOK Then
        Call GuardarCambios
  End If
End Sub

Private Sub GuardarCambios()
  Dim Sql As String
  Dim i As Integer
  
 On Error GoTo Error
     
    'If TvCentros.SelectedItem.Key = "Nuevo" Then
    '    If Not AgregarCentro Then
    '        Exit Sub
    '    End If
    'Else
        If Not ModificarCentro Then
            If Not Modificado Then
                Exit Sub
            End If
        End If
    'End If
 
     Conec.BeginTrans
        'Sql = "SpOcParametrosActualizar @P_EMailAutorizacion='" & TxtEmailAutorizar & "'"
        'Conec.Execute Sql
        
        'ACTUALIZA LAS RELACIONES DE LOS ARTÍCULOS
        Sql = "SpOCRelacionCentroDeCostoArticulosBorrar @R_CentroDeCosto = '" & C_Codigo & "'"
        Conec.Execute Sql
        For i = 1 To LvArticulos.ListItems.Count
            If LvArticulos.ListItems(i).Checked Then
                Sql = "SpOCRelacionCentroDeCostoArticulosAgregar " & _
                        "@R_CentroDeCosto ='" & C_Codigo & _
                      "', @R_Articulo = " & Val(LvArticulos.ListItems(i).Key)
                Conec.Execute Sql
            End If
        Next
        
        'ACTUELIZA LAS RELACIONES CON LAS CUENTAS CONTABLES
        Sql = "SpOCRelacionCentroDeCostoCuentaContableBorrar @R_CentroDeCosto = '" & C_Codigo & "'"
        Conec.Execute Sql
        For i = 1 To TvCuentas.Nodes.Count
            If TvCuentas.Nodes(i).Checked Or TvCuentas.Nodes(i).Tag = "X" Then
                Sql = "SpOCRelacionCentroDeCostoCuentaContableAgregar " & _
                         "@R_CentroDeCosto ='" & C_Codigo & _
                      "', @R_CuentaContable = '" & Mid(TvCuentas.Nodes(i).Key, 1, 4) & "'"
                Conec.Execute Sql
            End If
        Next
        Conec.CommitTrans

        Call CargarVecCentrosDeCostosEmisor
        Modificado = False
        Call CargarTvCentros
        Call CargarEmailAutorizacion
Error:
    If Err.Number = 0 Then
        MsgBox "La modificación se realizó correctamente", , "Modificación"
        Modificado = False
    Else
        Conec.RollbackTrans
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Function ModificarCentro() As Boolean
On Error GoTo Error
    Dim Sql As String
    Dim RsGuardar As New ADODB.Recordset

    ModificarCentro = True
    If BuscarCentroPadre(Val(C_Codigo)) = "" Then
        ModificarCentro = False
        Exit Function
    End If
    
    If CmbUsuarios.ListIndex > 0 Then
        Sql = "SpOCImporteSinPresupuestarAutorizacionDeCargaContable @CentroDeCosto ='" & C_Codigo & _
                 "', @NroAutorizacion =0" & _
                 " , @Periodo=" & FechaSQL(Date, "SQL")
        RsGuardar.Open Sql, Conec
        If RsGuardar!MontoSinPres > Val(TxtMonto.Text) Then
            MsgBox "El Monto Max. sin presupuestar no puede ser Modificado", vbInformation, "Uso Actual $" & RsGuardar!MontoSinPres
            TxtMonto.Text = RsGuardar!MontoSinPresupuestarMensual
        End If
        
        RsGuardar.Close
        Sql = "SpTA_CentrosDeCostosModificar @C_Codigo='" & C_Codigo & _
              "', @C_Descripcion ='" & TxtDescripcion.Text & _
              "', @@C_UsuarioResponsable = '" & CmbUsuarios.Text & _
              "', @@C_MontoSinPresupuestar = " & Val(TxtMonto.Text) & _
              " , @@C_ImporteAutorizar = " & Val(TxtMaxSinAutorizar) & _
              " , @@C_EmailAutorizacion ='" & TxtEmailAutorizar & "'"
               
        RsGuardar.Open Sql, Conec
          
        If RsGuardar!Ok = "OK" Then
            'TvCentros.SelectedItem.Text = TxtDescripcion.Text
            'Call CargarVecCentrosDeCostosEmisor
        Else
            MsgBox RsGuardar!Mensaje, vbInformation
            ModificarCentro = False
            'TxtDescripcion.SetFocus
            TxtDescripcion.SelStart = 0
            TxtDescripcion.SelLength = Len(TxtDescripcion.Text)
        End If
    Else
        MsgBox "Debe Seleccionar un Usuario Responsable"
        CmbUsuarios.SetFocus
        ModificarCentro = False
    End If
    
Error:
    If Err.Number <> 0 Then
        ModificarCentro = False
        Call ManipularError(Err.Number, Err.Description)
    End If
End Function

Private Sub CmdImpArt_Click()
    Call ConfImpresionDeArticulos
    ListA01_1400Articulos.Show
End Sub

Private Sub ConfImpresionDeArticulos()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
  Dim Nodo As Node
    
    RsListado.Fields.Append "Articulos", adVarChar, 60
    
    RsListado.Open
    i = 1
    While i <= LvArticulos.ListItems.Count
        If LvArticulos.ListItems(i).Checked Then
           RsListado.AddNew
           RsListado!Articulos = LvArticulos.ListItems(i).Text
        End If
        i = i + 1
    Wend
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    RsListado.Sort = "Articulos"
    
    ListA01_1400Articulos.TxtCentro = TxtDescripcion
    ListA01_1400Articulos.DataControl1.Recordset = RsListado
    ListA01_1400Articulos.Zoom = -1
End Sub

Private Sub CmdImpCuentas_Click()
    Call ConfImpresionDeCuentas
    ListA01_1400Cuentas.Show
End Sub

Private Sub ConfImpresionDeCuentas()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
  Dim Nodo As Node
    
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    
    RsListado.Open
    i = 1
    While i <= TvCuentas.Nodes.Count
        If TvCuentas.Nodes(i).Checked Then
           RsListado.AddNew
           RsListado!Cuenta = TvCuentas.Nodes(i).Text
        End If
        i = i + 1
    Wend
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    RsListado.Sort = "Cuenta"
    
    ListA01_1400Cuentas.TxtCentro = TxtDescripcion
    ListA01_1400Cuentas.DataControl1.Recordset = RsListado
    ListA01_1400Cuentas.Zoom = -1
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeCentros
    ListA01_1400.Show
End Sub

Private Sub ConfImpresionDeCentros()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
  Dim Nodo As Node
    
    RsListado.Fields.Append "Centro", adVarChar, 50
    RsListado.Fields.Append "SubCentro", adVarChar, 50
    RsListado.Fields.Append "NivelIntermedio", adVarChar, 50
    RsListado.Fields.Append "Responsable", adVarChar, 25
    
    RsListado.Open
    i = 1
    While i <= TvCentros.Nodes.Count
        
        If Not TvCentros.Nodes(i).Parent Is Nothing Then
            If TvCentros.Nodes(i).Child Is Nothing Then
                 RsListado.AddNew
                 If Not TvCentros.Nodes(i).Parent.Parent Is Nothing Then
                    RsListado!Centro = TvCentros.Nodes(i).Parent.Parent.Text
                    RsListado!Responsable = VecCentroDeCostoEmisor(TvCentros.Nodes(i).Parent.Parent.Index).C_UsuarioResponsable
                 Else
                    RsListado!Centro = TvCentros.Nodes(i).Parent.Text
                 RsListado!Responsable = VecCentroDeCostoEmisor(TvCentros.Nodes(i).Parent.Index).C_UsuarioResponsable
                 End If
                 If Not TvCentros.Nodes(i).Parent.Parent Is Nothing Then
                     RsListado!NivelIntermedio = TvCentros.Nodes(i).Parent.Text
                 Else
                     RsListado!NivelIntermedio = "(Sin Nivel Intermedio)"
                 End If
                 
                 RsListado!SubCentro = TvCentros.Nodes(i).Text
    
            End If
        End If
        i = i + 1
    Wend
    RsListado.MoveFirst
    RsListado.Sort = "Centro, NivelIntermedio, SubCentro"
    'RsListado.Sort = "NivelIntermedio"
    'RsListado.Sort = "SubCentro"


    ListA01_1400.DataControl1.Recordset = RsListado
    ListA01_1400.Zoom = -1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdExportarCta_Click()
    Dialogo.Filename = ""
    Call ArmarExcelCta(Dialogo)
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

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcelCta(ex, "Cuentas del Centro de Costo: " & TvCentros.SelectedItem.Text, 4)
        Call DatosExcelCta(ex, TvCuentas, 4)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        'For col = 1 To LvCuentas.ColumnHeaders.Count
            .Columns(LetraColumna(1) & ":" & LetraColumna(1)).EntireColumn.AutoFit
        'Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("B2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        
        ColorFondo = &HC0E0FF

        Call FormatearExcelCta(ex, TvCuentas, 4, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub Form_Load()
   TxtEmailAutorizar = EmailAutorizacion
   Call CrearEmcabezados
   Call CargarLV
   'Call CargarCmbCentrosDeCostosNivel2(CmbNivelIntermedio)
   Call CargarCmbUsuarios(CmbUsuarios)
   Modificado = False
   Call CargarTvCentros
   Modificado = False
End Sub

Private Sub CrearEmcabezados()
        LvArticulos.ColumnHeaders.Add , , "Artículo", LvArticulos.Width - 250
        'LvCuentas.ColumnHeaders.Add , , "Cuenta Contable", LvCuentas.Width - 250
End Sub

Private Sub CargarLV()
    Dim i As Integer
    Dim KeyPadre As String
    For i = 1 To UBound(VecCuentasContablesArbol)
        With VecCuentasContablesArbol(i)
            If .P_NIVEL = 1 Then
               TvCuentas.Nodes.Add , , .Codigo & "C", .Descripcion
            Else
               KeyPadre = BuscarPadre(.P_PADRE) & "C"
               TvCuentas.Nodes.Add KeyPadre, tvwChild, .Codigo & "C", .Descripcion & " - Cód. " & Trim(.Codigo)
            End If
        End With
    Next
    
    For i = 1 To UBound(VecArtCompra)
        LvArticulos.ListItems.Add , VecArtCompra(i).A_Codigo & "A", VecArtCompra(i).A_Descripcion
    Next
End Sub

Private Function BuscarPadre(Padre As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCuentasContablesArbol)
        If VecCuentasContablesArbol(i).P_JER = Padre Then
           BuscarPadre = VecCuentasContablesArbol(i).Codigo
           Exit Function
        End If
    Next
End Function

Private Sub CargarTvCentros()
 Dim Nodo As Node
  Dim i As Integer
 TvCentros.Nodes.Clear
     
     'carga el primer nivel
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        With VecCentroDeCostoEmisor(i)
           Set Nodo = TvCentros.Nodes.Add(, , .C_Jerarquia & "C", .C_Descripcion, 1)
        End With
    Next
    
    For i = 1 To UBound(VecCentroDeCostoNivel2)
        'carga el segundo nivel
        With VecCentroDeCostoNivel2(i)
             Set Nodo = TvCentros.Nodes.Add(CStr(.C_Padre) & "C", tvwChild, CStr(.C_Jerarquia) + "C", .C_Descripcion, 4)
        End With
    Next
  
    For i = 1 To UBound(VecCentroDeCosto)
        'carga hojas
        With VecCentroDeCosto(i)
             Set Nodo = TvCentros.Nodes.Add(CStr(.C_Padre) + "C", tvwChild, CStr(.C_Codigo) + "C", .C_Descripcion, 2)
        End With
    Next
    'TvCentros.Nodes.Add , , "Nuevo", "Nuevo Centros de Costo", 3
    
    TvCentros.Nodes(1).Selected = True
    
    Call TvCentros_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modificado Then
        Dim Rta As Integer
       Rta = MsgBox("Ha efectuado cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            Call GuardarCambios
         End If
       End If

    End If
End Sub

Private Sub LvArticulos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Modificado = True
  If Not Item.Checked Then
    Dim Sql As String
    Dim RsValidar As New ADODB.Recordset
    
    Sql = "SpVerificarArticulo @CentroDeCosto=" & Val(TvCentros.SelectedItem.Key) & _
                                  ", @Articulo= " & Val(Item.Key)
    RsValidar.Open Sql, Conec
    If RsValidar!Ok = "NO" Then
        Item.Checked = True
    End If
  End If
End Sub

Private Sub TvCentros_Click()
 Dim Nodo As Node
 Dim i As Integer
 Dim Rta As Integer
 Dim IndexEmisor As Integer
 
    If Modificado Then
       Rta = MsgBox("Ha efectuado cambio ¿Desea Guardarlos?", vbYesNo)
       If Rta = vbYes Then
            Call GuardarCambios
       End If
    End If
    CmdImpCuentas.Enabled = False
    CmdExportarCta.Enabled = False
    CmdImpArt.Enabled = False
    
   If TvCentros.SelectedItem.Key = "Nuevo" Then
        OptArt(1).Value = False
        OptArt(0).Value = False
        LbTipoArt.Enabled = True
        'OptArt(1).Enabled = True
        'OptArt(0).Enabled = True
        
        TxtDescripcion.Text = ""
        TxtDescripcion.Enabled = True
        CmbUsuarios.Enabled = True
        CmbUsuarios.ListIndex = 0
        'TxtSubordinado.Text = ""
        'LbSub.Enabled = False
        'TxtSubordinado.Enabled = False
        'CmdAgregar.Enabled = False
        'CmdModifSub.Visible = False
        'CMDBorrar.Enabled = False
        
        'limpia el LV de artículos
        For i = 1 To LvArticulos.ListItems.Count
            LvArticulos.ListItems(i).Checked = False
        Next
        'limpia el LV de Cuentas
        For i = 1 To TvCuentas.Nodes.Count
            TvCuentas.Nodes(i).Checked = False
        Next
        
        TvCuentas.Enabled = False
        LvArticulos.Enabled = False
        TxtDescripcion.SetFocus
        'CmbNivelIntermedio.Enabled = False
        'CmbNivelIntermedio.ListIndex = 0
   Else
        'CMDBorrar.Enabled = True

        Set Nodo = TvCentros.SelectedItem.Parent
        'TxtSubordinado.Text = ""
        'TxtSubordinado.Enabled = True
        'LbSub.Enabled = True

     If Nodo Is Nothing Then
        'si no tiene padre es un centro de costo
        Call UbicarUsuario(VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_UsuarioResponsable, CmbUsuarios)
        Call CheckArticulos(VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_Codigo)
        Call CheckCuentas(VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_Codigo)
       ' CmbNivelIntermedio.Enabled = False
        'CmbNivelIntermedio.ListIndex = 0
        
        TxtDescripcion.Enabled = True
        TvCuentas.Enabled = True
        TxtMonto.Text = VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_MontoSinPresupuestar
        TxtMaxSinAutorizar.Text = VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_ImporteAutorizar
        TxtEmailAutorizar.Text = VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_EmailAutorizacion
        TxtEmailAutorizar.Enabled = True
        TxtMonto.Enabled = True
        'si el campo C_TablaArticulos = "" se deja seleccionar los artículos
        LvArticulos.Enabled = VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_TablaArticulos = ""
        CmdAuxiliares.Enabled = True
        CmbUsuarios.Enabled = True
        'CmdModifSub.Visible = False
        'CmdModifSub.Enabled = False
        'CmdAgregar.Enabled = True
        'CmdAgregar.Visible = True
        'TxtSubordinado.Enabled = True
        
        LbTipoArt.Enabled = True
        'una letra T en el campo C_TablaArticulos significa que los art son de la tabla TA_Articulos
        If VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_TablaArticulos = "" Then
            OptArt(0).Value = True
        Else
            OptArt(1).Value = True
        End If

        TxtDescripcion.Text = TvCentros.SelectedItem.Text
        CmdImpCuentas.Enabled = True
        CmdExportarCta.Enabled = True
        CmdImpArt.Enabled = True
     Else
       'si tiene padre es un sub-centro de costo o Intermedio
       
        CmdAuxiliares.Enabled = False
        TxtDescripcion.Text = Nodo.Text
        TxtDescripcion.Enabled = False
        CmbUsuarios.Enabled = False
        TxtEmailAutorizar.Enabled = False
        TvCuentas.Enabled = False
        
        Dim NodoCuenta As Node
        For Each NodoCuenta In TvCuentas.Nodes
            NodoCuenta.Expanded = False
        Next
        LvArticulos.Enabled = False
        
        TxtMonto.Text = VecCentroDeCostoEmisor(IndexEmisor).C_MontoSinPresupuestar
        TxtMonto.Enabled = False
        
        OptArt(0).Enabled = False
        OptArt(1).Enabled = False
        'una letra T en el campo C_TablaArticulos significa que los art son de la tabla TA_Articulos
        'algo distinto de T significa que los art son de la tabla TA_ArticulosCompra
        If VecCentroDeCostoEmisor(IndexEmisor).C_TablaArticulos = "" Then
            OptArt(0).Value = True
        Else
            OptArt(1).Value = True
        End If
        
        Call UbicarUsuario(VecCentroDeCostoEmisor(IndexEmisor).C_UsuarioResponsable, CmbUsuarios)
        Call CheckArticulos(Val(VecCentroDeCostoEmisor(IndexEmisor).C_Codigo))
        Call CheckCuentas(Val(VecCentroDeCostoEmisor(IndexEmisor).C_Codigo))
     End If
   End If
     Modificado = False
     
     If TvCentros.SelectedItem.Parent Is Nothing Then
        C_Codigo = Mid(VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_Codigo, 1, 4)
     End If
End Sub

Private Sub CheckArticulos(CodCentro As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
    
    For i = 1 To LvArticulos.ListItems.Count
        LvArticulos.ListItems(i).Checked = False
    Next
    
    Sql = "SpOCRelacionCentroDeCostoArticulosTraer @R_CentroDeCosto='" & CodCentro & "'"
    With RsCargar
              
       .Open Sql, Conec, adOpenStatic, adLockReadOnly
        While Not .EOF
            LvArticulos.ListItems(!R_Articulo & "A").Checked = True
           .MoveNext
        Wend
       .Close
    End With
    Set RsCargar = Nothing
    
End Sub

Private Sub CheckCuentas(CodCentro As String)
On Error GoTo ErrorCta
 Dim KeyPadre As String
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
    MousePointer = vbHourglass

    For i = 1 To TvCuentas.Nodes.Count
        TvCuentas.Nodes(i).Checked = False
        TvCuentas.Nodes(i).BackColor = vbWhite
    Next
    
    Sql = "SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CodCentro & "'"
    With RsCargar
              
        .Open Sql, Conec, adOpenStatic, adLockReadOnly
        While Not .EOF
            TvCuentas.Nodes(!R_CuentaContable & "C").Checked = True
            'TvCuentas.Nodes(!R_CuentaContable & "C").Selected = True
            KeyPadre = !R_CuentaContable & "C"
            .MoveNext
            
            While Not TvCuentas.Nodes(KeyPadre).Parent Is Nothing
                TvCuentas.Nodes(KeyPadre).Parent.BackColor = &HFFC0C0
                KeyPadre = TvCuentas.Nodes(KeyPadre).Parent.Key
                'TvCuentas.SelectedItem.Parent.Selected = True
            Wend
        Wend
        .Close
    End With
    Set RsCargar = Nothing
ErrorCta:
  Call ManipularError(Err.Number, Err.Description)
  MousePointer = vbNormal
End Sub

Private Sub TvCentros_NodeClick(ByVal Node As MSComctlLib.Node)
    Call TvCentros_Click
End Sub

Private Sub TvCuentas_KeyUp(KeyCode As Integer, Shift As Integer)
    Call TvCuentas_MouseUp(0, 0, 0, 0)
End Sub

Private Sub TvCuentas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Dim Nodo As Node
    If Not xNodo Is Nothing Then
        If xNodo.Tag = "X" Then
            xNodo.Checked = True
        Else
            xNodo.Checked = False
        End If
        
        Set xNodo = Nothing
        Exit Sub
    End If
            
End Sub

Private Sub TvCuentas_NodeCheck(ByVal Node As MSComctlLib.Node)
  'Dim xNodo As Node
  Dim i As Integer
  
  If Node.Checked And Not Node.Child Is Nothing Then
    If Not Node.Child.Child Is Nothing Then
        MsgBox "La cuenta no es imputable"
        Set xNodo = Node
        Node.Checked = False
        Exit Sub
    End If
  End If
  If Not Node.Parent Is Nothing Then
    If Node.Parent.Checked Then
          MsgBox "La cuenta tiene Cuentas de Mayor Nivel Relacionadas"
          Set xNodo = Node
          Node.Checked = False
          Exit Sub
    End If
  End If
  
  If Not Node.Child Is Nothing Then
    For i = Node.Child.Index To Node.Child.Index + Node.Children
        If TvCuentas.Nodes(i).Checked Then
          MsgBox "La cuenta tiene Cuentas de menor Nivel Relacionadas"
          Set xNodo = Node
          Node.Checked = False
          Exit Sub
        End If
    Next
  End If
  
  Modificado = True
  
  If Not Node.Checked Then
      Dim Sql As String
      Dim RsValidar As New ADODB.Recordset
        
      Sql = "SpVerificarCuentaContable @CentroDeCosto='" & VecCentroDeCostoEmisor(TvCentros.SelectedItem.Index).C_Codigo & _
                                   "', @CuentaContable= '" & Mid(Node.Key, 1, 4) & "'"
      RsValidar.Open Sql, Conec
        If RsValidar!Ok = "NO" Then
            MsgBox "La Cuenta no puede ser desvinculada del Centro de Costo", vbInformation
            Set xNodo = Node
            Node.Checked = True
            Node.Tag = "X"
        End If
  End If

End Sub

Private Sub TxtDescripcion_Change()
    Call ColorObligatorio(TxtDescripcion, CmdConfirmar)
    If TxtDescripcion.Text = "" Then
        Modificado = False
    Else
        Modificado = True
    End If
End Sub

Private Sub TxtDescripcion_GotFocus()
On Error GoTo Errores
    SelText TxtDescripcion
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub TxtMaxSinAutorizar_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtMaxSinAutorizar, KeyAscii)
    Modificado = True
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
    End If
    Modificado = True
End Sub
