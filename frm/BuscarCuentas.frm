VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "El Pulqui"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerra&r"
      Height          =   510
      Left            =   4290
      TabIndex        =   7
      Top             =   7140
      Width           =   1635
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   510
      Left            =   2490
      TabIndex        =   6
      Top             =   7155
      Width           =   1635
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   7200
   End
   Begin VB.ComboBox CbBuscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6735
      Width           =   2220
   End
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3210
      TabIndex        =   3
      Top             =   6720
      Width           =   2550
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      Caption         =   "B&uscar"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   6735
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6165
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   10874
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Reducida 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1305
      TabIndex        =   10
      Top             =   7110
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Descripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   7455
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   8
      Top             =   7125
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LBTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuentas Contables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8325
   End
   Begin VB.Label LbConexion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lbconexion"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7290
      TabIndex        =   5
      Top             =   7110
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "BuscarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LlevaArt As Integer

Private Type VariablesGlobales
    Servidor As String
    Env As rdoEnvironment
    db As rdoConnection
    TbListado As rdoResultset
    VecBuscar() As Integer
    Cargar As Boolean
End Type

Const VgNumero = "#0.00" 'esta constante es el formato de los numeros
Dim v As VariablesGlobales

Public Sub CargarParametros(Conexion As String)
    LbConexion.Caption = Conexion
    Call CargarListado
End Sub

Private Sub CmdCerrar_Click()
    Hide
End Sub

Private Sub ConfImpresion()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    RsListado.Fields.Append "Codigo", adVarChar, 10
    RsListado.Fields.Append "Descripcion", adVarChar, 100
    RsListado.Fields.Append "Reducido", adVarChar, 50
    RsListado.Open
    i = 1
    For i = 1 To LVListado.ListItems.Count
        RsListado.AddNew
        With LVListado.ListItems(i)
            RsListado!Codigo = .Text
            RsListado!Descripcion = .SubItems(1)
            RsListado!Reducido = .SubItems(2)
        End With
    Next
    RsListado.MoveFirst
         
    RepCuentas.TxtFecha = Date
    RepCuentas.DataControl1.Recordset = RsListado
    RepCuentas.Zoom = -1
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresion
    RepCuentas.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errores
    TeclaPresionada ActiveControl, KeyAscii
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LbConexion_Change()
On Error GoTo errores
    Dim NombreConexion As String
    Dim Usuario As String
    Dim Clave As String
    Dim Opciones As String
    MousePointer = vbHourglass
    Set v.Env = rdoEnvironments(0)
    NombreConexion = BuscarString(LbConexion.Caption, "dsn=")
    Usuario = BuscarString(LbConexion.Caption, "UID=")
    Clave = BuscarString(LbConexion.Caption, "PWD=")
    If Trim(Clave) = "" Then
        Opciones = "UID=" & Usuario
    Else
        Opciones = "UID=" & Usuario & ";PWD=" & Clave
    End If
    
    v.Servidor = BuscarString(LbConexion.Caption, "servidor=")
    Set v.db = v.Env.OpenConnection(NombreConexion, rdDriverNoPrompt, False, Opciones)
    Call InicializarTodo
    MousePointer = vbNormal
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub CargarListado()
On Error GoTo errores
Dim sSQL As String
        MousePointer = vbHourglass
        sSQL = "SpTACuentas"
        Set v.TbListado = v.db.OpenResultset(sSQL)
        LVListado.ListItems.Clear
        LVListado.Sorted = False
        ' esto activa el timer para empezar a cargar los renglones
        
        Timer1.Enabled = True
        MousePointer = vbNormal

errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub InicializarTodo()
On Error GoTo errores
   ' InicializarTags
    Call CargarEncabezados
    Call CargarComboBuscar
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub CargarEncabezados()
On Error GoTo errores
Dim TamanioTotal As Integer
Dim TamanioColumna As Integer
    'tamaniocolumna es el tamaño que va a tener la columna indicada
    
    TamanioColumna = 1000
    TamanioTotal = TamanioColumna
    LVListado.ColumnHeaders.Add 1, "Código", "Código", TamanioColumna
    
    TamanioColumna = 4300
    TamanioTotal = TamanioTotal + TamanioColumna
    LVListado.ColumnHeaders.Add 2, "Descripción", "Descripción", TamanioColumna
    
    LVListado.ColumnHeaders.Add 3, "Reducida", "Reducida", LVListado.Width - 5550
    'el tamaño de la última columna se actualiza automáticamente para acomodarse al listview
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub


Private Sub CMDBuscar_Click()
    
On Error GoTo errores
    Buscar
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LVListado_DblClick()
    If LVListado.ListItems.Count > 0 Then
        Codigo.Caption = Trim(LVListado.ListItems(LVListado.SelectedItem.Index).Text)
        Descripcion.Caption = Trim(LVListado.ListItems(LVListado.SelectedItem.Index).SubItems(1))
        Reducida.Caption = Trim(LVListado.ListItems(LVListado.SelectedItem.Index).SubItems(2))
    End If
    Me.Visible = False

End Sub

Private Sub Timer1_Timer()
'On Error GoTo errores
Dim i As Integer
    i = 1
    With v.TbListado
    'cargo de a 25 renglones para no perder tanto tiempo y luego vuelve a arrancar el timer
    
        While Not .EOF And i < 25
            LVListado.ListItems.Add
            LVListado.ListItems(LVListado.ListItems.Count).Text = VerificarNulo(!C_Codigo)
            LVListado.ListItems(LVListado.ListItems.Count).SubItems(1) = Trim(VerificarNulo(!C_Descripcion))
            LVListado.ListItems(LVListado.ListItems.Count).SubItems(2) = Trim(VerificarNulo(!C_Reducida))
            'el vector vectotales es un acumulador para calcular los totales
            i = i + 1
            .MoveNext
        Wend
    If .EOF Then
        'cuando terminé de calcular todos los renglones deshabilito el timer
        .Close
        Timer1.Enabled = False
    End If
    End With
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub



Private Sub CargarComboBuscar()
On Error GoTo errores
    'el vector vecbuscar indica el campo por el que se tiene que buscar en el listado
    ReDim v.VecBuscar(1)
    CbBuscar.AddItem "Código"
    v.VecBuscar(0) = 0
    CbBuscar.AddItem "Descripción"
    v.VecBuscar(1) = 1
    CbBuscar.ListIndex = 0
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub


Private Sub Buscar()
On Error GoTo errores
'este procedimiento busca en el listview el valor ingresado en el textbox txtbuscar
' y basándose en el vector vecbuscar y el combo cbuscar
Dim Encontro As Boolean
    Encontro = RealizarBusqueda(LVListado.SelectedItem.Index + 1, LVListado.ListItems.Count)
    If Not Encontro Then
        Encontro = RealizarBusqueda(1, LVListado.SelectedItem.Index)
        If Not Encontro Then
            MsgBox "no existe ningún campo que coincida con su criterio de búsqueda"
        End If
    End If
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Function RealizarBusqueda(desde As Integer, hasta As Integer) As Boolean
On Error GoTo errores
Dim Encontro As Boolean
Dim i As Integer
    Encontro = False
    i = desde
    While Not Encontro And (i <= hasta)
        If v.VecBuscar(CbBuscar.ListIndex) = 0 Then ' si es por el primer campo del listview
            If InStr(1, UCase(LVListado.ListItems(i).Text), UCase(TxtBuscar.Text)) > 0 Then
                Encontro = True
            End If
        Else
            If InStr(1, UCase(LVListado.ListItems(i).SubItems(v.VecBuscar(CbBuscar.ListIndex))), UCase(TxtBuscar.Text)) > 0 Then
                Encontro = True
            End If
        End If
        i = i + 1
    Wend
    If Encontro Then
        LVListado.SetFocus
        LVListado.ListItems(i - 1).Selected = True
        LVListado.ListItems(i - 1).EnsureVisible
    End If
    RealizarBusqueda = Encontro
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Function

Private Sub LvListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo errores
   ' Cuando se hace clic en un objeto ColumnHeader, el
   ' control ListView se ordena por los subelementos de
   ' esa columna.
   ' Establece el SortKey como el Index del ColumnHeader - 1
   ' Asigna a Sorted el valor True para ordenar la lista.
   LVListado.SortKey = ColumnHeader.Index - 1
   If LVListado.SortOrder = lvwAscending Then
        LVListado.SortOrder = lvwDescending
   Else
       LVListado.SortOrder = lvwAscending
   End If
   LVListado.Sorted = True
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

