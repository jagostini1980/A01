VERSION 5.00
Begin VB.Form A01_1100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas"
   ClientHeight    =   2805
   ClientLeft      =   3600
   ClientTop       =   3210
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4680
      Top             =   405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuentas "
      Height          =   2475
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   7620
      Begin VB.TextBox TxtReducida 
         Height          =   315
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1350
         Width           =   3060
      End
      Begin VB.CommandButton CMDSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5895
         TabIndex        =   8
         Top             =   1875
         Width           =   1455
      End
      Begin VB.CommandButton CMDGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Guardar"
         Height          =   375
         Left            =   225
         TabIndex        =   5
         Top             =   1875
         Width           =   1455
      End
      Begin VB.CommandButton CMDBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4005
         TabIndex        =   7
         Top             =   1875
         Width           =   1455
      End
      Begin VB.CommandButton CMDBorrar 
         Appearance      =   0  'Flat
         Caption         =   "Borrar"
         Height          =   375
         Left            =   2115
         TabIndex        =   6
         Top             =   1875
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   855
         Width           =   6135
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   315
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LbReducida 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reducida:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1395
         Width           =   735
      End
      Begin VB.Label LbDescripcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   945
         Width           =   885
      End
      Begin VB.Label LbCodigo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   555
         TabIndex        =   1
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.Label LbConexion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "A01_1100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type VariablesGlobales
    Servidor As String
    Env As rdoEnvironment
    db As rdoConnection
    Modificado As Boolean
End Type

Private Type VariablesImpresion
    TamanioLetra As Single
    SeparacionConceptos As Integer
    CantidadLetras As Integer
    ancho As Integer
    VecPosiciones() As Integer        'es la posicion donde se imprime cada encabezado
    VecPosicionesCheque() As Integer        'es la posicion donde se imprime cada encabezado
End Type


Const VgNumero = "#0.00" 'esta constante es el formato de los numeros
Dim v As VariablesGlobales
Dim Vi As VariablesImpresion

Private Sub CMDBorrar_Click()
Dim sSQL As String
Dim Tabla As rdoResultset
Dim Pregunta As Integer
    If v.Modificado Then
       Pregunta = MsgBox("¿Esta seguro que desea borrar el registro actual?", vbQuestion + vbOKCancel, "PulquiPack")
       If Pregunta = vbOK Then
            If ValidarBorrar Then
                 sSQL = "SpTACuentasBorrar @Codigo='" & TxtCodigo.Text & "'"
                 Set Tabla = v.db.OpenResultset(sSQL)
                 MsgBox "El registro fue borrado correctamente", vbInformation + vbOKOnly
                 BorrarTodo
                 TxtCodigo.SetFocus
                 CargarCuentasContables v.db
            End If
        End If
    End If
End Sub

Private Function ValidarBorrar() As Boolean
On Error GoTo Errores

Dim TbTabla As rdoResultset
Dim sSQL As String
Dim i As Integer
Dim j As Integer
Dim Continuar As Boolean
Dim CambioNumero As Boolean
    ValidarBorrar = True
        sSQL = "SpTACuentasValidarBorrar @Codigo='" & TxtCodigo.Text & "'"
        Set TbTabla = v.db.OpenResultset(sSQL)
        If UCase(TbTabla!Mensaje) <> "OK" Then
            MsgBox TbTabla!Mensaje, 16
            ValidarBorrar = False
            Exit Function
        End If
Errores:
    ManipularError Err.Number, Err.Description
End Function

Private Function Validar() As Boolean
On Error GoTo Errores
Dim TbTabla As rdoResultset
Dim sSQL As String
    Validar = True
        If TxtCodigo.Text = "" Then
            MsgBox " Debe Ingresar un Codigo", 16
            Validar = False
            Exit Function
        End If
        If TxtDescripcion.Text = "" Then
            MsgBox " Debe Ingresar una Descripcion", 16
            Validar = False
            Exit Function
        End If
        If TxtReducida.Text = "" Then
            MsgBox " Debe Ingresar una Descripcion Reducida", 16
            Validar = False
            Exit Function
        End If
        
        sSQL = "SpTACuentasValidarAgregar @Codigo='" & TxtCodigo.Text & "', " & _
                                    " @Modificado=" & BooleanoSQL2(Str(v.Modificado), v.Servidor)
        Set TbTabla = v.db.OpenResultset(sSQL)
        If UCase(TbTabla!Mensaje) <> "OK" Then
            MsgBox TbTabla!Mensaje
            Validar = False
            Exit Function
        End If
Errores:
    ManipularError Err.Number, Err.Description
End Function

Private Sub CMDBuscar_Click()
On Error GoTo ManejoError
    BuscarCuentas.Codigo.Caption = " "
    BuscarCuentas.Descripcion.Caption = " "
    BuscarCuentas.Reducida.Caption = ""
    BuscarCuentas.CargarParametros LbConexion.Caption
    BuscarCuentas.Show vbModal
ManejoError:
    Timer2.Enabled = True
End Sub

Private Sub CMDGuardar_Click()
    Call GuardarCambios
End Sub

Private Sub GuardarCambios()
On Error GoTo Errores
Dim TbGuardar As rdoResultset
Dim sSQL As String
Dim Pregunta As Integer
    ComenzarTransaccion v.db, v.Servidor
    Pregunta = MsgBox("¿Desea grabar los datos?", vbQuestion + vbOKCancel, "PulquiPack")
    If Pregunta = vbOK Then
        If Validar Then
           If Not v.Modificado Then
               sSQL = "SpTACuentasAgregar @Codigo='" & TxtCodigo.Text & "', " & _
                                "@Descripcion='" & Trim(TxtDescripcion.Text) & "', " & _
                                "@Reducida='" & Trim(TxtReducida.Text) & "'"
               Set TbGuardar = v.db.OpenResultset(sSQL, 2, 3)
               MsgBox "Los datos fueron grabados correctamente", vbInformation + vbOKOnly
               Call BorrarTodo
               TxtCodigo.SetFocus
           Else
               sSQL = "SpTACuentasModificar @Codigo='" & TxtCodigo.Text & "', " & _
                                "@Descripcion='" & Trim(TxtDescripcion.Text) & "', " & _
                                "@Reducida='" & Trim(TxtReducida.Text) & "'"
                                
               Set TbGuardar = v.db.OpenResultset(sSQL)
               MsgBox "Los datos fueron modificados correctamente", vbInformation + vbOKOnly
               Call BorrarTodo
               TxtCodigo.SetFocus
           End If
           Call CargarCuentasContables(v.db)
           'Call CargarVecCuentasContables
           'Call CargarVecCuentasContables
        End If
     End If
    TerminarTransaccion v.db, v.Servidor
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub Timer2_Timer()
Dim BancoAnterior As String
Dim SucursalAnterior As String
Dim VariableBoba As Boolean
        If Trim(BuscarCuentas.Codigo.Caption) <> "" Then
            TxtCodigo.Text = BuscarCuentas.Codigo.Caption
            TxtDescripcion.Text = BuscarCuentas.Descripcion.Caption
            TxtReducida.Text = BuscarCuentas.Reducida.Caption
            v.Modificado = True
        End If
            Timer2.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LbConexion.Caption = "servidor=sql;dsn=ElPulqui1;uid=todos;PWD=todos"
    InicializarTodo
End Sub

Private Sub TXTCodigo_LostFocus()
Dim sSQL As String
Dim Tabla As rdoResultset
    sSQL = "SpTACuentasRecuperar @Codigo='" & TxtCodigo.Text & "'"
    Set Tabla = v.db.OpenResultset(sSQL)
            
    If Not Tabla.EOF Then
        TxtDescripcion.Text = Trim(Tabla!C_Descripcion)
        TxtReducida.Text = Trim(Tabla!C_Reducida)
        v.Modificado = True
    Else
        TxtDescripcion.Text = ""
        TxtReducida.Text = ""
        v.Modificado = False
   End If
    
End Sub

Private Sub BorrarTodo()
    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
    TxtReducida.Text = ""
End Sub

Private Sub Txtcodigo_GotFocus()
On Error GoTo Errores
    SelText TxtCodigo
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub TxtDescripcion_GotFocus()
On Error GoTo Errores
    SelText TxtDescripcion
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub LbConexion_Change()
On Error GoTo Errores
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
    
    v.Servidor = UCase(BuscarString(LbConexion.Caption, "servidor="))
    Set v.db = v.Env.OpenConnection(NombreConexion, rdDriverNoPrompt, False, Opciones)
    InicializarTodo
    MousePointer = vbNormal

Errores:
    ManipularError Err.Number, Err.Description
    
End Sub

Private Sub InicializarTodo()
On Error GoTo Errores
    MousePointer = vbHourglass
    InicializarTags
    ReDim Vi.VecPosiciones(2)
    Vi.VecPosiciones(0) = 1
    Vi.VecPosiciones(1) = 15
    Vi.VecPosiciones(2) = 40
    Vi.ancho = 100
    Vi.TamanioLetra = 10
    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
    TxtReducida.Text = ""
    v.Modificado = False
    
    MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub InicializarTags()
On Error GoTo Errores
    TxtCodigo.Tag = Repetir("X", 4)
    TxtDescripcion.Tag = Repetir("X", 50)
    TxtReducida.Tag = Repetir("X", 30)
    'un tag con cero es para dígitos y la coma y el punto
Errores:
    ManipularError Err.Number, Err.Description
End Sub



