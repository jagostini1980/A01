VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Ingreso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmUsuario 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Identifidacion de Usuario "
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin MSComctlLib.ProgressBar PrBar 
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2513
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1095
         TabIndex        =   5
         Top             =   1665
         Width           =   1095
      End
      Begin VB.TextBox TxtContraseña 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label LBVersion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Version: 1.0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3375
         TabIndex        =   9
         Top             =   2880
         Width           =   840
      End
      Begin VB.Label LbContraseña 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contraseña:"
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
         Left            =   780
         TabIndex        =   3
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label LbUsuario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario:"
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
         Left            =   1095
         TabIndex        =   1
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.Label LbConexion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Ingreso"
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

Dim v As VariablesGlobales

Private Sub CmdAceptar_Click()
Dim TbTabla As rdoResultset
Dim sSQL As String
Dim RsNivel As New ADODB.Recordset

    If Validar Then
        
        sSQL = "SpACExisteUsuario @Usuario='" & TxtUsuario.Text & "'"

        Set TbTabla = v.db.OpenResultset(sSQL)
        
        If Not TbTabla.EOF Then
            If Trim(TxtContraseña.Text) <> Trim(TbTabla!U_Contrasena) Then
                MsgBox "La contraseña no corresponde"
            Else
        'acá inicialozo la donección ADO que se
        'utilizará en todo el sistema exepto en
        'las ventanas que ya existen y trabajan con RDO
                Call InicializarConexionADO
        'pone el usuario en una variable Global para luego
        'obtener el Nivel de la tabla AC_AccesosDeGruposAModulos
        'en la col N_Nivel
                Usuario = TxtUsuario.Text
                NombreUsuario = TbTabla!U_Nombre
                Call DescargarComponentes
                Call CargarVectores
                
                Unload Ingreso
                MenuEmisionOrdenCompra.Show

            End If
        Else
            MsgBox "El Usuario no existe"
        End If
    End If
End Sub

Private Sub DescargarComponentes()
On Error GoTo ErrorFtp
  Dim DirSystem As String
  Dim mFTP As cFTP
  Set mFTP = New cFTP
  
  mFTP.SetModeActive
  mFTP.SetTransferBinary
  DirSystem = Obtiene_WinDIR
  
  MousePointer = vbHourglass
  
  If Not FileExist(DirSystem & "\sg20.ocx") Then
     If mFTP.OpenConnection("svrppack.dyndns.org", "", "") Then
        Call mFTP.SetFTPDirectory("Componentes")
        If mFTP.FTPDownloadFile(DirSystem & "\" & "sg20.ocx", "sg20.ocx") Then
           Call SelfRegisterDLL(DirSystem & "\" & "sg20.ocx")
        Else
           MsgBox "Error al Descargar sg20.ocx", vbCritical, "Error de Actualizacion"
        End If
     Else
        MsgBox "Error al intentar conectarse al Servidor", vbCritical, "Error de Actualizacion"
     End If
  End If
  
  If Not FileExist(DirSystem & "\mswinsck.ocx") Then
     If mFTP.OpenConnection("svrppack.dyndns.org", "", "") Then
        Call mFTP.SetFTPDirectory("Componentes")
        If mFTP.FTPDownloadFile(DirSystem & "\" & "mswinsck.ocx", "mswinsck.ocx") Then
           Call SelfRegisterDLL(DirSystem & "\" & "mswinsck.ocx")
        Else
           MsgBox "Error al Descargar mswinsck.ocx", vbCritical, "Error de Actualizacion"
        End If
     Else
        MsgBox "Error al intentar conectarse al Servidor", vbCritical, "Error de Actualizacion"
     End If
  End If
  
  If Not FileExist(DirSystem & "\vbSendMail.dll") Then
     If mFTP.OpenConnection("svrppack.dyndns.org", "", "") Then
        Call mFTP.SetFTPDirectory("Componentes")
        If mFTP.FTPDownloadFile(DirSystem & "\" & "vbSendMail.dll", "vbSendMail.dll") Then
           Call SelfRegisterDLL(DirSystem & "\" & "vbSendMail.dll")
        Else
           MsgBox "Error al Descargar vbSendMail.dll", vbCritical, "Error de Actualizacion"
        End If
     Else
        MsgBox "Error al intentar conectarse al Servidor", vbCritical, "Error de Actualizacion"
     End If
  End If
  
  If Not FileExist(DirSystem & "\MSMAPI32.OCX") Then
     If mFTP.OpenConnection("svrppack.dyndns.org", "", "") Then
        Call mFTP.SetFTPDirectory("Componentes")
        If mFTP.FTPDownloadFile(DirSystem & "\" & "MSMAPI32.OCX", "MSMAPI32.OCX") Then
           Call SelfRegisterDLL(DirSystem & "\" & "MSMAPI32.OCX")
        Else
           MsgBox "Error al Descargar MSMAPI32.OCX", vbCritical, "Error de Actualizacion"
        End If
     Else
        MsgBox "Error al intentar conectarse al Servidor", vbCritical, "Error de Actualizacion"
     End If
  End If
  
ErrorFtp:
    MousePointer = vbNormal
    Call ManipularError(Err.Number, Err.Description)
End Sub


Private Function Validar() As Boolean
On Error GoTo Errores

    Validar = True
    If TxtUsuario.Text = "" Then
        MsgBox " Debe Ingresar un usuario", 16
        Validar = False
        Exit Function
    End If
    If TxtContraseña.Text = "" Then
        MsgBox " Debe Ingresar una Contraseña", 16
        Validar = False
        Exit Function
    End If
        
Errores:
    ManipularError Err.Number, Err.Description
End Function

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LbConexion.Caption = "servidor=sql;dsn=ElPulqui1;uid=todos;PWD=todos"
    PrBar.Align = vbAlignBottom
    PrBar.Visible = False
    LBVersion = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision
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
    'TxtUsuario.Text = ""
    'TxtContraseña.Text = ""
    v.Modificado = False
    MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub CargarVectores()
'On Error GoTo Errores
    MousePointer = vbHourglass
    PrBar.Min = 1
    PrBar.Max = 25
    PrBar.Visible = True
    'Establece Min como valor de Value.
    PrBar.Value = PrBar.Min
      Call CargarEmailAutorizacion
    PrBar.Value = 2
      Call CargarVecRubrosContables
    PrBar.Value = 3
      Call CargarEmpresas(v.db)
    PrBar.Value = 4
      Call CargarVecCentrosDeCostos
    PrBar.Value = 5
      Call CargarVecArtTaller
    PrBar.Value = 6
      Call CargarParametrosSeguro
    PrBar.Value = 7
      Call CargarVecCentrosDeCostosEmisor
    PrBar.Value = 8
      Call CargarProveedores(v.db)
    PrBar.Value = 9
      Call CargarVecUsuarios
    PrBar.Value = 10
      Call CargarVecCentrosDeCostosNivel2
    PrBar.Value = 11
      Call CargarCuentasContables(v.db)
    PrBar.Value = 12
        If CentroEmisor = "" Then
            CentroEmisor = BuscarAuxiliar(Usuario)
        End If
        MaxSinAutorizacion = VecCentroDeCostoEmisor(BuscarIndexCentroEmisor(CentroEmisor)).C_ImporteAutorizar
    PrBar.Value = 13
        Call CargarVecArtMotoVan
    PrBar.Value = 14
        Call CargarVecArtCompras
    PrBar.Value = 15
        Call CargarCoches
        Call CargarTiposDeCoche
    PrBar.Value = 16
        Call CargarVecLineas
        Call CargarVecProvMinicenas
    PrBar.Value = 17
        Call CargarVecLugaresDeEntrega
    PrBar.Value = 18
        Call CargarRubros(v.db)
    PrBar.Value = 19
        Call CargarGrupos(v.db)
    PrBar.Value = 20
        Call CargarUnidadesDeMedida(v.db)
    PrBar.Value = 21
        Call CargarMarcas(v.db)
    PrBar.Value = 22
        Call CargarVecFormasDePago
    PrBar.Value = 23
        Call CargarTalleres(v.db)
    PrBar.Value = 24
        Call CargarVecUnidadesDeNegocio
    PrBar.Value = 25
        Call CargarVecPaquetes

    PrBar.Visible = False
    PrBar.Value = PrBar.Min
    MousePointer = vbNormal
    
Errores:
    'MsgBox "Vector " & PrBar.Value
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub InicializarTags()
On Error GoTo Errores
    TxtUsuario.Tag = Repetir("X", 15)
    TxtContraseña.Tag = Repetir("X", 10)
    
    'un tag con cero es para dígitos y la coma y el punto
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub



Private Sub UserControl_KeyPress(KeyAscii As Integer)
    TeclaPresionada ActiveControl, KeyAscii
End Sub





