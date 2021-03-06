VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form IngresoSeg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
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
         Left            =   1073
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TxtContrase�a 
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
      Begin VB.Label LbContrase�a 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contrase�a:"
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
Attribute VB_Name = "IngresoSeg"
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
Dim RsNivel As ADODB.Recordset
Set RsNivel = New ADODB.Recordset

    If Validar Then
        
        sSQL = "SpACExisteUsuario @Usuario='" & TxtUsuario.Text & "'"

        Set TbTabla = v.db.OpenResultset(sSQL)
        
        If Not TbTabla.EOF Then
            If Trim(TxtContrase�a.Text) <> Trim(TbTabla!U_Contrasena) Then
                MsgBox "La contrase�a no corresponde"
            Else
        'ac� inicialozo la donecci�n ADO que se
        'utilizar� en todo el sistema exepto en
        'las ventanas que ya existen y trabajan con RDO
                Call InicializarConexionADO
        'pone el usuario en una variable Global para luego
        'obtener el Nivel de la tabla AC_AccesosDeGruposAModulos
        'en la col N_Nivel
                Usuario = TxtUsuario.Text
                NombreUsuario = TbTabla!U_Nombre
                
                Call CargarVectores
                
                Unload IngresoSeg
                MenuA01Seg.Show

            End If
        Else
            MsgBox "El Usuario no existe"
        End If
    End If
End Sub

Private Function Validar() As Boolean
On Error GoTo Errores

    Validar = True
    If TxtUsuario.Text = "" Then
        MsgBox " Debe Ingresar un usuario", 16
        Validar = False
        Exit Function
    End If
    If TxtContrase�a.Text = "" Then
        MsgBox " Debe Ingresar una Contrase�a", 16
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
    LBVersion = "Versi�n: " & App.Major & "." & App.Minor & "." & App.Revision
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
    'TxtContrase�a.Text = ""
    v.Modificado = False
    MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub CargarVectores()
    MousePointer = vbHourglass
    PrBar.Min = 1
    PrBar.Max = 16
    PrBar.Visible = True
    'Establece Min como valor de Value.
    PrBar.Value = PrBar.Min
    
    PrBar.Value = 2
      'Call CargarVecArtCompras
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
    PrBar.Value = 13
        'Call CargarVecArtMotoVan
    PrBar.Value = 14
        'Call CargarVecArtCompras
    PrBar.Value = 15
        Call CargarCoches
    PrBar.Value = 16
        Call CargarTiposDeCoche
    
    PrBar.Visible = False
    PrBar.Value = PrBar.Min
    MousePointer = vbNormal
    
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub InicializarTags()
On Error GoTo Errores
    TxtUsuario.Tag = Repetir("X", 15)
    TxtContrase�a.Tag = Repetir("X", 10)
    
    'un tag con cero es para d�gitos y la coma y el punto
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub



Private Sub UserControl_KeyPress(KeyAscii As Integer)
    TeclaPresionada ActiveControl, KeyAscii
End Sub





