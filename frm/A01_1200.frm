VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form A01_1200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   4515
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedores"
      Height          =   4155
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7455
      Begin MSMask.MaskEdBox TxtCuit 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   3165
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.TextBox TxtTelefono 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   7
         Text            =   " "
         Top             =   2724
         Width           =   4095
      End
      Begin VB.TextBox TxtCodigoPostal 
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2330
         Width           =   1215
      End
      Begin VB.TextBox TxtLocalidad 
         Height          =   315
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1936
         Width           =   3495
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1542
         Width           =   4095
      End
      Begin VB.TextBox TxtReducida 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1148
         Width           =   2535
      End
      Begin VB.CommandButton CMDSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5580
         TabIndex        =   12
         Top             =   3570
         Width           =   1455
      End
      Begin VB.CommandButton CMDGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Guardar"
         Height          =   375
         Left            =   315
         TabIndex        =   9
         Top             =   3555
         Width           =   1455
      End
      Begin VB.CommandButton CMDBuscar 
         Caption         =   "Buscar "
         Height          =   375
         Left            =   3825
         TabIndex        =   11
         Top             =   3570
         Width           =   1455
      End
      Begin VB.CommandButton CMDBorrar 
         Appearance      =   0  'Flat
         Caption         =   "Borrar"
         Height          =   375
         Left            =   2070
         TabIndex        =   10
         Top             =   3570
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   754
         Width           =   4815
      End
      Begin VB.TextBox TxtCodigo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LbCuit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "CUIT:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1155
         TabIndex        =   21
         Top             =   3195
         Width           =   420
      End
      Begin VB.Label LbTelefono 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Teléfono:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   900
         TabIndex        =   20
         Top             =   2790
         Width           =   675
      End
      Begin VB.Label LbCodigoPostal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo Postal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   555
         TabIndex        =   19
         Top             =   2385
         Width           =   1020
      End
      Begin VB.Label LbLocalidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Localidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   18
         Top             =   2025
         Width           =   735
      End
      Begin VB.Label LbDireccion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dirección:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   855
         TabIndex        =   17
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label LbReducida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reducida:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label LbDescripcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razón Social:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   585
         TabIndex        =   14
         Top             =   810
         Width           =   990
      End
      Begin VB.Label LbCodigo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   1035
         TabIndex        =   13
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.Label LbConexion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   765
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "A01_1200"
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
                 sSQL = "SpTAProveedoresBorrar @Codigo=" & TxtCodigo.Text
                 Set Tabla = v.db.OpenResultset(sSQL)
                 MsgBox "El registro fue borrado correctamente", vbInformation + vbOKOnly
                 BorrarTodo
                 TxtCodigo.SetFocus
                 TraerUltimoProveedor
                 CargarProveedores v.db
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
        sSQL = "SpTAProveedoresValidarBorrar @Codigo=" & TxtCodigo.Text
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
        If TxtCodigo.Text = "" Or Val(TxtCodigo.Text) = 0 Then
            MsgBox " Debe Ingresar un Codigo", 16
            Validar = False
            Exit Function
        End If
        If Not IsNumeric(TxtCodigo.Text) Then
            MsgBox " Debe Ingresar un Codigo Numerico", 16
            Validar = False
            Exit Function
        End If

        If TxtDescripcion.Text = "" Then
            MsgBox " Debe Ingresar una Descripcion", 16
            Validar = False
            Exit Function
        End If
        If Trim(TxtCuit.Text) <> "" Then
            If Not Valcuit(TxtCuit.Text) Then
               MsgBox " El CUIT es invalido", 16
               Validar = False
               Exit Function
            End If
        End If

        sSQL = "SpTAProveedoresValidarAgregar @Codigo=" & TxtCodigo.Text & ", " & _
                                    " @Modificado=" & BooleanoSQL2(Str(v.Modificado), v.Servidor)
        Set TbTabla = v.db.OpenResultset(sSQL)
        If UCase(TbTabla!Mensaje) <> "OK" Then
            MsgBox TbTabla!Mensaje
            Validar = False
            Exit Function
        End If
        sSQL = "SpTAProveedoresValidarAgregarCuit @Cuit='" & TxtCuit.Text & "', " & _
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
    BuscarProveedores.Codigo.Caption = " "
    BuscarProveedores.Descripcion.Caption = " "
    BuscarProveedores.Reducida.Caption = ""
    BuscarProveedores.Direccion.Caption = ""
    BuscarProveedores.Localidad.Caption = ""
    BuscarProveedores.CodigoPostal.Caption = ""
    BuscarProveedores.Telefono.Caption = ""
    BuscarProveedores.CUIT.Caption = ""
    BuscarProveedores.CargarParametros LbConexion.Caption
    BuscarProveedores.Show vbModal
ManejoError:
    Timer2.Enabled = True
End Sub

Private Sub CMDGuardar_Click()
    GuardarCambios
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
               sSQL = "SpTAProveedoresAgregar @Codigo=" & TxtCodigo.Text & ", " & _
                                " @Descripcion='" & Trim(TxtDescripcion.Text) & "', " & _
                                " @Reducida='" & Trim(TxtReducida.Text) & "', " & _
                                " @Direccion='" & TxtDireccion.Text & "', " & _
                                " @Localidad='" & TxtLocalidad.Text & "', " & _
                                " @CodigoPostal='" & TxtCodigoPostal.Text & "', " & _
                                " @Telefono='" & TxtTelefono.Text & "', " & _
                                " @Cuit='" & TxtCuit.Text & "'"
               Set TbGuardar = v.db.OpenResultset(sSQL, 2, 3)
               MsgBox "Los datos fueron grabados correctamente", vbInformation + vbOKOnly
               BorrarTodo
               TxtCodigo.SetFocus
               TraerUltimoProveedor
           Else
               sSQL = "SpTAProveedoresModificar @Codigo=" & TxtCodigo.Text & ", " & _
                                " @Descripcion='" & Trim(TxtDescripcion.Text) & "', " & _
                                " @Reducida='" & Trim(TxtReducida.Text) & "', " & _
                                " @Direccion='" & TxtDireccion.Text & "', " & _
                                " @Localidad='" & TxtLocalidad.Text & "', " & _
                                " @CodigoPostal='" & TxtCodigoPostal.Text & "', " & _
                                " @Telefono='" & TxtTelefono.Text & "', " & _
                                " @Cuit='" & TxtCuit.Text & "'"
               Set TbGuardar = v.db.OpenResultset(sSQL)
               MsgBox "Los datos fueron modificados correctamente", vbInformation + vbOKOnly
               BorrarTodo
               TxtCodigo.SetFocus
               TraerUltimoProveedor
           End If
           CargarProveedores v.db
        End If
     End If
    TerminarTransaccion v.db, v.Servidor
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub Timer2_Timer()
Dim VariableBoba As Boolean
        If Trim(BuscarProveedores.Codigo.Caption) <> "" Then
            TxtCodigo.Text = BuscarProveedores.Codigo.Caption
            TxtDescripcion.Text = BuscarProveedores.Descripcion.Caption
            TxtReducida.Text = BuscarProveedores.Reducida.Caption
            TxtDireccion.Text = BuscarProveedores.Direccion.Caption
            TxtLocalidad.Text = BuscarProveedores.Localidad.Caption
            TxtCodigoPostal.Text = BuscarProveedores.CodigoPostal.Caption
            TxtTelefono.Text = BuscarProveedores.Telefono.Caption
            TxtCuit.Text = BuscarProveedores.CUIT.Caption
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
    If IsNumeric(TxtCodigo.Text) And Val(TxtCodigo.Text) <> 0 Then
        sSQL = "SpTAProveedoresRecuperar @Codigo=" & TxtCodigo.Text
        Set Tabla = v.db.OpenResultset(sSQL)
      With Tabla
        If Not .EOF Then
            TxtDescripcion.Text = Trim(!P_Descripcion)
            TxtReducida.Text = Trim(!P_Reducida)
            TxtDireccion.Text = VerificarNulo(!P_Direccion, "N")
            TxtLocalidad.Text = VerificarNulo(!P_Localidad)
            TxtCodigoPostal.Text = VerificarNulo(!P_CodigoPostal)
            TxtTelefono.Text = Trim(!P_Telefono)
            TxtCuit.Text = Trim(!P_Cuit)
            v.Modificado = True
        Else
            TxtDescripcion.Text = ""
            TxtReducida.Text = ""
            TxtDireccion.Text = ""
            TxtLocalidad.Text = ""
            TxtCodigoPostal.Text = ""
            TxtTelefono.Text = ""
            TxtCuit.Text = ""
            v.Modificado = False
        End If
      End With
    Else
        If Not IsNumeric(TxtCodigo.Text) Then
           MsgBox " Debe Ingresar un Codigo Numerico", 16
           TxtCodigo.Text = 0
           TxtCodigo.SetFocus
        End If
    End If

    
End Sub
Private Sub BorrarTodo()
    TxtCodigo.Text = 0
    TxtDescripcion.Text = ""
    TxtReducida.Text = ""
    TxtDireccion.Text = ""
    TxtLocalidad.Text = ""
    TxtCodigoPostal.Text = ""
    TxtTelefono.Text = ""
    TxtCuit.Text = ""
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
    TxtCodigo.Text = 0
    TraerUltimoProveedor
    TxtDescripcion.Text = ""
    TxtReducida.Text = ""
    TxtDireccion.Text = ""
    TxtLocalidad.Text = ""
    TxtCodigoPostal.Text = ""
    TxtTelefono.Text = ""
    TxtCuit.Text = ""
    
    v.Modificado = False
    
    MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub
Private Sub TraerUltimoProveedor()
Dim sSQL As String
Dim Tabla As rdoResultset
    sSQL = "SpTAProveedoresTraerUltimo"
    Set Tabla = v.db.OpenResultset(sSQL)
           
    If Not Tabla.EOF Then
       TxtCodigo.Text = Trim(Tabla!UltimoProveedor) + 1
    End If
End Sub
Private Sub InicializarTags()
On Error GoTo Errores
    TxtCodigo.Tag = Repetir("0", 4)
    TxtDescripcion.Tag = Repetir("X", 40) 'un tag con X es para que se pueda ingresar cualquier dígito
    TxtReducida.Tag = Repetir("X", 20)
    TxtDireccion.Tag = Repetir("X", 40)
    TxtLocalidad.Tag = Repetir("X", 30)
    TxtCodigoPostal.Tag = Repetir("X", 8)
    TxtTelefono.Tag = Repetir("X", 40)
    TxtCuit.Mask = "99-99999999-9"
    'un tag con cero es para dígitos y la coma y el punto
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    TeclaPresionada ActiveControl, KeyAscii
End Sub



