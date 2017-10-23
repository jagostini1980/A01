VERSION 5.00
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form FrmCrearArticulo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Artículos"
   ClientHeight    =   1830
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Artículo"
      Height          =   1635
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7020
      Begin Controles.ComboEsp CmbCuentaContable 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   720
         Width           =   5190
         _ExtentX        =   9155
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
      Begin VB.CommandButton CMDSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3735
         TabIndex        =   3
         Top             =   1125
         Width           =   1200
      End
      Begin VB.CommandButton CMDGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2295
         TabIndex        =   2
         Top             =   1125
         Width           =   1200
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   315
         Width           =   5190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Por Defecto:"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label LbDescripcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   660
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmCrearArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VecCuentasContables() As CuentasContables

Private Sub CMDGuardar_Click()
    Call GuardarCambios
End Sub

Private Sub GuardarCambios()
On Error GoTo errores
Dim Sql As String
Dim Pregunta As Integer
    
    Pregunta = MsgBox("¿Desea grabar los datos?", vbQuestion + vbOKCancel, "PulquiPack")
    If Pregunta = vbOK Then
        Dim RsGuardar As New ADODB.Recordset
        
     ' si A_Codigo =0 es un alta
       If ArtCompraActual.A_Codigo = 0 Then
          Sql = "SpTaArticulosComprasAgregaryRelacionar @A_Descripcion='" & TxtDescripcion.Text & _
                                                    "', @A_CuentaPordefecto ='" & VecCuentasContables(CmbCuentaContable.ListIndex).Codigo & _
                                                    "', @CentroDeCostos ='" & CentroEmisor & "'"

          RsGuardar.Open Sql, Conec
          If RsGuardar!OK = "OK" Then
                MsgBox RsGuardar!Mensaje
                Unload Me
          Else
                MsgBox RsGuardar!Mensaje, vbInformation
                TxtDescripcion.SetFocus
                TxtDescripcion.SelText = TxtDescripcion.Text
                Exit Sub
          End If
               
       End If
           
    End If
    'limpia los campos
    TxtDescripcion.Text = ""
    CmbCuentaContable.ListIndex = 0
    Call CargarVecArtCompras
errores:
    If Err.Number <> 0 Then
       ManipularError Err.Number, Err.Description
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim RsCuentas As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
    
      ReDim VecCuentasContables(0)
      'en esta sección carga las cuentas
    With RsCuentas
        Sql = "SpOCRelacionCentroDeCostoCuentaContable @R_CentroDeCosto='" & CentroEmisor & "'"
        .Open Sql, Conec, adOpenStatic, adLockReadOnly
        
          For i = 1 To UBound(Ayudas.VecCuentasContables)
              .Find "R_CuentaContable = " & Ayudas.VecCuentasContables(i).Codigo, , , 1
             If Not .EOF Then
                ReDim Preserve VecCuentasContables(UBound(VecCuentasContables) + 1)
                VecCuentasContables(UBound(VecCuentasContables)) = Ayudas.VecCuentasContables(i)
             End If
          Next
    End With
    Call CargarCmbCuentasContables(CmbCuentaContable)
End Sub

Private Sub TxtDescripcion_Change()
    Call ColorObligatorio(TxtDescripcion, CMDGuardar)
End Sub

Private Sub TxtDescripcion_GotFocus()
On Error GoTo errores
    SelText TxtDescripcion
errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCmbCuentasContables(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo errores
Dim i As Integer

    Cmb.Clear
    
    If Tipo = "Elegir" Then
       Cmb.AddItem "Seleccione una Cuenta Contable"
    Else
       Cmb.AddItem "Todas las Cuentas Contables"
    End If

    For i = 1 To UBound(VecCuentasContables)
        Cmb.AddItem Trim(VecCuentasContables(i).Descripcion)
    Next
        
    Cmb.ListIndex = 0
errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub UbicarCmbCuentas(Codigo As String, CbCuentasContables As ComboEsp)
Dim i As Integer

    CbCuentasContables.ListIndex = 0
    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = Codigo Then
            CbCuentasContables.ListIndex = i
            Exit For
        End If
    Next
End Sub
