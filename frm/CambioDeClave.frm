VERSION 5.00
Begin VB.Form CambioDeClave 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Clave de Acceso"
   ClientHeight    =   2775
   ClientLeft      =   5670
   ClientTop       =   4560
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2250
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Frame FrmCambiodeClave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Clave de Acceso"
      Height          =   1965
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4440
      Begin VB.TextBox TxtConfirmarContraseña 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1350
         Width           =   2000
      End
      Begin VB.TextBox TxtContraseñaActual 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   2000
      End
      Begin VB.TextBox TxtNuevaContraseña 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   855
         Width           =   2000
      End
      Begin VB.Label LbConfirmarcontraeña 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Confirmar Contraseña:"
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
         Left            =   150
         TabIndex        =   6
         Top             =   1440
         Width           =   1890
      End
      Begin VB.Label LbContraseñaActual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contraseña Actual:"
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
         Left            =   405
         TabIndex        =   5
         Top             =   435
         Width           =   1635
      End
      Begin VB.Label LbNuevaContraseña 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nueva Contraseña:"
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
         Left            =   390
         TabIndex        =   4
         Top             =   930
         Width           =   1650
      End
   End
End
Attribute VB_Name = "CambioDeClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()

Dim TbTabla As ADODB.Recordset
Set TbTabla = New ADODB.Recordset
Dim sSQL As String
On Error GoTo ErrorCambio
    If Validar Then
        
        sSQL = "SpACExisteUsuario @Usuario='" & Usuario & "'"

         TbTabla.Open sSQL, Conec
        
        If Not TbTabla.EOF Then
            If Trim(TxtContraseñaActual.Text) <> Trim(TbTabla!U_Contrasena) Then
                MsgBox "La contraseña actual es correcta"
                Exit Sub
            Else
                sSQL = "SpACModificarContraseña @Usuario='" & Usuario & "', " & _
                                "@Contraseña='" & TxtNuevaContraseña.Text & "'"
                                
                Conec.Execute sSQL
            End If
        Else
            MsgBox "El Usuario no existe"
        End If
    Else
        Exit Sub
    End If
ErrorCambio:
     If Err.Number = 0 Then
        MsgBox "La Contraseña ha sido actualizada correctamente", vbInformation
     Else
        MsgBox "Error de actualización", vbCritical
     End If
        Unload Me

End Sub

Private Function Validar() As Boolean
On Error GoTo Errores

        Validar = True
        If Trim(TxtConfirmarContraseña.Text) <> Trim(TxtNuevaContraseña.Text) Then
            MsgBox " La Nueva contraseña no coincide con la confirmación de contraseña", 16
            Validar = False
            Exit Function
        End If
        
Errores:
    ManipularError Err.Number, Err.Description
End Function

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub TxtConfirmarContraseña_Change()
    Call ColorObligatorio(TxtConfirmarContraseña, CmdAceptar)
End Sub

Private Sub TxtContraseñaActual_Change()
    Call ColorObligatorio(TxtContraseñaActual, CmdAceptar)
End Sub

Private Sub TxtNuevaContraseña_Change()
    Call ColorObligatorio(TxtNuevaContraseña, CmdAceptar)

End Sub
