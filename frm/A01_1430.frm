VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_1430 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Usuarios Auxiliares"
   ClientHeight    =   8190
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Centro de costo"
      Height          =   1410
      Left            =   87
      TabIndex        =   3
      Top             =   45
      Width           =   4155
      Begin Controles.ComboEsp CmbCentroDeCosto 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   450
         Width           =   3930
         _ExtentX        =   6932
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
      Begin VB.TextBox TxtUsuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   990
         Width           =   3930
      End
      Begin VB.Label LbDescripcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo:"
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
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label LBUsuario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ususario Responsable:"
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
         Left            =   135
         TabIndex        =   4
         Top             =   765
         Width           =   1965
      End
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   833
      TabIndex        =   1
      Top             =   7695
      Width           =   1230
   End
   Begin VB.CommandButton CMDSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2228
      TabIndex        =   2
      Top             =   7695
      Width           =   1200
   End
   Begin MSComctlLib.ListView LvUsuarios 
      Height          =   6060
      Left            =   90
      TabIndex        =   0
      Top             =   1530
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   10689
      View            =   3
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
End
Attribute VB_Name = "A01_1430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private Nivel As Integer
Private ListIndex As Integer
Public CentroDeCastoActual As String

Private Sub CmbCentroDeCosto_Click()
    
    If Modificado Then
       Dim Rta As Integer
       Rta = MsgBox("Ha efectuado cambio �Desea Guardarlos?", vbYesNo)
       
       If Rta = vbYes Then
          Call GuardarCambios
       End If
    End If
    TxtUsuario.Text = VecCentroDeCostoEmisor(CmbCentroDeCosto.ListIndex).C_UsuarioResponsable
    Call CheckAuxiliar(VecCentroDeCostoEmisor(CmbCentroDeCosto.ListIndex).C_Codigo)
    ListIndex = CmbCentroDeCosto.ListIndex
    Modificado = False
End Sub

Private Sub CmdConfirmar_Click()
 Dim Sql As String
 Dim Pregunta As Integer
 Dim i As Integer
  Pregunta = MsgBox("�Desea Modificar?", vbQuestion + vbOKCancel, "Pulqui")
  
  If Pregunta = vbOK Then
        Call GuardarCambios
  End If
End Sub

Private Sub GuardarCambios()
  Dim Sql As String
  Dim i As Integer
  
 On Error GoTo Error
     
     Conec.BeginTrans
        'ACTUALIZA LAS RELACIONES DE LOS AUXILIARES
        Sql = "SpTaAuxiliaresCentosDeCostosPorCentroDeCostoBorrar @A_CentroDeCosto = '" & VecCentroDeCostoEmisor(ListIndex).C_Codigo & "'"
        Conec.Execute Sql
        For i = 1 To LvUsuarios.ListItems.Count
            If LvUsuarios.ListItems(i).Checked And Not LvUsuarios.ListItems(i).Bold _
               And LvUsuarios.ListItems(i).ForeColor = vbBlack Then
                Sql = "SpTaAuxiliaresCentosDeCostosPorCentroDeCostoAgregar " & _
                         "@A_CentroDeCosto ='" & VecCentroDeCostoEmisor(ListIndex).C_Codigo & _
                      "', @A_Usuario = '" & LvUsuarios.ListItems(i).Text & "'"
                      
                Conec.Execute Sql
            End If
        Next
        
        Modificado = False

     Conec.CommitTrans
Error:
    If Err.Number = 0 Then
        MsgBox "La modificaci�n se realiz� correctamente", , "Modificaci�n"
        Modificado = False
    Else
        Conec.RollbackTrans
        Call ManipularError(Err.Number, Err.Description)
    End If

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'TxtUsuario.Text = Usuario
    Call CrearEmcabezados
    'Nivel = TraerNivel("A011430")
'cargo los usuarios en el Lv
    For i = 1 To UBound(VecUsuarios)
        LvUsuarios.ListItems.Add , VecUsuarios(i), VecUsuarios(i)
    Next
    
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCosto)
    Call BuscarCentroEmisor(CentroDeCastoActual, CmbCentroDeCosto)

    CmbCentroDeCosto.Enabled = False ' Nivel = 2
    Modificado = False
End Sub

Private Sub CrearEmcabezados()
   LvUsuarios.ColumnHeaders.Add , , "Usuarios", LvUsuarios.Width - 250
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modificado Then
        Dim Rta As Integer
       Rta = MsgBox("Ha efectuado cambio �Desea Guardarlos?", vbYesNoCancel)
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

Private Sub Lvusuarios_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Bold Then
        Item.Checked = True
        MsgBox "El usuario es responsable del Centro de Costo", vbInformation
        Exit Sub
    End If
    
    If Item.ForeColor = &H808080 Then
        Item.Checked = True
        MsgBox "El usuario pertenece a otro Centro de Costo", vbInformation
    Else
        Modificado = True
    End If
End Sub

Private Sub CheckAuxiliar(CodCentro As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
 On Error Resume Next
    For i = 1 To LvUsuarios.ListItems.Count
        LvUsuarios.ListItems(i).Checked = False
        LvUsuarios.ListItems(i).Bold = False
        LvUsuarios.ListItems(i).ForeColor = vbBlack
    Next
    
    Sql = "SpTaAuxiliaresCentosDeCostosTraer"
    With RsCargar
        
        .Open Sql, Conec, adOpenStatic, adLockReadOnly
        While Not .EOF
            LvUsuarios.ListItems(CStr(!A_Usuario)).Checked = True
            If !A_CentroDeCosto <> CodCentro Then
               LvUsuarios.ListItems(CStr(!A_Usuario)).ForeColor = &H808080
            End If
            
            If TxtUsuario.Text = !A_Usuario Then
               LvUsuarios.ListItems(CStr(!A_Usuario)).Bold = True
            End If
            
            .MoveNext
        Wend
        .Close
    End With
    Set RsCargar = Nothing
    
End Sub

