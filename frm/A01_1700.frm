VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_1700 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lugares de Entrega"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameServ 
      BackColor       =   &H00E0E0E0&
      Height          =   1860
      Left            =   60
      TabIndex        =   8
      Top             =   5130
      Width           =   6495
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1215
         TabIndex        =   3
         Top             =   1440
         Width           =   1300
      End
      Begin VB.TextBox TxtLugar 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   1
         Top             =   165
         Width           =   4740
      End
      Begin VB.TextBox TxtEMail 
         Height          =   675
         Left            =   90
         MaxLength       =   500
         TabIndex        =   2
         Top             =   705
         Width           =   6315
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2610
         TabIndex        =   4
         Top             =   1440
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4005
         TabIndex        =   5
         Top             =   1440
         Width           =   1300
      End
      Begin VB.Label LbCta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lugar de Entrega"
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
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   1485
      End
      Begin VB.Label LbCodCuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-Mail"
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
         TabIndex        =   9
         Top             =   495
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1972
      TabIndex        =   6
      Top             =   7110
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3412
      TabIndex        =   7
      Top             =   7110
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   5070
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8943
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
End
Attribute VB_Name = "A01_1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private VecLugaresDeEntregaTmp() As TipoLugarDeEntrega

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = LVListado.SelectedItem.Index
      ReDim Preserve VecLugaresDeEntregaTmp(UBound(VecLugaresDeEntregaTmp) + 1)
      Modificado = True
     'lo pone en el LV
      VecLugaresDeEntregaTmp(i).L_Descripcion = TxtLugar.Text
      VecLugaresDeEntregaTmp(i).L_EMail = TxtEMail
      VecLugaresDeEntregaTmp(i).Estado = "A"
      
      LVListado.ListItems(i).Text = TxtLugar.Text
      LVListado.ListItems(i).SubItems(1) = TxtEMail.Text
      LVListado.ListItems(i).Checked = True
      LVListado.ListItems.Add
      LVListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LVListado.SelectedItem)
      Modificado = True
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Function ValidarCarga() As Boolean
Dim i As Integer

        ValidarCarga = True
        If TxtLugar.Text = "" Then
            ValidarCarga = False
            MsgBox "Debe Ingresar La Descripcion", vbInformation
            TxtLugar.SetFocus
            Exit Function
        End If
        
        For i = 1 To UBound(VecLugaresDeEntregaTmp)
          If LVListado.SelectedItem.Index <> i Then
            If VecLugaresDeEntregaTmp(i).L_Descripcion = TxtLugar Then
                ValidarCarga = False
                MsgBox "El Lugar ya fue Cargardo", vbInformation
                TxtLugar.SetFocus
                Exit Function
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
On Error GoTo ErrorGrabar
    Dim Sql As String
    Dim i As Integer
    Conec.BeginTrans

    For i = 1 To LVListado.ListItems.Count - 1
        With VecLugaresDeEntregaTmp(i)
            If Not LVListado.ListItems(i).Checked Then
                .Estado = "B"
            End If
            
            Select Case .Estado
            Case "A"
                Sql = "SpTaLugaresDeEntregaAgregar @L_Descripcion='" & .L_Descripcion & _
                                               "', @L_EMail ='" & .L_EMail & "'"
                Conec.Execute Sql
            Case "M"
                Sql = "SpTaLugaresDeEntregaModificar @L_Codigo =" & .L_Codigo & _
                                                  ", @L_Descripcion='" & .L_Descripcion & _
                                                 "', @L_EMail ='" & .L_EMail & "'"
                Conec.Execute Sql
            Case "B"
                Sql = "SpTaLugaresDeEntregaBorrar @L_Codigo =" & .L_Codigo
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
        Call CargarVecLugaresDeEntrega
        Call CargarLV
        MsgBox "Los Lugares de Entrega se grabaron correctamente", vbInformation, "Grabado"
    End If
End Sub

Private Sub CmdEliminar_Click()
    LVListado.SelectedItem.Checked = False
    'VecLugaresDeEntregaTmp(LvListado.SelectedItem.Index).Estado = "B"
    Modificado = True
End Sub

Private Sub CargarLV()
On Error GoTo ErrorCarga

    Dim i As Integer

    ReDim VecLugaresDeEntregaTmp(UBound(VecLugaresDeEntrega))
 
    LVListado.ListItems.Clear
    For i = 1 To UBound(VecLugaresDeEntrega)
        VecLugaresDeEntregaTmp(i).L_Codigo = VecLugaresDeEntrega(i).L_Codigo
        VecLugaresDeEntregaTmp(i).L_Descripcion = VecLugaresDeEntrega(i).L_Descripcion
        VecLugaresDeEntregaTmp(i).L_EMail = VecLugaresDeEntrega(i).L_EMail
        
        LVListado.ListItems.Add
        LVListado.ListItems(i).Text = VecLugaresDeEntrega(i).L_Descripcion
        LVListado.ListItems(i).SubItems(1) = VecLugaresDeEntrega(i).L_EMail
        LVListado.ListItems(i).Checked = True
    Next
    
    LVListado.ListItems.Add
    LVListado.ListItems(LVListado.ListItems.Count).Selected = True
    Call LvListado_ItemClick(LVListado.SelectedItem)
    CmdConfirnar.Enabled = True
    Modificado = False
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdModificar_Click()
Dim i As Integer
   If ValidarCarga Then
      i = LVListado.SelectedItem.Index
      Modificado = True
     'lo pone en el LV
      VecLugaresDeEntregaTmp(i).L_Descripcion = TxtLugar.Text
      VecLugaresDeEntregaTmp(i).L_EMail = TxtEMail
      If VecLugaresDeEntregaTmp(i).L_Codigo = 0 Then
         VecLugaresDeEntregaTmp(i).Estado = "A"
      Else
         VecLugaresDeEntregaTmp(i).Estado = "M"
      End If
      
      LVListado.ListItems(i).Text = TxtLugar.Text
      LVListado.ListItems(i).SubItems(1) = TxtEMail.Text
      LVListado.ListItems(i).Checked = True
      LVListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LVListado.SelectedItem)
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

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Errores
'NO SE TOCA
   If Item.Index < LVListado.ListItems.Count Then
       Call CargarEnModificar(Item.Index)
       CmdEliminar.Enabled = True
       CmdAgregar.Enabled = False
       CmdModificar.Enabled = True
    Else
        TxtLugar = ""
        TxtEMail = ""
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
        CmdModificar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarEnModificar(i As Integer)
    With VecLugaresDeEntregaTmp(i)
       TxtLugar = .L_Descripcion
       TxtEMail = .L_EMail
    End With
End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarLV
    LVListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LVListado.SelectedItem)
End Sub

Private Sub CrearEncabezado()
    LVListado.ColumnHeaders.Add , , "Lugar de Entrega", (LVListado.Width - 250) / 2
    LVListado.ColumnHeaders.Add , , "E-Mail", (LVListado.Width - 250) / 2
    LVListado.ListItems.Add
End Sub

