VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_1B100 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
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
      Height          =   1005
      Left            =   60
      TabIndex        =   7
      Top             =   5985
      Width           =   6495
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1170
         TabIndex        =   2
         Top             =   540
         Width           =   1300
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   1
         Top             =   165
         Width           =   4740
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2565
         TabIndex        =   3
         Top             =   540
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3960
         TabIndex        =   4
         Top             =   540
         Width           =   1300
      End
      Begin VB.Label LbCta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Forma de Pago:"
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
         TabIndex        =   8
         Top             =   225
         Width           =   1350
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1927
      TabIndex        =   5
      Top             =   7065
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3367
      TabIndex        =   6
      Top             =   7065
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   5925
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10451
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
Attribute VB_Name = "A01_1B100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private VecFormasDePagoTmp() As TipoFormaDePago

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = LvListado.SelectedItem.Index
      ReDim Preserve VecFormasDePagoTmp(UBound(VecFormasDePagoTmp) + 1)
      Modificado = True
     'lo pone en el LV
      VecFormasDePagoTmp(i).F_Descripcion = TxtDescripcion.Text
      VecFormasDePagoTmp(i).Estado = "A"
      
      LvListado.ListItems(i).Text = TxtDescripcion.Text
     
      LvListado.ListItems(i).Checked = True
      LvListado.ListItems.Add
      LvListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LvListado.SelectedItem)
      Modificado = True
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Function ValidarCarga() As Boolean
Dim i As Integer

        ValidarCarga = True
        If TxtDescripcion.Text = "" Then
            ValidarCarga = False
            MsgBox "Debe Ingresar La Descripcion", vbInformation
            TxtDescripcion.SetFocus
            Exit Function
        End If
        
        For i = 1 To UBound(VecFormasDePagoTmp)
          If LvListado.SelectedItem.Index <> i Then
            If VecFormasDePagoTmp(i).F_Descripcion = TxtDescripcion Then
                ValidarCarga = False
                MsgBox "La Forma de Pago ya fue Cargarda", vbInformation
                TxtDescripcion.SetFocus
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

    For i = 1 To LvListado.ListItems.Count - 1
        With VecFormasDePagoTmp(i)
            If Not LvListado.ListItems(i).Checked Then
                .Estado = "B"
            End If
            
            Select Case .Estado
            Case "A"
                Sql = "SpTaFormasDePagoAgregar @F_Descripcion='" & .F_Descripcion & "'"
                Conec.Execute Sql
            Case "M"
                Sql = "SpTaFormasDePagoModificar @F_Codigo =" & .F_Codigo & _
                                              ", @F_Descripcion='" & .F_Descripcion & "'"
                Conec.Execute Sql
            Case "B"
                Sql = "SpTaFormasDePagoBorrar @F_Codigo =" & .F_Codigo
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
        Call CargarVecFormasDePago
        Call CargarLV
        MsgBox "Las Formas de Pago se grabaron correctamente", vbInformation, "Grabado"
    End If
End Sub

Private Sub CmdEliminar_Click()
    LvListado.SelectedItem.Checked = False
    'VecLugaresDeEntregaTmp(LvListado.SelectedItem.Index).Estado = "B"
    Modificado = True
End Sub

Private Sub CargarLV()
On Error GoTo ErrorCarga

    Dim i As Integer

    ReDim VecFormasDePagoTmp(UBound(VecFormasDePago))
 
    LvListado.ListItems.Clear
    For i = 1 To UBound(VecFormasDePagoTmp)
        VecFormasDePagoTmp(i).F_Codigo = VecFormasDePago(i).F_Codigo
        VecFormasDePagoTmp(i).F_Descripcion = VecFormasDePago(i).F_Descripcion
       
        
        LvListado.ListItems.Add
        LvListado.ListItems(i).Text = VecFormasDePagoTmp(i).F_Descripcion
       
        LvListado.ListItems(i).Checked = True
    Next
    
    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    CmdConfirnar.Enabled = True
    Modificado = False
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdModificar_Click()
Dim i As Integer
   If ValidarCarga Then
      i = LvListado.SelectedItem.Index
      Modificado = True
     'lo pone en el LV
      VecFormasDePagoTmp(i).F_Descripcion = TxtDescripcion.Text

      If VecFormasDePagoTmp(i).F_Codigo = 0 Then
         VecFormasDePagoTmp(i).Estado = "A"
      Else
         VecFormasDePagoTmp(i).Estado = "M"
      End If
      
      LvListado.ListItems(i).Text = TxtDescripcion.Text
      LvListado.ListItems(i).Checked = True
      LvListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LvListado.SelectedItem)
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
   If Item.Index < LvListado.ListItems.Count Then
       Call CargarEnModificar(Item.Index)
       CmdEliminar.Enabled = True
       CmdAgregar.Enabled = False
       CmdModificar.Enabled = True
    Else
        TxtDescripcion = ""
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
        CmdModificar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarEnModificar(i As Integer)
    With VecFormasDePagoTmp(i)
       TxtDescripcion = .F_Descripcion

    End With
End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarLV
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Forma de Pago", LvListado.Width - 250
    
    LvListado.ListItems.Add
End Sub

