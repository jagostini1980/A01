VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_1900 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas no Utilizadas En Financiero"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameServ 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   60
      TabIndex        =   7
      Top             =   5265
      Width           =   6495
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   2
         Top             =   255
         Width           =   645
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1937
         TabIndex        =   3
         Top             =   630
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3377
         TabIndex        =   4
         Top             =   630
         Width           =   1300
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   255
         Width           =   3570
         _ExtentX        =   6297
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
      Begin VB.Label LbCta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta Contable:"
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
         Left            =   75
         TabIndex        =   9
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label LbCodCuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod."
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
         Left            =   5265
         TabIndex        =   8
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1972
      TabIndex        =   5
      Top             =   6435
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3412
      TabIndex        =   6
      Top             =   6435
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   5205
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9181
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
Attribute VB_Name = "A01_1900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private Type TipoCuenta
    Descripcion As String
    Codigo As String
    Estado As String
End Type

Private VecCuentasNoUtilizadas() As TipoCuenta

Private Sub CmbCuentas_Click()
    If CmbCuentas.ListIndex > 0 Then
       TxtCodCuenta.Text = VecCuentasContables(CmbCuentas.ListIndex).Codigo
    End If
End Sub

Private Sub CmdAgregar_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = LvListado.SelectedItem.Index
      ReDim Preserve VecCuentasNoUtilizadas(UBound(VecCuentasNoUtilizadas) + 1)
      Modificado = True
     'lo pone en el LV
      VecCuentasNoUtilizadas(i).Descripcion = CmbCuentas.Text
      VecCuentasNoUtilizadas(i).Codigo = VecCuentasContables(CmbCuentas.ListIndex).Codigo
      VecCuentasNoUtilizadas(i).Estado = "A"
      
      LvListado.ListItems(i).Text = CmbCuentas.Text
      LvListado.ListItems(i).Checked = True
      LvListado.ListItems.Add
      LvListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LvListado.SelectedItem)
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Function ValidarCarga() As Boolean
Dim i As Integer
        ValidarCarga = True
        If CmbCuentas.ListIndex = 0 Then
            ValidarCarga = False
            MsgBox "Debe Seleccionar una Cuenta", vbInformation
            CmbCuentas.SetFocus
            Exit Function
        End If
        
        For i = 1 To UBound(VecCuentasNoUtilizadas)
          If LvListado.SelectedItem.Index <> i Then
            If VecCuentasNoUtilizadas(i).Codigo = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
                ValidarCarga = False
                MsgBox "La Cuenta ya fue cargarda", vbInformation
                CmbCuentas.SetFocus
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
    Rta = MsgBox("¿Desea Guardar los datos?", vbYesNo)
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
        With VecCuentasNoUtilizadas(i)
            If Not LvListado.ListItems(i).Checked Then
                .Estado = "B"
            End If
            
            Select Case .Estado
            Case "A"
                Sql = "SpOcCuentasNoUtilizadasEnFinancieroAgregar @C_CuentaContable='" & .Codigo & "'"
                Conec.Execute Sql
            Case "B"
                Sql = "SpOcCuentasNoUtilizadasEnFinancieroBorrar @C_CuentaContable='" & .Codigo & "'"
                Conec.Execute Sql
            End Select
        End With
    Next
    Conec.CommitTrans
    Call CargarLV

    MsgBox "Las Cuentas se grabaron correctamente", vbInformation, "Grabado"
ErrorGrabar:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdEliminar_Click()
    LvListado.SelectedItem.Checked = False
    VecCuentasNoUtilizadas(LvListado.SelectedItem.Index).Estado = "B"
    Modificado = True
End Sub

Private Sub CargarLV()
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    Sql = "SpOcCuentasNoUtilizadasEnFinanciero"
    With RsCargar
        .Open Sql, Conec
        ReDim VecCuentasNoUtilizadas(.RecordCount)
        i = 1
        LvListado.ListItems.Clear
        While Not .EOF
            VecCuentasNoUtilizadas(i).Codigo = !C_CuentaContable
            VecCuentasNoUtilizadas(i).Descripcion = BuscarDescCta(!C_CuentaContable)
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = VecCuentasNoUtilizadas(i).Descripcion
            LvListado.ListItems(i).Checked = True
            
            i = i + 1
            .MoveNext
        Wend
        LvListado.ListItems.Add
        .Close
    End With
    LvListado.ListItems(LvListado.ListItems.Count).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    CmdConfirnar.Enabled = True
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Errores
'NO SE TOCA
   If Item.Index < LvListado.ListItems.Count Then
       Call CargarEnModificar(Item.Index)
       CmdEliminar.Enabled = True
       CmdAgregar.Enabled = False
    Else
        CmbCuentas.Enabled = True
        TxtCodCuenta.Enabled = True
        CmbCuentas.ListIndex = 0
        TxtCodCuenta.Text = ""
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarEnModificar(i As Integer)
    With VecCuentasNoUtilizadas(i)
        Call UbicarCuentaContable(.Codigo, CmbCuentas)
        CmbCuentas.Enabled = False
        TxtCodCuenta.Enabled = False
    End With
End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarCmbCuentasContables(CmbCuentas)
    Call CargarLV
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
End Sub

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call UbicarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Cuenta Contable", LvListado.Width - 250
    LvListado.ListItems.Add
End Sub

Public Sub UbicarCuentaContable(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
   
    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = C_Codigo Then
           Cmb.ListIndex = i
           Exit For
        End If
    Next
    
End Sub

