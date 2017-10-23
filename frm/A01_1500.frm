VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_1500 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificación financiera de Cuentas"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCopiar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar Período"
      Height          =   600
      Left            =   3420
      TabIndex        =   17
      Top             =   45
      Width           =   3300
      Begin VB.CommandButton CmdCopiar 
         Caption         =   "Copiar"
         Height          =   315
         Left            =   2115
         TabIndex        =   18
         Top             =   200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalCopiar 
         Height          =   315
         Left            =   915
         TabIndex        =   19
         Top             =   200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   53739523
         UpDown          =   -1  'True
         CurrentDate     =   39071
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período:"
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
         TabIndex        =   20
         Top             =   260
         Width           =   750
      End
   End
   Begin VB.Frame FrameServ 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   90
      TabIndex        =   12
      Top             =   4905
      Width           =   6630
      Begin VB.TextBox TxtCantMeses 
         Height          =   315
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   5
         Top             =   630
         Width           =   645
      End
      Begin VB.TextBox TxtCodCuenta 
         Height          =   315
         Left            =   5895
         MaxLength       =   4
         TabIndex        =   4
         Top             =   255
         Width           =   645
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2642
         TabIndex        =   7
         Top             =   990
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4082
         TabIndex        =   8
         Top             =   990
         Width           =   1300
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1202
         TabIndex        =   6
         Top             =   990
         Width           =   1300
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   255
         Width           =   3525
         _ExtentX        =   6218
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad de Meses:"
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
         TabIndex        =   16
         Top             =   690
         Width           =   1695
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
         Left            =   210
         TabIndex        =   15
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
         Left            =   5400
         TabIndex        =   14
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2047
      TabIndex        =   9
      Top             =   6435
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3487
      TabIndex        =   10
      Top             =   6435
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
      Height          =   600
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   3300
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   2115
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   53739523
         UpDown          =   -1  'True
         CurrentDate     =   39071
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período:"
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
         TabIndex        =   13
         Top             =   240
         Width           =   750
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   4125
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   7276
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
Attribute VB_Name = "A01_1500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean

Private Type TipoClasificacionCta
    C_Periodo As String
    C_Cuenta As String
    C_CantidadMeses As Integer
    Estado As String
End Type

Private PeriodoActual As String
Private VecClasificacionCta() As TipoClasificacionCta

Private Sub CalPeriodo_Validate(Cancel As Boolean)
    If PeriodoActual <> Format(CalPeriodo, "MM/yyyy") Then
        ReDim VecClasificacionCta(0)
        LvListado.ListItems.Clear
        LvListado.ListItems.Add
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    End If
End Sub

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
      ReDim Preserve VecClasificacionCta(UBound(VecClasificacionCta) + 1)
      Modificado = True
     'lo pone en el LV
      VecClasificacionCta(i).C_CantidadMeses = Val(TxtCantMeses)
      VecClasificacionCta(i).C_Cuenta = VecCuentasContables(CmbCuentas.ListIndex).Codigo
      VecClasificacionCta(i).C_Periodo = Format(CalPeriodo, "MM/yyyy")
      VecClasificacionCta(i).Estado = "A"
      
      LvListado.ListItems(i).Text = CmbCuentas.Text
      LvListado.ListItems(i).SubItems(1) = TxtCantMeses
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

        If Val(TxtCantMeses) = 0 Then
            ValidarCarga = False
            MsgBox "La Cantidad de Meses debe ser mayor a 0", vbInformation
            TxtCantMeses.SetFocus
            Exit Function
        End If
        
        For i = 1 To UBound(VecClasificacionCta)
          If LvListado.SelectedItem.Index <> i Then
            If VecClasificacionCta(i).C_Cuenta = VecCuentasContables(CmbCuentas.ListIndex).Codigo Then
                ValidarCarga = False
                MsgBox "La Cuenta ya fue cargarda para este período", vbInformation
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
    Rta = MsgBox("¿Desea Cerrar los Períodos?", vbYesNo)
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
        With VecClasificacionCta(i)
            If Not LvListado.ListItems(i).Checked Then
                .Estado = "B"
            End If
            
            Select Case .Estado
            Case "A"
                Sql = "SpTaClasificacionFinancieraCuentasContablesAgregar @C_Periodo ='" & .C_Periodo & _
                                                                      "', @C_Cuenta ='" & .C_Cuenta & _
                                                                      "', @C_CantidadMeses =" & .C_CantidadMeses
                Conec.Execute Sql
            Case "M"
                Sql = "SpTaClasificacionFinancieraCuentasContablesModificar @C_Periodo ='" & .C_Periodo & _
                                                                        "', @C_Cuenta ='" & .C_Cuenta & _
                                                                        "', @C_CantidadMeses =" & .C_CantidadMeses
                Conec.Execute Sql
            Case "B"
                Sql = "SpTaClasificacionFinancieraCuentasContablesBorrar @C_Periodo ='" & .C_Periodo & _
                                                                     "', @C_Cuenta ='" & .C_Cuenta & "'"
                Conec.Execute Sql
            End Select
        End With
    Next
    Conec.CommitTrans
    Call CargarLV(Format(CalPeriodo, "MM/yyyy"))

    MsgBox "Los períodos se grabaron correctamente", vbInformation, "Grabado"
ErrorGrabar:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdCopiar_Click()
    Dim i As Integer
    Call CargarLV(Format(CalCopiar, "MM/yyyy"))
    'emula que son todos nuevos
    For i = 1 To UBound(VecClasificacionCta)
        VecClasificacionCta(i).Estado = "A"
    Next
End Sub

Private Sub CmdEliminar_Click()
    LvListado.SelectedItem.Checked = False
    VecClasificacionCta(LvListado.SelectedItem.Index).Estado = "B"
    Modificado = True
End Sub

Private Sub CmdModif_Click()
'On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = LvListado.SelectedItem.Index
      Modificado = True
     'lo pone en el LV
      VecClasificacionCta(i).C_CantidadMeses = Val(TxtCantMeses)
      VecClasificacionCta(i).C_Cuenta = VecCuentasContables(CmbCuentas.ListIndex).Codigo
      VecClasificacionCta(i).C_Periodo = Format(CalPeriodo, "MM/yyyy")
      If VecClasificacionCta(i).Estado <> "A" Then
         VecClasificacionCta(i).Estado = "M"
      End If

      LvListado.ListItems(i).Text = CmbCuentas.Text
      LvListado.ListItems(i).SubItems(1) = TxtCantMeses
      LvListado.ListItems(i).Checked = True
      LvListado.ListItems(i + 1).Selected = True
      Call LvListado_ItemClick(LvListado.SelectedItem)
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV(Format(CalPeriodo, "MM/yyyy"))
End Sub

Private Sub CargarLV(Periodo As String)
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    Sql = "SpTaClasificacionFinancieraCuentasContablesTraer @Periodo ='" & Periodo & "'"
    With RsCargar
        .Open Sql, Conec
        ReDim VecClasificacionCta(.RecordCount)
        i = 1
        LvListado.ListItems.Clear
        While Not .EOF
            VecClasificacionCta(i).C_CantidadMeses = !C_CantidadMeses
            VecClasificacionCta(i).C_Cuenta = !C_Cuenta
            VecClasificacionCta(i).C_Periodo = Format(CalPeriodo, "MM/yyyy")
            
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = BuscarDescCta(!C_Cuenta)
            LvListado.ListItems(i).SubItems(1) = !C_CantidadMeses
            LvListado.ListItems(i).Checked = True
            
            i = i + 1
            .MoveNext
        Wend
        PeriodoActual = Format(CalPeriodo, "MM/yyyy")
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
       CmdModif.Enabled = True
       CmdEliminar.Enabled = True
       CmdAgregar.Enabled = False
    Else
        CmbCuentas.Enabled = True
        TxtCodCuenta.Enabled = True
        CmbCuentas.ListIndex = 0
        TxtCodCuenta.Text = ""
        TxtCantMeses.Text = ""
        CmdAgregar.Enabled = True
        CmdModif.Enabled = False
        CmdEliminar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarEnModificar(i As Integer)
    With VecClasificacionCta(i)
        Call UbicarCuentaContable(.C_Cuenta, CmbCuentas)
        CmbCuentas.Enabled = False
        TxtCodCuenta.Enabled = False

        TxtCantMeses = .C_CantidadMeses
    End With
End Sub

Private Sub Form_Load()
    Call CrearEncabezado
    Call CargarCmbCuentasContables(CmbCuentas)
    ReDim VecClasificacionCta(0)
    CalPeriodo.Value = Date
    CalCopiar.Value = DateAdd("M", -1, Date)
    Call CargarLV(Format(CalPeriodo, "MM/yyyy"))
    
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
End Sub

Private Sub TxtCantMeses_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
    End If
End Sub

Private Sub TxtCodCuenta_LostFocus()
    If TxtCodCuenta <> "" Then
       Call UbicarCuentaContable(TxtCodCuenta, CmbCuentas)
    End If
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Cuenta Contable", LvListado.Width - 1750
    LvListado.ColumnHeaders.Add , , "Cant. De Meses", 1500, 1
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

