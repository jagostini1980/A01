VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_1410 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Períodos Cerrados"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameServ 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cierre de Períodos"
      Height          =   1185
      Left            =   90
      TabIndex        =   12
      Top             =   4905
      Width           =   4515
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   315
         Left            =   3015
         TabIndex        =   4
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   39071
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   315
         Left            =   990
         TabIndex        =   3
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   23592963
         UpDown          =   -1  'True
         CurrentDate     =   39071
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1575
         TabIndex        =   6
         Top             =   675
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3015
         TabIndex        =   7
         Top             =   675
         Width           =   1300
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   135
         TabIndex        =   5
         Top             =   675
         Width           =   1300
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
         Left            =   165
         TabIndex        =   14
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha:"
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
         Left            =   2340
         TabIndex        =   13
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Confirmar"
      Height          =   350
      Left            =   1935
      TabIndex        =   8
      Top             =   6165
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3375
      TabIndex        =   9
      Top             =   6165
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Año"
      Height          =   600
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   4470
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   1665
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalAño 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   23592963
         UpDown          =   -1  'True
         CurrentDate     =   39071
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Año:"
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
         Left            =   225
         TabIndex        =   11
         Top             =   240
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   4125
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
Attribute VB_Name = "A01_1410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Modificado As Boolean
Private AñoPeriodo As String

Private Sub CmdAgregar_Click()
On Error GoTo Errores
Dim i As Integer
   If ValidarCarga Then
      i = LvListado.SelectedItem.Index
       
      Modificado = True
   'lo pone en el LV
      LvListado.ListItems(i).Text = Format(CalPeriodo.Value, "MM/yyyy")
      LvListado.ListItems(i).SubItems(1) = CalFecha.Value
      LvListado.ListItems.Add
      Call LvListado_ItemClick(LvListado.SelectedItem)
   End If
Errores:
   Call ManipularError(Err.Number, Err.Description)

End Sub

Private Function ValidarCarga() As Boolean
Dim i As Integer
        ValidarCarga = True
        For i = 1 To LvListado.ListItems.Count
          If LvListado.SelectedItem.Index <> i Then
            If LvListado.ListItems(i).Text = Format(CalPeriodo.Value, "MM/yyyy") Then
                ValidarCarga = False
                MsgBox "El Perído ya está Cerrado", vbInformation, "Periodo"
                CalPeriodo.SetFocus
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
    Rta = MsgBox("¿Desea Cerrar los Períodos?", vbYesNo, "Cerrar Períodos")
    If Rta = vbYes Then
        Call Grabar
    End If
End Sub

Private Sub Grabar()
    Dim Sql As String
    Dim i As Integer
    Conec.BeginTrans
        Sql = "SpOCCierrePeriodoBorrar @Año =" & AñoPeriodo
        Conec.Execute Sql
    For i = 1 To LvListado.ListItems.Count - 1
        Sql = "SpOCCierrePeriodoAgregar @C_Periodo ='" & LvListado.ListItems(i).Text & _
                                 "', @C_Fecha =" & FechaSQL(LvListado.ListItems(i).SubItems(1), "SQL")
        Conec.Execute Sql

    Next
    Conec.CommitTrans
    MsgBox "Los períodos se grabaron correctamente", vbInformation, "Grabado"
End Sub

Private Sub CmdEliminar_Click()
    LvListado.ListItems.Remove (LvListado.SelectedItem.Index)
    Modificado = True
End Sub

Private Sub CmdModif_Click()
On Error GoTo Errores
Dim i As Integer
  If ValidarCarga Then
    i = LvListado.SelectedItem.Index
       
    Modificado = True
   'lo pone en el LV
         LvListado.ListItems(i).Text = Format(CalPeriodo.Value, "MM/yyyy")
         LvListado.ListItems(i).SubItems(1) = CalFecha.Value
  End If
Errores:
  Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub CmdTraer_Click()
    Call CargarLV(CStr(Year(CalAño.Value)))
End Sub

Private Sub CargarLV(Año As String)
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Sql = "SpOCCierrePeriodoTraer @Año = '" & Año & "'"
    With RsCargar
        .Open Sql, Conec
        ReDim VecPeriodoCierre(.RecordCount)
        i = 1
        LvListado.ListItems.Clear
        While Not .EOF
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = !C_Periodo
            LvListado.ListItems(i).SubItems(1) = !C_Fecha
            LvListado.ListItems(i).Checked = True
            
            i = i + 1
            .MoveNext
        Wend
        LvListado.ListItems.Add
        .Close
    End With
    If CalPeriodo.MaxDate > "01/01/" & Año Then
        CalPeriodo.MinDate = "01/01/" & Año
        CalPeriodo.MaxDate = "31/12/" & Año
    Else
        CalPeriodo.MaxDate = "31/12/" & Año
        CalPeriodo.MinDate = "01/01/" & Año
    End If
    Call LvListado_ItemClick(LvListado.ListItems(1))
    AñoPeriodo = Año
    CmdConfirnar.Visible = True
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
        CalPeriodo.Value = IIf(CalPeriodo.MaxDate < Date, CalPeriodo.MaxDate, IIf(CalPeriodo.MinDate > Date, CalPeriodo.MinDate, Date))
        CalFecha.Value = Date
        CmdAgregar.Enabled = True
        CmdModif.Enabled = False
        CmdEliminar.Enabled = False
   End If
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarEnModificar(i As Integer)
   CalPeriodo.Value = "01/" & LvListado.ListItems(i).Text
   CalFecha.Value = LvListado.ListItems(i).SubItems(1)
End Sub

Private Sub Form_Load()
    CalAño.Value = Date
    Call CrearEncabezado
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Período Cerrado", (LvListado.Width - 300) / 2
    LvListado.ColumnHeaders.Add , , "Fecha", (LvListado.Width - 300) / 2
    'LvListado.ListItems.Add
End Sub
