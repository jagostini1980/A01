VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_2400 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de presupuestos"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3308
      TabIndex        =   8
      Top             =   7920
      Width           =   1230
   End
   Begin MSComctlLib.ListView LvRenglones 
      Height          =   4515
      Left            =   67
      TabIndex        =   3
      Top             =   3285
      Width           =   9100
      _ExtentX        =   16060
      _ExtentY        =   7964
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4703
      TabIndex        =   4
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   74
      TabIndex        =   5
      Top             =   45
      Width           =   9100
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   330
         Left            =   2295
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   225
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   23527427
         UpDown          =   -1  'True
         CurrentDate     =   38993
      End
      Begin VB.Label Label1 
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
         Left            =   270
         TabIndex        =   6
         Top             =   293
         Width           =   750
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2400
      Left            =   67
      TabIndex        =   2
      Top             =   810
      Width           =   9100
      _ExtentX        =   16060
      _ExtentY        =   4233
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label LbTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Aprobado: $"
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
      Left            =   10215
      TabIndex        =   7
      Top             =   7920
      Width           =   1545
   End
End
Attribute VB_Name = "A01_2400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoPresupuesto
    P_NumeroPresupuesto As Integer
    'P_CodigoArticulo As Integer
    P_CuentaContable As String
    'P_CentroDeCosto As Integer
    'P_Cantidad As Long
    'P_ImporteUnitario As Double
    Monto As Double
    P_FechaAprobacion As String
    P_CentroDeCostoEmisor As Integer
End Type

Private VecCabecera() As TipoOrdenDeCompra
Private VecRenglones() As TipoPresupuesto

Private Sub CalFecha_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            Call CmdTraer_Click
        End If
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDePresupuesto
    RepPresupuesto.Show
End Sub

Private Sub ConfImpresionDePresupuesto()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 150
    RsListado.Fields.Append "Observaciones", adVarChar, 150
    RsListado.Fields.Append "Monto", adDouble
    RsListado.Open
    i = 1
    While i <= LvRenglones.ListItems.Count
        
      With LvRenglones.ListItems(i)
        If .Checked Then
           RsListado.AddNew
           RsListado!CuentaContable = .SubItems(1)
           RsListado!Monto = Val(Replace(.SubItems(2), ",", "."))
           RsListado!Observaciones = .SubItems(3)
           
        End If
      End With
        i = i + 1
    Wend
    
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    With LvListado.SelectedItem
        RepPresupuesto.TxtFecha = .SubItems(1)
        RepPresupuesto.TxtNroPresupuesto.Text = Trim(.Text)
        RepPresupuesto.TxtCentroEmisor.Text = .SubItems(2)
        RepPresupuesto.TxtResp.Text = .SubItems(4)
        RepPresupuesto.TxtObservaciones = .SubItems(5)
        'RepPresupuesto.TxtAnulada.Visible = LBAnulada.Visible
        'RepPresupuesto.TxtAnulada.Text = LBAnulada.Caption
        RepPresupuesto.TxtPeriodo.Text = Format(CalFecha.Value, "MMMM/yyyy")
    End With
    RepPresupuesto.DataControl1.Recordset = RsListado
    RepPresupuesto.Zoom = -1

End Sub

Private Sub CmdTraer_Click()
    Call CargarLV(CalFecha.Value)
End Sub

Private Sub CargarLV(Periodo As Date)
On Error GoTo Error
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Sql = "SpOCPresupuestosCabeceraTraerPeriodoPorUsuario @Periodo ='" & Format(CalFecha.Value, "MM/yyyy") & _
            "', @Usuario ='" & Usuario & "'"
    With RsCargar
        .Open Sql, Conec
        i = 1
        LvListado.ListItems.Clear
        While Not .EOF
            LvListado.ListItems.Add
            LvListado.ListItems(i).Text = Format(!P_NumeroPresupuesto, "            00000000")
            LvListado.ListItems(i).SubItems(1) = !P_FechaEmision
            LvListado.ListItems(i).SubItems(2) = Convertir(!C_Descripcion)
            LvListado.ListItems(i).SubItems(3) = !P_CentroDeCostoEmisor
            LvListado.ListItems(i).SubItems(4) = !P_Responsable
            LvListado.ListItems(i).SubItems(5) = !P_Observaciones
            
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
    
    LvRenglones.ListItems.Clear
    If LvListado.ListItems.Count > 0 Then
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    End If
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    
    CalFecha.Value = Date
       
    ReDim VecCabecera(0)
    ReDim VecRenglones(0)
End Sub

Private Sub CrearEncabezados()
    
    LvListado.ColumnHeaders.Add , , "Presupuesto Nº", 1400
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Centro de Costo Emisor", (LvListado.Width - 2800) / 3
    LvListado.ColumnHeaders.Add , , "CentroEmisor", 0
    LvListado.ColumnHeaders.Add , , "Responsable", (LvListado.Width - 2800) / 3
    LvListado.ColumnHeaders.Add , , "Observaciones", (LvListado.Width - 2800) / 3
    
    LvRenglones.ColumnHeaders.Add , , "Aprobado", 900
    LvRenglones.ColumnHeaders.Add , , "Cuenta Contable", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Monto", 1000, 1
    LvRenglones.ColumnHeaders.Add , , "Observaciones Pres.", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Observaciones Aprov.", (LvRenglones.Width - 2200) / 3
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
'NO SE TOCA
    Call CargarLvRenglones(Val(Item.Text))
    CmdImprimir.Enabled = True
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarLvRenglones(NroPresupuesto As Integer)
  Dim i As Integer
  Dim Sql As String
  Dim RsCargar As ADODB.Recordset
  Set RsCargar = New ADODB.Recordset

On Error GoTo Error
    LvRenglones.ListItems.Clear
    Sql = "SpOCPresupuestosRenglonesTraer2 @NroPresupuesto = " & NroPresupuesto & _
                                        ", @P_CentroDeCostoEmisor = '" & LvListado.SelectedItem.SubItems(3) & "'"
    
      With RsCargar
        .Open Sql, Conec
        i = 1
        ReDim VecRenglones(.RecordCount)
        
        While Not .EOF
            LvRenglones.ListItems.Add
            LvRenglones.ListItems(i).SubItems(1) = BuscarDescCta(!P_CuentaContable) & " Cód. " & !P_CuentaContable
            LvRenglones.ListItems(i).SubItems(2) = Format(!Monto, "#.00")
            LvRenglones.ListItems(i).SubItems(3) = VerificarNulo(!P_ObservacionesPresupuesto)
            LvRenglones.ListItems(i).SubItems(4) = VerificarNulo(!P_Observaciones)
            
            LvRenglones.ListItems(i).Checked = Not IsNull(!P_FechaAprobacion)
            
            VecRenglones(i).P_NumeroPresupuesto = !P_NumeroPresupuesto
            VecRenglones(i).P_CuentaContable = !P_CuentaContable
            VecRenglones(i).Monto = !Monto
            VecRenglones(i).P_FechaAprobacion = VerificarNulo(!P_FechaAprobacion)
            VecRenglones(i).P_CentroDeCostoEmisor = !P_CentroDeCostoEmisor
            i = i + 1
            .MoveNext
        Wend
     End With
     Call CalcularTotal
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub LvRenglones_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = VecRenglones(Item.Index).P_FechaAprobacion <> ""
End Sub

Private Sub CalcularTotal()
   Dim Total As Double
   Dim i As Integer
   
    For i = 1 To LvRenglones.ListItems.Count
       If LvRenglones.ListItems(i).Checked Then
          Total = Total + Val(Replace(LvRenglones.ListItems(i).SubItems(2), ",", "."))
       End If
    Next
    
    LbTotal.Caption = "Total Aprobado: $" & Format(Total, "0.00")
End Sub

