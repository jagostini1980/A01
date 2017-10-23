VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_4200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobación de presupuestos"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3169
      TabIndex        =   15
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   45
      TabIndex        =   9
      Top             =   7020
      Width           =   10230
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar"
         Height          =   330
         Left            =   7920
         TabIndex        =   14
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   315
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   13
         Top             =   405
         Width           =   6495
      End
      Begin VB.TextBox TxtMonto 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   405
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones Aprov."
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
         Left            =   1350
         TabIndex        =   12
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         TabIndex        =   10
         Top             =   180
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView LvRenglones 
      Height          =   3615
      Left            =   45
      TabIndex        =   3
      Top             =   3420
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
      Left            =   5892
      TabIndex        =   5
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   10230
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   330
         Left            =   3105
         TabIndex        =   1
         Top             =   180
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   49610755
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
         TabIndex        =   7
         Top             =   225
         Width           =   750
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2760
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   4868
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confir&mar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4484
      TabIndex        =   4
      Top             =   7920
      Width           =   1300
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
      Left            =   8685
      TabIndex        =   8
      Top             =   8010
      Width           =   1545
   End
End
Attribute VB_Name = "A01_4200"
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
    P_CentroDeCostoEmisor As String
    P_Observaciones As String
    P_ObservacionesPresupuesto As String
    Estado As String
End Type

Private VecCabecera() As TipoOrdenDeCompra

Private VecRenglones() As TipoPresupuesto
Private Modificado As Boolean
Private LvIndex As Integer

Private Sub CalFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call CmdTraer_Click
    End If
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
Dim Rta As Integer

    Rta = MsgBox("¿Desea confirmar la Aprobación del Presupuesto?", vbYesNo)
    If Rta = vbYes Then
        Call AprobarPresupuesto
    End If
End Sub

Private Sub AprobarPresupuesto()
  Dim Sql As String
  Dim i As Integer
  Dim CentroDeCostoEmisor As String
On Error GoTo ErrorInsert

    CentroDeCostoEmisor = VecRenglones(1).P_CentroDeCostoEmisor
    Conec.BeginTrans
    For i = 1 To UBound(VecDistribucionPresupuesto)
      With VecDistribucionPresupuesto(i)
        
        Sql = "SpOcPresupuestosDistrubucionActualizar @P_NumeroPresupuesto =" & .P_NumeroPresupuesto & _
                " , @P_CentroDeCostosEmisor ='" & CentroDeCostoEmisor & _
                "', @P_SubCentroDeCosto ='" & .P_SubCentroDeCosto & _
                "', @P_CuentaContable ='" & .P_CuentaContable & _
                "', @P_Periodo ='" & Format(CalFecha.Value, "MM/yyyy") & _
                "', @P_Importe =" & Replace(.P_Importe, ",", ".")
      End With
        Conec.Execute Sql
    Next
    Conec.CommitTrans
    Conec.BeginTrans
    For i = 1 To UBound(VecRenglones)
      With VecRenglones(i)
        If .Estado = "M" Then
            
             Sql = "SpOCPresupuestosRenglonesModificar @P_NumeroPresupuesto = " & .P_NumeroPresupuesto & _
                                                   " , @P_CuentaContable = '" & .P_CuentaContable & _
                                                   "', @P_CentroDeCostoEmisor = '" & .P_CentroDeCostoEmisor & _
                                                   "', @Monto = " & .Monto & _
                                                   " , @P_Observaciones = '" & .P_Observaciones & _
                                                   "', @P_ObservacionesPresupuesto = '" & .P_ObservacionesPresupuesto & "'"
            If .P_FechaAprobacion <> "" Then
                Sql = Sql & ", @@P_FechaAprobacion= " & FechaSQL(Mid(.P_FechaAprobacion, 1, 10), "SQL")
            End If
             Conec.Execute Sql
        End If
        

        If LvRenglones.ListItems(i).Checked And .P_FechaAprobacion = "" Then
            Sql = "SpOCPresupuestosRenglonesAprobar @P_NumeroPresupuesto = " & .P_NumeroPresupuesto & _
                                                 ", @P_CuentaContable = '" & .P_CuentaContable & _
                                                "', @P_CentroDeCostoEmisor = '" & .P_CentroDeCostoEmisor & "'"
            Conec.Execute Sql
        End If
        
        If Not LvRenglones.ListItems(i).Checked And .P_FechaAprobacion <> "" Then
             Sql = "SpOCPresupuestosRenglonesDesaprobar @P_NumeroPresupuesto = " & .P_NumeroPresupuesto & _
                                                     ", @P_CuentaContable = '" & .P_CuentaContable & _
                                                    "', @P_CentroDeCostoEmisor = '" & .P_CentroDeCostoEmisor & "'"
             Conec.Execute Sql
        End If
                    
      End With
    Next
    
    Modificado = False
    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       MsgBox "La Aprobación Del Presupuesto se realizó correctamente"
    Else
       Conec.RollbackTrans
       Call TratarError(Err.Number, Err.Description)
    End If
End Sub

Private Sub CmdModificar_Click()
    Dim i As Integer
    Dim Sql As String
    Dim RsValidar As New ADODB.Recordset
    
    If LvRenglones.SelectedItem.Checked Then
       Sql = "SpOCPresupuestosRenglonesValidarMonto @CuentaContable = '" & VecRenglones(LvRenglones.SelectedItem.Index).P_CuentaContable & _
                                                "', @CentroEmisor = '" & VecRenglones(LvRenglones.SelectedItem.Index).P_CentroDeCostoEmisor & _
                                                "', @Periodo = " & Format(CalFecha.Value, "'MM/yyyy'")
       RsValidar.Open Sql, Conec
       If RsValidar!MontoSinUso < VecRenglones(LvRenglones.SelectedItem.Index).Monto And _
          VecRenglones(LvRenglones.SelectedItem.Index).P_FechaAprobacion <> "" Then
        'Si RsValidar!CantNoAsignada < VecRenglones(Item.Index).P_Cantidad
        'el renglon no se puede anular
            MsgBox "El Renglón no puede ser Modificado por haber Órdenes creadas para esao Cuenta Contable", vbInformation, "Desaprobar"
            Exit Sub
       End If
    End If
    If ValN(TxtMonto) > 0 Then
        If VecRenglones(LvRenglones.SelectedItem.Index).P_CuentaContable = "5121" Then
            'para distribuir mobilidades
            A01_4110.JerarquiaCentro = BuscarJerarquiaCentro(VecRenglones(LvRenglones.SelectedItem.Index).P_CentroDeCostoEmisor)
            A01_4110.TxtNroPresupuesto = VecRenglones(LvRenglones.SelectedItem.Index).P_NumeroPresupuesto
            A01_4110.TxtCentroDeCostoEmisor = BuscarDescCentroEmisor(VecRenglones(LvRenglones.SelectedItem.Index).P_CentroDeCostoEmisor)
            A01_4110.TxtTotalPres = ValN(TxtMonto) 'VecRenglones(LvRenglones.SelectedItem.Index).Monto
            A01_4110.CalPeriodo = CalFecha
            A01_4110.TxtCuentaContable = BuscarDescCta(VecRenglones(LvRenglones.SelectedItem.Index).P_CuentaContable)
            A01_4110.Show vbModal
            If Not A01_4110.Ok Then
                 Exit Sub
            End If
        End If
    End If
    If ValN(TxtMonto) > 0 Then
   
        VecRenglones(LvRenglones.SelectedItem.Index).Monto = ValN(TxtMonto.Text)
        VecRenglones(LvRenglones.SelectedItem.Index).P_Observaciones = TxtObservaciones.Text
        VecRenglones(LvRenglones.SelectedItem.Index).Estado = "M"
        
        LvRenglones.SelectedItem.SubItems(2) = Format(Val(TxtMonto.Text), "#.00")
        LvRenglones.SelectedItem.SubItems(4) = TxtObservaciones.Text
        Modificado = True
        
        Call CalcularTotal
    Else
        MsgBox "El monto debe ser mayor a 0", vbInformation
    End If
End Sub

Private Sub CmdTraer_Click()
    Call CargarLV(CalFecha.Value)
End Sub

Private Sub CargarLV(Periodo As Date)
    Dim Sql As String
    Dim i As Integer
    LvIndex = 0
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    Sql = "SpOCPresupuestosCabeceraTraerPeriodo @Periodo ='" & Format(CalFecha.Value, "MM/yyyy") & "'"
    With RsCargar
        .Open Sql, Conec
        i = 1
        LvListado.Sorted = False
        LvListado.ListItems.Clear
        While Not .EOF
            LvListado.ListItems.Add
            LvListado.ListItems(i).SubItems(1) = Format(!P_NumeroPresupuesto, "00000000")
            LvListado.ListItems(i).SubItems(2) = !P_FechaEmision
            LvListado.ListItems(i).SubItems(3) = Convertir(!C_Descripcion)
            LvListado.ListItems(i).SubItems(4) = !P_CentroDeCostoEmisor
            LvListado.ListItems(i).SubItems(5) = !P_Responsable
            LvListado.ListItems(i).SubItems(6) = !P_Observaciones
            LvListado.ListItems(i).Checked = !RengDesaprobados = 0
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
    LvListado.Sorted = True
    
    LvRenglones.ListItems.Clear
    If LvListado.ListItems.Count > 0 Then
        CmdConfirmar.Enabled = True
        LvListado.ListItems(1).Selected = True
        Call LvListado_ItemClick(LvListado.SelectedItem)
    Else
        CmdConfirmar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call CrearEncabezados
    Modificado = False
    
    CalFecha.Value = Date
    
    LvIndex = 0
    ReDim VecCabecera(0)
    ReDim VecRenglones(0)
    ReDim VecDistribucionPresupuesto(0)
End Sub

Private Sub CrearEncabezados()
    
    LvListado.ColumnHeaders.Add , , "Aprobado", 900
    LvListado.ColumnHeaders.Add , , "Presupuesto Nº", 1300, 1
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Centro de Costo Emisor", (LvListado.Width - 3550) / 3
    LvListado.ColumnHeaders.Add , , "CentroEmisor", 0
    LvListado.ColumnHeaders.Add , , "Responsable", (LvListado.Width - 3750) / 3
    LvListado.ColumnHeaders.Add , , "Observaciones", (LvListado.Width - 3600) / 3
    
    LvRenglones.ColumnHeaders.Add , , "Aprobado", 900
    LvRenglones.ColumnHeaders.Add , , "Cuenta Contable", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Monto", 1000, 1
    LvRenglones.ColumnHeaders.Add , , "Observaciones Pres.", (LvRenglones.Width - 2200) / 3
    LvRenglones.ColumnHeaders.Add , , "Observaciones Aprob.", (LvRenglones.Width - 2200) / 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Rta As Integer
    If Modificado Then
       Rta = MsgBox("Ha efectuedo cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            Call AprobarPresupuesto
         End If
       End If
    End If
End Sub



Private Sub LvListado_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim Sql As String
    Dim RsValidar As ADODB.Recordset
    Set RsValidar = New ADODB.Recordset

    Modificado = True
    
    If Item.Index <> LvListado.SelectedItem.Index Then
        Item.Selected = True
        Call LvListado_ItemClick(Item)
    End If
     For i = 1 To LvRenglones.ListItems.Count
        Sql = "SpOCPresupuestosRenglonesValidarMonto @CuentaContable = '" & VecRenglones(i).P_CuentaContable & _
                                                 "', @CentroEmisor = '" & VecRenglones(i).P_CentroDeCostoEmisor & _
                                                  "', @Periodo = " & Format(CalFecha.Value, "'MM/yyyy'")
         
         RsValidar.Open Sql, Conec
         LvRenglones.ListItems(i).Checked = Item.Checked
         
       If RsValidar!MontoSinUso < VecRenglones(i).Monto And _
          VecRenglones(i).P_FechaAprobacion <> "" Then
            LvRenglones.ListItems(i).Checked = True
            Modificado = True
       End If
       
         RsValidar.Close
     Next
     
     For i = 1 To LvRenglones.ListItems.Count
            Item.Checked = True
        If Not LvRenglones.ListItems(i).Checked Then
            Item.Checked = False
            Exit For
        End If
     Next
     
     Call CalcularTotal

End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
Dim Rta As Integer
Dim i As Integer

    If LvIndex = Item.Index Then
        Exit Sub
    End If
    
    If Modificado Then
       Rta = MsgBox("Ha efectuedo cambio ¿Desea Guardarlos?", vbYesNoCancel)
        'If Rta = vbCancel Or Rta = vbNo Then
        '    For i = 1 To LvRenglones.ListItems.Count
        '     If Not LvRenglones.ListItems(i).Checked Then
        '         LvListado.ListItems(LvIndex).Checked = False
        '         Exit For
        '     End If
        '    Next
        'End If
       If Rta = vbCancel Then
          LvListado.ListItems(LvIndex).Selected = True
          Exit Sub
       Else
         If Rta = vbYes Then
            Call AprobarPresupuesto
         End If
       End If

    End If
    Call CargarLvRenglones(Val(Item.SubItems(1)))
    CmdImprimir.Enabled = True
    LvIndex = Item.Index
    Modificado = False
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub CargarLvRenglones(NroPresupuesto As Integer)
  Dim i As Integer
  Dim Sql As String
  Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LvRenglones.ListItems.Clear
    Sql = "SpOCPresupuestosRenglonesTraer2 @NroPresupuesto = " & NroPresupuesto & _
                                        ", @P_CentroDeCostoEmisor = '" & LvListado.SelectedItem.SubItems(4) & "'"
    
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
            VecRenglones(i).P_Observaciones = VerificarNulo(!P_Observaciones)
            VecRenglones(i).P_ObservacionesPresupuesto = VerificarNulo(!P_ObservacionesPresupuesto)
            i = i + 1
            .MoveNext
        Wend
        .Close
         Sql = "SpOcPresupuestosDistrubucionTraer @P_NumeroPresupuesto =" & NroPresupuesto & _
                                               ", @P_CentroDeCostosEmisor ='" & LvListado.SelectedItem.SubItems(4) & "'"
        .Open Sql, Conec
        ReDim VecDistribucionPresupuesto(.RecordCount)
        For i = 1 To .RecordCount
           VecDistribucionPresupuesto(i).P_CentroDeCostosEmisor = !P_CentroDeCostosEmisor
           VecDistribucionPresupuesto(i).P_CuentaContable = !P_CuentaContable
           VecDistribucionPresupuesto(i).P_Importe = !P_Importe
           VecDistribucionPresupuesto(i).P_NumeroPresupuesto = !P_NumeroPresupuesto
           VecDistribucionPresupuesto(i).P_Periodo = !P_Periodo
           VecDistribucionPresupuesto(i).P_SubCentroDeCosto = !P_SubCentroDeCosto
           .MoveNext
        Next
     End With
     
     TxtMonto.Text = ""
     TxtObservaciones.Text = ""
     
     Call CalcularTotal
Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub LvRenglones_DblClick()
    If LvRenglones.SelectedItem.Index = 0 Then
        Exit Sub
    End If
    If VecRenglones(LvRenglones.SelectedItem.Index).P_CuentaContable = "5121" Then
        A01_4110.JerarquiaCentro = BuscarJerarquiaCentro(VecRenglones(LvRenglones.SelectedItem.Index).P_CentroDeCostoEmisor)
        A01_4110.TxtNroPresupuesto = VecRenglones(LvRenglones.SelectedItem.Index).P_NumeroPresupuesto
        A01_4110.TxtCentroDeCostoEmisor = BuscarDescCentroEmisor(VecRenglones(LvRenglones.SelectedItem.Index).P_CentroDeCostoEmisor)
        A01_4110.TxtTotalPres = VecRenglones(LvRenglones.SelectedItem.Index).Monto
        A01_4110.CalPeriodo = CalFecha
        A01_4110.FrameAsig.Enabled = False
        A01_4110.CmpAceptar.Enabled = False
        A01_4110.TxtCuentaContable = BuscarDescCta(VecRenglones(LvRenglones.SelectedItem.Index).P_CuentaContable)
        A01_4110.Show vbModal
    End If
End Sub

Private Sub LvRenglones_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim Sql As String
    Dim RsValidar As New ADODB.Recordset
    
  If Not Item.Checked Then
    Sql = "SpOCPresupuestosRenglonesValidarMonto @CuentaContable = '" & VecRenglones(Item.Index).P_CuentaContable & _
                                             "', @CentroEmisor = '" & VecRenglones(Item.Index).P_CentroDeCostoEmisor & _
                                             "', @Periodo = " & Format(CalFecha.Value, "'MM/yyyy'")
    RsValidar.Open Sql, Conec
    If RsValidar!MontoSinUso < VecRenglones(Item.Index).Monto And _
       VecRenglones(Item.Index).P_FechaAprobacion <> "" Then
     'Si RsValidar!CantNoAsignada < VecRenglones(Item.Index).P_Cantidad
     'el renglon no se puede anular
         MsgBox "El Renglón no puede ser Desaprobado por haber Órdenes creadas para esao Cuenta Contable", vbInformation, "Desaprobar"
         Item.Checked = True
         Exit Sub
    End If
 End If

    Modificado = True

    For i = 1 To LvRenglones.ListItems.Count
        If Not LvRenglones.ListItems(i).Checked Then
                Exit For
        End If
    Next
    If i - 1 = LvRenglones.ListItems.Count Then
        LvListado.ListItems(LvListado.SelectedItem.Index).Checked = True
    Else
        LvListado.ListItems(LvListado.SelectedItem.Index).Checked = False
    End If
    Call CalcularTotal
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

Private Sub LvRenglones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TxtMonto.Text = Replace(VecRenglones(Item.Index).Monto, ",", ".")
    TxtObservaciones.Text = VecRenglones(Item.Index).P_Observaciones
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    Call TxtNumerico(TxtMonto, KeyAscii)
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDePresupuesto
    RepPresupuesto.Show vbModal
End Sub

Private Sub ConfImpresionDePresupuesto()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 150
    RsListado.Fields.Append "Monto", adDouble
    RsListado.Fields.Append "Observaciones", adVarChar, 150
    RsListado.Open
    i = 1
    While i <= LvRenglones.ListItems.Count
        
      With LvRenglones.ListItems(i)
         
            RsListado.AddNew
            RsListado!CuentaContable = .SubItems(1)
            RsListado!Monto = Val(Replace(.SubItems(2), ",", "."))
            RsListado!Observaciones = .SubItems(3)
      End With
        i = i + 1
    Wend
    
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    With LvListado.SelectedItem
        RepPresupuesto.TxtFecha = .SubItems(2)
        RepPresupuesto.TxtNroPresupuesto.Text = Trim(.SubItems(1))
        RepPresupuesto.TxtCentroEmisor.Text = .SubItems(3)
        RepPresupuesto.TxtResp.Text = .SubItems(5)
        RepPresupuesto.TxtObservaciones = .SubItems(6)
        'RepPresupuesto.TxtAnulada.Visible = LBAnulada.Visible
        'RepPresupuesto.TxtAnulada.Text = LBAnulada.Caption
        RepPresupuesto.TxtPeriodo.Text = Format(CalFecha.Value, "MMMM/yyyy")
    End With
    RepPresupuesto.DataControl1.Recordset = RsListado

End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

