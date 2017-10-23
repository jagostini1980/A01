VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5800 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Evolución Presupuestado/Real"
   ClientHeight    =   8115
   ClientLeft      =   3225
   ClientTop       =   2250
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5220
      TabIndex        =   11
      Top             =   7650
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   195
      Top             =   7590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Height          =   375
      Left            =   3555
      TabIndex        =   10
      Top             =   7650
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   11790
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   10575
         TabIndex        =   2
         Top             =   225
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   330
         Left            =   1530
         TabIndex        =   0
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   114360323
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   330
         Left            =   4320
         TabIndex        =   1
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   114360323
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   7020
         TabIndex        =   8
         Top             =   210
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label LbCentroEmisor 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro Emisor:"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Desde:"
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
         Left            =   105
         TabIndex        =   7
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label LBFechaHasta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Hasta:"
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
         Left            =   2985
         TabIndex        =   6
         Top             =   285
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6885
      TabIndex        =   4
      Top             =   7650
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6825
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   12039
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
Attribute VB_Name = "A01_5800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Per1 As Integer
Public Per2 As Integer
Public Per3 As Integer

Private Sub CargarLV(FechaDesde As Date, FechaHasta As Date)
  Dim Sql As String
   Dim RsListado As New ADODB.Recordset
   Dim FechaDesdeReal As String
   Dim FechaHastaReal As String
   Dim Periodo As String
   Dim NroPeriodo As Integer
   Dim Cuenta As String
   Dim pos As Integer
   
   Dim i As Integer
   
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   MousePointer = vbHourglass

On Error GoTo ErrorTraer:
   If CalFechaDesde.Value > CalFechaHasta.Value Then
        MsgBox "Rango de fechas no válido", vbInformation
        CalFechaHasta.SetFocus
        Exit Sub
   Else
        LvListado.ColumnHeaders.Clear
        LvListado.ColumnHeaders.Add , , "Cuenta Contable", 3000
        For i = 0 To DateDiff("M", CalFechaDesde, CalFechaHasta)
            LvListado.ColumnHeaders.Add , , "Pres.", 1300, 1
            LvListado.ColumnHeaders.Add , , "Real", 1300, 1
            LvListado.ColumnHeaders.Add , , "Diferencia " & Format(DateAdd("M", i, CalFechaDesde), "MMM/yy"), 1500, 1
        Next
        LvListado.ColumnHeaders.Add , , , 0
   End If
   
   FechaDesdeReal = "01/" & Format(FechaDesde, "MM/yyyy")
   FechaHastaReal = LenMes(Month(FechaHasta), Year(FechaHasta)) & "/" & Format(FechaHasta, "MM/yyyy")
   
   Sql = "SpOcConsultaEvolucionPresupuestadoReal @PeriodoDesde ='" + Format(FechaDesde, "MM/yyyy") + _
                                             "', @PeriodoHasta ='" + Format(FechaHasta, "MM/yyyy") + _
                                             "', @CentroDeCosto='" + VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo + _
                                             "', @PeriodoDesdeReal =" + FechaSQL(FechaDesdeReal, "SQL") + _
                                             " , @PeriodoHastaReal =" + FechaSQL(FechaHastaReal, "SQL")
  
   LvListado.Sorted = False
   RsListado.Open Sql, Conec
   With RsListado
   'limpia el LV
   LvListado.ListItems.Clear
   CmdImprimir.Enabled = .RecordCount > 0
   If .RecordCount > 0 Then
    
      For i = 1 To .RecordCount
        
        If Cuenta <> !CuentaContable Then
            LvListado.ListItems.Add
            pos = pos + 1
            LvListado.ListItems(pos).Text = BuscarDescCta(!CuentaContable) & " (Cod. " & !CuentaContable & ")"
            Cuenta = !CuentaContable
            
        End If
        
        NroPeriodo = DateDiff("M", CalFechaDesde, !Periodo)
        LvListado.ListItems(pos).SubItems(1 + NroPeriodo * 3) = Format(VerificarNulo(!TotPres, "N"), "0.00")
        LvListado.ListItems(pos).SubItems(2 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N"), "0.00")
        LvListado.ListItems(pos).SubItems(3 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N") - VerificarNulo(!TotPres, "N"), "0.00")
        'LvListado.ListItems(i).SubItems(4) = VerificarNulo(!A_Observaciones)
        .MoveNext
      Next
    End If
   End With
    LvListado.Sorted = True
    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal

End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Cuenta Contable", 3000
    LvListado.ColumnHeaders.Add , , "Pres.", 1300, 1
    LvListado.ColumnHeaders.Add , , "Real", 1300, 1
    LvListado.ColumnHeaders.Add , , "Diferencia", 1300, 1
End Sub

Private Sub CalFechaDesde_Change()
'con esto controlo que el fecha hasta no sea menos que fecha desde
    CalFechaHasta.MinDate = CalFechaDesde.Value
End Sub

Private Sub CalFechaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    'si se preciona enter se carga el liste view
        Call CmdTraer_Click
    End If

End Sub

Private Sub CalFechaHasta_Change()
'con esto controlo que fecha desde no sea mayor que fecha hasta
    CalFechaDesde.MaxDate = CalFechaHasta.Value
End Sub

Private Sub CalFechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
   'si se preciona enter se carga el liste view
        Call CmdTraer_Click
    End If

End Sub

Private Sub CmbLineas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Beep
End Sub

Private Sub CmdExp_Click()
    Dialogo.Filename = ""
    Call ArmarExcel(Dialogo)
    If Dialogo.Filename <> "" Then
        MousePointer = vbHourglass
        Call GenerarPlanilla(Dialogo.Filename, Dialogo.FilterIndex)
        MousePointer = vbNormal
    End If

End Sub

Private Sub GenerarPlanilla(NombreArchivo As String, Filtro As Integer)
Dim ex As Excel.Application
Dim col As Integer
Dim ColorFondo As Long

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcel(ex, LvListado, Caption, 6)
        Call DatosExcel(ex, LvListado, 6)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To LvListado.ColumnHeaders.Count
            .Columns(LetraColumna(col) & ":" & LetraColumna(col)).EntireColumn.AutoFit
        Next
        '.Columns("D:D").ColumnWidth = 25
        '-----ESTO LO PONGO LUEGO DE AJUSTAR LAS COLUMNAS PORQUE SINO SALEN MAL --------
        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Fecha: " & Date
        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Hora: " & Time
        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Periodo Desde : " & Me.CalFechaDesde & " Hasta: " & CalFechaHasta
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Centro de Costo: " & CmbCentroDeCostoEmisor.Text

        ColorFondo = &HC0E0FF
        Call FormatearExcel(ex, LvListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5800.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    Per1 = 0
    Per2 = -1
    Per3 = -1
    
    If DateDiff("M", CalFechaDesde, CalFechaHasta) >= 1 Then
        A01_5810.PeriodoFin = CalFechaHasta
        A01_5810.PeriodoInicio = CalFechaDesde
        A01_5810.Show vbModal
    End If
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 100
    RsListado.Fields.Append "Pres1", adDouble
    RsListado.Fields.Append "Real1", adDouble
    RsListado.Fields.Append "Diferencia1", adDouble
    RsListado.Fields.Append "Pres2", adDouble
    RsListado.Fields.Append "Real2", adDouble
    RsListado.Fields.Append "Diferencia2", adDouble
    RsListado.Fields.Append "Pres3", adDouble
    RsListado.Fields.Append "Real3", adDouble
    RsListado.Fields.Append "Diferencia3", adDouble
    
    RsListado.Open
    i = 1
    While i <= LvListado.ListItems.Count
       RsListado.AddNew
       With LvListado.ListItems(i)
            RsListado!CuentaContable = .Text
            RsListado!Pres1 = ValN(.SubItems(1 + Per1 * 3))
            RsListado!Real1 = ValN(.SubItems(2 + Per1 * 3))
            RsListado!Diferencia1 = ValN(.SubItems(3 + Per1 * 3))
            If Per2 >= 1 Then
                RsListado!Pres2 = ValN(.SubItems(1 + Per2 * 3))
                RsListado!Real2 = ValN(.SubItems(2 + Per2 * 3))
                RsListado!Diferencia2 = ValN(.SubItems(3 + Per2 * 3))
            End If
            If Per3 >= 2 Then
                RsListado!Pres3 = ValN(.SubItems(1 + Per3 * 3))
                RsListado!Real3 = ValN(.SubItems(2 + Per3 * 3))
                RsListado!Diferencia3 = ValN(.SubItems(3 + Per3 * 3))
            End If
       End With
       i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    RsListado.Sort = "CuentaContable"
    ListA01_5800.LbPeriodo1 = Format(DateAdd("M", Per1, CalFechaDesde), "MMMM/yyyy")
    If Per2 = -1 Then
        ListA01_5800.LbDif2.Visible = False
        ListA01_5800.LbPeriodo2.Visible = False
        ListA01_5800.LbPres2.Visible = False
        ListA01_5800.LbReal2.Visible = False
        ListA01_5800.TxtDif2.Visible = False
        ListA01_5800.TxtTotDif2.Visible = False
        ListA01_5800.TxtTotPres2.Visible = False
        ListA01_5800.TxtTotReal2.Visible = False
    Else
        ListA01_5800.LbPeriodo2 = Format(DateAdd("M", Per2, CalFechaDesde), "MMMM/yyyy")
    End If
    
    If Per3 = -1 Then
        ListA01_5800.LbDif3.Visible = False
        ListA01_5800.LbPeriodo3.Visible = False
        ListA01_5800.LbPres3.Visible = False
        ListA01_5800.LbReal3.Visible = False
        ListA01_5800.TxtDif3.Visible = False
        ListA01_5800.TxtTotDif3.Visible = False
        ListA01_5800.TxtTotPres3.Visible = False
        ListA01_5800.TxtTotReal3.Visible = False
    Else
        ListA01_5800.LbPeriodo3 = Format(DateAdd("M", Per3, CalFechaDesde), "MMMM/yyyy")
    End If

    ListA01_5800.TxtCentro.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5800.TxtPeriodo.Text = "Desde: " & Format(CalFechaDesde.Value, "MMMM/yyyy") & " Hasta: " & Format(CalFechaHasta.Value, "MMMM/yyyy")
    
    ListA01_5800.DataControl1.Recordset = RsListado
    ListA01_5800.Zoom = -1
End Sub


Private Sub CmdTraer_Click()

    Call CargarLV(CalFechaDesde.Value, CalFechaHasta.Value)
End Sub

Private Sub Form_Load()
    CalFechaDesde.Value = Date
    CalFechaDesde.MaxDate = Date
    CalFechaHasta.Value = Date
    CalFechaHasta.MinDate = Date
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    CmbCentroDeCostoEmisor.ListIndex = 0
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A015800") = 2
    Call CrearEncabezado
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

