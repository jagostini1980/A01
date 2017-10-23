VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5B800 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evolución Presupuestado/Real por Cuenta"
   ClientHeight    =   8115
   ClientLeft      =   3225
   ClientTop       =   2250
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDetalleCont 
      Caption         =   "Detalle &Contable"
      Enabled         =   0   'False
      Height          =   375
      Left            =   225
      TabIndex        =   12
      Top             =   7650
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1755
      TabIndex        =   11
      Top             =   7650
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   150
      Top             =   7590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Height          =   375
      Left            =   3300
      TabIndex        =   10
      Top             =   7650
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   38
      TabIndex        =   5
      Top             =   15
      Width           =   6255
      Begin VB.CommandButton CmdTraer 
         Caption         =   "Traer"
         Height          =   315
         Left            =   5055
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalFechaDesde 
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   100597763
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   315
         Left            =   3660
         TabIndex        =   1
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   100597763
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCuentas 
         Height          =   315
         Left            =   1590
         TabIndex        =   8
         Top             =   570
         Width           =   3420
         _ExtentX        =   6033
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
      Begin VB.Label LbCentroEmisor 
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
         Left            =   60
         TabIndex        =   9
         Top             =   630
         Width           =   1485
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
         Left            =   180
         TabIndex        =   7
         Top             =   255
         Width           =   1320
      End
      Begin VB.Label LBFechaHasta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta:"
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
         Left            =   3030
         TabIndex        =   6
         Top             =   255
         Width           =   570
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4815
      TabIndex        =   4
      Top             =   7650
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6525
      Left            =   60
      TabIndex        =   3
      Top             =   1035
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11509
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
Attribute VB_Name = "A01_5B800"
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
   Dim Nombre As String
   Dim i As Integer
   
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   MousePointer = vbHourglass

On Error GoTo ErrorTraer:
   If CalFechaDesde.Value > CalFechaHasta.Value Then
        MsgBox "Rango de fechas no válido", vbInformation
        CalFechaHasta.SetFocus
        Exit Sub
   End If
   
    Sql = "SpOcConsultaEvolucionCuentaContablePresRealTraer @PeriodoDesde ='" + Format(FechaDesde, "yyyy/MM") + _
                                                        "', @PeriodoHasta ='" + Format(FechaHasta, "yyyy/MM") + _
                                                        "', @CuentaContable='" + VecCuentasContables(CmbCuentas.ListIndex).Codigo & "'"

   With RsListado
         LvListado.Sorted = False
         .Open Sql, Conec
         'limpia el LV
         LvListado.ListItems.Clear
         If .RecordCount > 0 Then

             If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    LvListado.ListItems.Add
                    LvListado.ListItems(i).Text = !Periodo
                    LvListado.ListItems(i).SubItems(1) = Format(VerificarNulo(!TotPres, "N"), "0.00")
                    LvListado.ListItems(i).SubItems(2) = Format(VerificarNulo(!TotReal, "N"), "0.00")
                    LvListado.ListItems(i).SubItems(3) = Format(VerificarNulo(!TotPres, "N") - VerificarNulo(!TotReal, "N"), "0.00")
                    .MoveNext
                Next
             End If
         End If
   End With

    'LvListado.Sorted = True
    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal
    CmdImprimir.Enabled = True
    CmdDetalleCont.Visible = True
    CmdDetalleCont.Enabled = True
    
End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Periodo", (LvListado.Width - 270) / 4
    LvListado.ColumnHeaders.Add , , "Pres.", (LvListado.Width - 270) / 4, 1
    LvListado.ColumnHeaders.Add , , "Real", (LvListado.Width - 270) / 4, 1
    LvListado.ColumnHeaders.Add , , "Diferencia", (LvListado.Width - 270) / 4, 1
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

Private Sub CmdDetalleCont_Click()
    If Not LvListado.SelectedItem Is Nothing Then
        
        Call A01_5B810.Traer(CalFechaDesde.Value, CalFechaHasta.Value, VecCuentasContables(CmbCuentas.ListIndex).Codigo)
        A01_5B810.Show vbModal
    End If
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
        .ActiveCell.FormulaR1C1 = "Periodo Desde : " & Format(CalFechaDesde, "MM/yyyy") & " Hasta: " & Format(CalFechaHasta, "MM/yyyy")
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Cuenta Contable: " & CmbCuentas.Text

        ColorFondo = &HC0E0FF
        Call FormatearExcel(ex, LvListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeConsulta
    ListA01_5B800.Show
End Sub


Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim j As Integer
  Dim RsListado As New ADODB.Recordset
  
    RsListado.Fields.Append "Presupuestado", adVarChar, 25
    RsListado.Fields.Append "Periodo", adVarChar, 25
    RsListado.Fields.Append "RealContable", adVarChar, 25
    RsListado.Fields.Append "Diferencia", adVarChar, 25
   
    RsListado.Open
    i = 0

    For i = 1 To LvListado.ListItems.Count
        RsListado.AddNew
        RsListado!Periodo = LvListado.ListItems(i).Text
        RsListado!Presupuestado = LvListado.ListItems(i).ListSubItems(1)
        RsListado!RealContable = LvListado.ListItems(i).ListSubItems(2)
        RsListado!Diferencia = LvListado.ListItems(i).ListSubItems(3)
    Next
    
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    ListA01_5B800.TxtCuentaContable = CmbCuentas.Text
    ListA01_5B800.TxtPeriodo = Format(CalFechaDesde.Value, "MMMM/yyyy") & " Hasta " & Format(CalFechaHasta.Value, "MMMM/yyyy")
    ListA01_5B800.DataControl1.Recordset = RsListado
    ListA01_5B800.Zoom = -1
End Sub


Private Sub CmdTraer_Click()
    Call CargarLV(CalFechaDesde.Value, CalFechaHasta.Value)
End Sub

Private Sub Form_Load()
    CalFechaDesde.Value = Date
    CalFechaDesde.MaxDate = Date
    CalFechaHasta.Value = Date
    CalFechaHasta.MinDate = Date
    Call CargarComboCuentasContables(CmbCuentas)
    Call CrearEncabezado
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub


