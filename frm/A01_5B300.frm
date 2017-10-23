VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5B300 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Evolución Movilidades Presupuestado/Real"
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
      Left            =   4462
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
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   49676291
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin MSComCtl2.DTPicker CalFechaHasta 
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   49676291
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   7050
         TabIndex        =   8
         Top             =   225
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
         Top             =   285
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
      Left            =   5977
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
Attribute VB_Name = "A01_5B300"
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
   Dim Nombre As String
   Dim pos As Integer
   Dim j As Integer
   Dim CodSC As String
   Dim i As Integer
   
   RsListado.CursorLocation = adUseClient
   RsListado.CursorType = adOpenKeyset
   MousePointer = vbHourglass

'On Error GoTo ErrorTraer:
   If CalFechaDesde.Value > CalFechaHasta.Value Then
        MsgBox "Rango de fechas no válido", vbInformation
        CalFechaHasta.SetFocus
        Exit Sub
   Else
        LvListado.ColumnHeaders.Clear
        LvListado.ColumnHeaders.Add , , "Nombre", 3000
        For i = 0 To DateDiff("M", CalFechaDesde, CalFechaHasta)
            LvListado.ColumnHeaders.Add , , "Pres.", 1300, 1
            LvListado.ColumnHeaders.Add , , "Real", 1300, 1
            LvListado.ColumnHeaders.Add , , "Diferencia " & Format(DateAdd("M", i, CalFechaDesde), "MMM/yy"), 1500, 1
        Next
        LvListado.ColumnHeaders.Add , , , 0
   End If
   
   Sql = "SpOcConsultaEvolucionMovilidadesPresupuestadoTraer @PeriodoDesde ='" & "01/" & Format(CalFechaDesde, "MM/yyyy") + _
                                                         "', @PeriodoHasta ='" & "01/" & Format(CalFechaHasta, "MM/yyyy") + _
                                                         "', @CentroDeCostosEmisor='" + VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
   
   With RsListado
        
         LvListado.Sorted = False
         .Open Sql, Conec
         
         'limpia el LV
         LvListado.ListItems.Clear
         If .RecordCount > 0 Then
          
            For i = 1 To .RecordCount
              
              If Nombre <> !C_Descripcion Then
                  LvListado.ListItems.Add
                  pos = pos + 1
                  LvListado.ListItems(pos).Text = Convertir(!C_Descripcion)
                  LvListado.ListItems(pos).Tag = !P_SubCentroDeCosto
                  Nombre = !C_Descripcion
              End If
              
              NroPeriodo = DateDiff("M", CalFechaDesde, !P_Periodo)
              LvListado.ListItems(pos).SubItems(1 + NroPeriodo * 3) = Format(VerificarNulo(!TotPres, "N"), "0.00")
              'LvListado.ListItems(pos).SubItems(2 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N"), "0.00")
              'LvListado.ListItems(pos).SubItems(3 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N") - VerificarNulo(!TotPres, "N"), "0.00")
              .MoveNext
            Next
            .Close
            Sql = "SpOcConsultaEvolucionMovilidadesRealTraer @PeriodoDesde ='" + Format(FechaDesde, "MM/yyyy") + _
                                                         "', @PeriodoHasta ='" + Format(FechaHasta, "MM/yyyy") + _
                                                         "', @Emisor='" + VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia & "'"
             .Open Sql, Conec
             If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    CodSC = BuscarCodigoPorCodSecundario(Format(!C_Rubro, "000"))
                    For j = 1 To LvListado.ListItems.Count
                        If LvListado.ListItems(j).Tag = CodSC Then
                            Exit For
                        End If
                    Next
                    NroPeriodo = DateDiff("M", CalFechaDesde, !Periodo)
                    If j > LvListado.ListItems.Count Then
                        LvListado.ListItems.Add
                        LvListado.ListItems(j).Tag = CodSC
                        LvListado.ListItems(j).Text = BuscarDescCentro(CodSC)
                        LvListado.ListItems(j).SubItems(2 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N"), "0.00")
                    Else
                        LvListado.ListItems(j).SubItems(2 + NroPeriodo * 3) = Format(VerificarNulo(!TotReal, "N"), "0.00")
                    End If
                    .MoveNext
                Next
             End If
         End If
   End With
   For i = 1 To LvListado.ListItems.Count
       For j = 3 To LvListado.ColumnHeaders.Count - 1 Step 3
            With LvListado.ListItems(i)
                .SubItems(j) = ValN(.SubItems(j - 1)) - ValN(.SubItems(j - 2))
            End With
       Next
       
   Next
    LvListado.Sorted = True
    Set RsListado = Nothing
ErrorTraer:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal

End Sub

Private Sub CrearEncabezado()
    LvListado.ColumnHeaders.Add , , "Nombre", 3000
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
    CmbCentroDeCostoEmisor.Enabled = TraerNivel("A015B300") = 2
    Call CrearEncabezado
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

