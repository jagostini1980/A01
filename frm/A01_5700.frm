VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5700 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Presupuestado/Contable/SGP"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDetallePres 
      Caption         =   "Detalle Presupuestado"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2460
      TabIndex        =   9
      Top             =   7065
      Width           =   1860
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   360
      Top             =   7065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   2460
      TabIndex        =   7
      Top             =   7470
      Width           =   1550
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4085
      TabIndex        =   8
      Top             =   7470
      Width           =   1550
   End
   Begin VB.CommandButton CmdDetalleSGP 
      Caption         =   "&Detalle SGP de la Cuenta"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6780
      TabIndex        =   11
      Top             =   7065
      Width           =   2085
   End
   Begin VB.CommandButton CmdDetalleContable 
      Caption         =   "&Detalle Contable de la Cuenta"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4380
      TabIndex        =   10
      Top             =   7050
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   880
      Left            =   45
      TabIndex        =   14
      Top             =   0
      Width           =   11235
      Begin VB.CheckBox ChkConSgp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Con SGP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9105
         TabIndex        =   18
         Top             =   585
         Width           =   1590
      End
      Begin VB.OptionButton OptConsulta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Presupuestado > Contable"
         Height          =   195
         Index           =   1
         Left            =   3375
         TabIndex        =   17
         Top             =   585
         Width           =   2265
      End
      Begin VB.OptionButton OptConsulta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
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
         Index           =   3
         Left            =   8100
         TabIndex        =   4
         Top             =   585
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptConsulta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Presupuestado < Contable"
         Height          =   195
         Index           =   2
         Left            =   5760
         TabIndex        =   3
         Top             =   585
         Width           =   2265
      End
      Begin VB.OptionButton OptConsulta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Presupuestado = Contable"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   585
         Width           =   2265
      End
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H80000003&
         Caption         =   "Traer"
         Height          =   315
         Left            =   9000
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   218
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   210
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM/yyyy"
         Format          =   53673987
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   5310
         TabIndex        =   1
         Top             =   210
         Width           =   3570
         _ExtentX        =   6297
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo Emisor:"
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
         Left            =   3150
         TabIndex        =   16
         Top             =   278
         Width           =   2055
      End
      Begin VB.Label LbFechaDesde 
         Alignment       =   1  'Right Justify
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
         Left            =   225
         TabIndex        =   15
         Top             =   285
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   7335
      TabIndex        =   13
      Top             =   7470
      Width           =   1550
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5710
      TabIndex        =   12
      Top             =   7470
      Width           =   1550
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6015
      Left            =   45
      TabIndex        =   6
      Top             =   945
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   10610
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
Attribute VB_Name = "A01_5700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Nivel As Integer
Private VecRepCuentas() As String

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdDetallePres_Click()
    MousePointer = vbHourglass
    A01_5730.Cuenta = LvListado.SelectedItem.SubItems(6)
    A01_5730.Periodo = CalPeriodo.Value
    A01_5730.CentroEmisor = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia
    A01_5730.Show vbModal
    MousePointer = vbNormal

End Sub

Private Sub CmdDetalleSGP_Click()
    MousePointer = vbHourglass
    A01_5710.Cuenta = LvListado.SelectedItem.SubItems(6)
    A01_5710.Periodo = CalPeriodo.Value
    A01_5710.CentroEmisor = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo
    A01_5710.Show vbModal
    MousePointer = vbNormal
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

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5700.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 100
    RsListado.Fields.Append "TotPres", adDouble
    RsListado.Fields.Append "TotContable", adDouble
    RsListado.Fields.Append "TotSGP", adDouble
    RsListado.Fields.Append "DesvioPorc", adVarChar, 20
    RsListado.Fields.Append "DesvioPorcSGP", adVarChar, 20
    
    RsListado.Open
    i = 1
    While i < LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
            RsListado!CuentaContable = .Text
            RsListado!TotPres = ValN(.SubItems(1))
            RsListado!TotContable = ValN(.SubItems(2))
            RsListado!TotSGP = ValN(.SubItems(3))
            RsListado!DesvioPorc = .SubItems(4)
            RsListado!DesvioPorcSGP = .SubItems(5)
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    RsListado.Sort = "CuentaContable"
    For i = 0 To OptConsulta.Count - 1
        If OptConsulta(i).Value Then
            ListA01_5700.TxtConsultado = OptConsulta(i).Caption
        End If
    Next
    ListA01_5700.TxtCentro.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5700.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5700.DataControl1.Recordset = RsListado
    ListA01_5700.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
   Call CargarListado(CalPeriodo.Value)
  
  'si se carga algún nodo se Habilita la impresión
   CmdImprimir.Enabled = LvListado.ListItems.Count > 0
   CmdExpPdf.Enabled = LvListado.ListItems.Count > 0
End Sub

Private Sub CargarListado(Periodo As Date)
Dim Sql As String
Dim i As Integer
Dim TotPres As Double
Dim TotContable As Double
Dim TotSGP As Double
Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LvListado.ListItems.Clear
    MousePointer = vbHourglass
    'Realiza la consulta
    If OptConsulta(0).Value Then
       Sql = "SpOcConsultaDesvioPresupuestoContableSgpPresIgualCont"
    End If
    If OptConsulta(1).Value Then
       Sql = "SpOcConsultaDesvioPresupuestoContableSgpPresMayorCont"
    End If
    If OptConsulta(2).Value Then
       Sql = "SpOcConsultaDesvioPresupuestoContableSgpPresMenorCont"
    End If
    If OptConsulta(3).Value Then
        If ChkConSgp.Value = 0 Then
            Sql = "SpOcConsultaDesvioPresupuestoContableSGP"
        Else
            Sql = "SpOcConsultaDesvioPresupuestoContableSGPSoloConSgp"
        End If
    End If
        
    Sql = Sql & " @Periodo = '" & Format(CalPeriodo.Value, "MM/yyyy") & _
              "', @PeriodoSGP = " & FechaSQL(CalPeriodo.Value, "SQL") & _
               ", @CentroEmisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'" & _
               ", @Emisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia & "'"
    
    With RsCargar
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .Open Sql, Conec
        
        ReDim VecRepCuentas(.RecordCount)
      LvListado.Sorted = False
        While Not .EOF
            i = i + 1
            LvListado.ListItems.Add
            
            LvListado.ListItems(i).Text = BuscarDescCta(!CuentaContable) & " - Cod. " & !CuentaContable
            LvListado.ListItems(i).SubItems(1) = Format(ValN(!TotalPres), "0.00")
            LvListado.ListItems(i).SubItems(2) = Format(ValN(!Contable), "0.00")
            LvListado.ListItems(i).SubItems(3) = Format(ValN(!TotalSGP), "0.00")
            'LVListado.ListItems(i).SubItems(4) = Format(0 - ValN(!TotalPres), "0.00")
            If ValN(!TotalPres) = 0 Then
                LvListado.ListItems(i).SubItems(4) = "100.00 %"
            Else
               LvListado.ListItems(i).SubItems(4) = Format((ValN(!Contable) - ValN(!TotalPres)) / ValN(!TotalPres), "0.00 %")
            End If
            If ValN(!TotalSGP) = 0 Then
                LvListado.ListItems(i).SubItems(5) = "100.00 %"
            Else
               LvListado.ListItems(i).SubItems(5) = Format((ValN(!Contable) - ValN(!TotalSGP)) / ValN(!TotalSGP), "0.00 %")
            End If
            LvListado.ListItems(i).SubItems(6) = !CuentaContable
            TotPres = TotPres + ValN(!TotalPres)
            TotContable = TotContable + ValN(!Contable)
            TotSGP = TotSGP + ValN(!TotalSGP)
            .MoveNext
        Wend
    End With
    LvListado.Sorted = True
    LvListado.Sorted = False
    
    LvListado.ListItems.Add
    LvListado.ListItems(LvListado.ListItems.Count).Text = "Totales ==>"
    LvListado.ListItems(LvListado.ListItems.Count).SubItems(1) = Format(TotPres, "0.00")
    LvListado.ListItems(LvListado.ListItems.Count).SubItems(2) = Format(TotContable, "0.00")
    LvListado.ListItems(LvListado.ListItems.Count).SubItems(3) = Format(TotSGP, "0.00")
    
    LvListado.ListItems(LvListado.ListItems.Count).Bold = True
    LvListado.ListItems(LvListado.ListItems.Count).ListSubItems(1).Bold = True
    LvListado.ListItems(LvListado.ListItems.Count).ListSubItems(2).Bold = True
    LvListado.ListItems(LvListado.ListItems.Count).ListSubItems(3).Bold = True

Error:
    MousePointer = vbNormal
    Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub CmdDetalleContable_Click()
    MousePointer = vbHourglass
    A01_5720.Cuenta = LvListado.SelectedItem.SubItems(6)
    A01_5720.Periodo = CalPeriodo.Value
    A01_5720.CentroEmisor = VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia
    A01_5720.Show vbModal
    MousePointer = vbNormal
End Sub

Private Sub Form_Load()
   
    CalPeriodo.Value = Date
    Call CrearEncabezado
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    
    Nivel = TraerNivel("A015700")
    If Nivel = 2 Then
        CmbCentroDeCostoEmisor.ListIndex = 0
    Else
        Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
        CmbCentroDeCostoEmisor.Enabled = False
        CmdTraer.Enabled = CentroEmisor <> ""
        If CentroEmisor = "" Then
            MsgBox "Ud. No tiene un Centro de Costos Asociado", vbInformation
        End If
    End If

End Sub

Private Sub CalcularTotales()
  Dim i As Integer
  Dim Total As Double
  Dim TotalPres As Double
  
    For i = 1 To LvListado.ListItems.Count
        Total = Total + LvListado.ListItems(i).SubItems(1)
        TotalPres = TotalPres + LvListado.ListItems(i).SubItems(2)
    Next
    
End Sub

Private Sub CrearEncabezado()

    LvListado.ColumnHeaders.Add , , "Cuenta Contable", LvListado.Width - 7350
    LvListado.ColumnHeaders.Add , , "Presupuestado", 1300, 1
    LvListado.ColumnHeaders.Add , , "Real Contable", 1200, 1
    LvListado.ColumnHeaders.Add , , "Real SGP", 1200, 1
    LvListado.ColumnHeaders.Add , , "Desvio % Pres./Cont.", 1700, 1
    LvListado.ColumnHeaders.Add , , "Desvio % Cont./SGP", 1700, 1
    LvListado.ColumnHeaders.Add , , "CodCuenta", 0
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CmdDetalleContable.Enabled = Item.Index < LvListado.ListItems.Count
    CmdDetalleSGP.Enabled = Item.Index < LvListado.ListItems.Count
    CmdDetallePres.Enabled = Item.Index < LvListado.ListItems.Count
End Sub

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5700.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export ListA01_5700.Pages
         Unload ListA01_5700
  End If
Error:
    If Err.Number = 0 Then
        MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
    Else
        Call ManipularError(Err.Number, Err.Description)
    End If
End Sub

Private Sub GenerarPlanilla(NombreArchivo As String, Filtro As Integer)
Dim ex As Excel.Application
Dim col As Integer
Dim ColorFondo As Long
Dim i As Integer

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcel(ex, LvListado, Caption, 8)
        Call DatosExcel(ex, LvListado, 8)
        
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
        .ActiveCell.FormulaR1C1 = "Periodo: " & CalPeriodo
        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Centro de Costo: " & CmbCentroDeCostoEmisor.Text
        .Range("A6").Select
        For i = 0 To OptConsulta.Count - 1
            If OptConsulta(i).Value Then
                .ActiveCell.FormulaR1C1 = "Consultado Por: " & OptConsulta(i).Caption
            End If
        Next

        ColorFondo = &HC0E0FF
        Call FormatearExcel(ex, LvListado, 8, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub

Private Sub OptConsulta_Click(Index As Integer)
    ChkConSgp.Enabled = Index = 3
End Sub
