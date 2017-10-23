VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5600 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Desvio Presupuestado/Contable"
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
   Begin VB.CommandButton CmdExp 
      Caption         =   "&Exportar Excel"
      Height          =   350
      Left            =   5865
      TabIndex        =   10
      Top             =   7455
      Width           =   1365
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7335
      TabIndex        =   9
      Top             =   7470
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   11235
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H80000003&
         Caption         =   "Traer"
         Height          =   315
         Left            =   9045
         MaskColor       =   &H00000000&
         TabIndex        =   2
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
         Format          =   24248323
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   5355
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
         Left            =   3195
         TabIndex        =   8
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
         TabIndex        =   6
         Top             =   285
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10035
      TabIndex        =   4
      Top             =   7470
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8685
      TabIndex        =   3
      Top             =   7470
      Width           =   1230
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6690
      Left            =   45
      TabIndex        =   7
      Top             =   675
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   11800
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
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   45
      Top             =   7380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "A01_5600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Nivel As Integer

Private Sub CmdCerra_Click()
    Unload Me
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
   ListA01_5600.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "CuentaContable", adVarChar, 100
    RsListado.Fields.Append "TotPres", adDouble
    RsListado.Fields.Append "TotContable", adDouble
    RsListado.Fields.Append "Desvio", adDouble
    RsListado.Fields.Append "DesvioPorc", adVarChar, 20
    
    RsListado.Open
    i = 1
    While i < LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
            RsListado!CuentaContable = .Text
            RsListado!TotPres = ValN(.SubItems(1))
            RsListado!TotContable = ValN(.SubItems(2))
            RsListado!Desvio = ValN(.SubItems(3))
            RsListado!DesvioPorc = .SubItems(4)
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
       RsListado.MoveFirst
    End If
    
    ListA01_5600.TxtCentro.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5600.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5600.DataControl1.Recordset = RsListado
    ListA01_5600.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
   Call CargarListado(CalPeriodo.Value)
  
  'si se carga algún nodo se Habilita la impresión
   CmdImprimir.Enabled = LVListado.ListItems.Count > 0
   CmdExpPdf.Enabled = LVListado.ListItems.Count > 0
End Sub

Private Sub CargarListado(Periodo As Date)
Dim Sql As String
Dim i As Integer
Dim TotPres As Double
Dim TotContable As Double
Dim RsCargar As New ADODB.Recordset

On Error GoTo Error
    LVListado.ListItems.Clear

    'Realiza la consulta
    Sql = "SpOcConsultaDesvioPresupuestoContable @Periodo = '" & Format(CalPeriodo.Value, "MM/yyyy") & _
                                             "', @CentroEmisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                                             "', @Emisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Jerarquia & "'"
    With RsCargar
        .Open Sql, Conec
        LVListado.Sorted = False
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add
            LVListado.ListItems(i).Text = BuscarDescCta(!P_CuentaContable) & " - Cod. " & !P_CuentaContable
            LVListado.ListItems(i).SubItems(1) = Format(ValN(!TotalPres), "0.00")
            LVListado.ListItems(i).SubItems(2) = Format(!Contable, "0.00")
            LVListado.ListItems(i).SubItems(3) = Format(!Contable - ValN(!TotalPres), "0.00")
            If ValN(!TotalPres) <> 0 Then
                LVListado.ListItems(i).SubItems(4) = Format((!Contable - ValN(!TotalPres)) / ValN(!TotalPres), "0.00 %")
            End If
            
            TotPres = TotPres + ValN(!TotalPres)
            TotContable = TotContable + ValN(!Contable)

            .MoveNext
        Wend
    End With
    
    LVListado.Sorted = True
    LVListado.Sorted = False
    
    LVListado.ListItems.Add
    LVListado.ListItems(LVListado.ListItems.Count).Text = "Totales ==>"
    LVListado.ListItems(LVListado.ListItems.Count).SubItems(1) = Format(TotPres, "0.00")
    LVListado.ListItems(LVListado.ListItems.Count).SubItems(2) = Format(TotContable, "0.00")
    
    LVListado.ListItems(LVListado.ListItems.Count).Bold = True
    LVListado.ListItems(LVListado.ListItems.Count).ListSubItems(1).Bold = True
    LVListado.ListItems(LVListado.ListItems.Count).ListSubItems(2).Bold = True

Error:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
   
    CalPeriodo.Value = Date
    Call CrearEncabezado
    
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    
    Nivel = TraerNivel("A015600")
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
  
    For i = 1 To LVListado.ListItems.Count
        Total = Total + LVListado.ListItems(i).SubItems(1)
        TotalPres = TotalPres + LVListado.ListItems(i).SubItems(2)
    Next
    
End Sub

Private Sub CrearEncabezado()

    LVListado.ColumnHeaders.Add , , "Cuenta Contable", LVListado.Width - 5250
    LVListado.ColumnHeaders.Add , , "Presupuestado", 1400, 1
    LVListado.ColumnHeaders.Add , , "Real Contable", 1200, 1
    LVListado.ColumnHeaders.Add , , "Desvio", 1200, 1
    LVListado.ColumnHeaders.Add , , "Desvio %", 1200, 1
    
End Sub


Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeConsulta
         ListA01_5600.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export ListA01_5600.Pages
         Unload ListA01_5600
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

    Set ex = New Excel.Application
    With ex
        '---------GENERO LIBRO Y HOJA ---------------------------
        .Workbooks.Add
        .Sheets.Add
        '-------- GENERO LOS DATOS ------------------------------
        Call EncabezadoExcel(ex, LVListado, Caption, 6)
        Call DatosExcel(ex, LVListado, 6)
        
        '--------AJUSTO LOS TAMAÑOS DE LAS COLUMNAS
        For col = 1 To LVListado.ColumnHeaders.Count
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

        ColorFondo = &HC0E0FF
        Call FormatearExcel(ex, LVListado, 6, ColorFondo)
    End With
    Call GuardarPlanilla(ex, NombreArchivo, Filtro)
    ex.ActiveWorkbook.Close
    MsgBox "Exportacion Finalizada"
End Sub


