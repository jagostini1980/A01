VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5200 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por Cuenta Contable - Centro de Costo"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   7020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5200.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5200.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5200.frx":2ACC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TVListado 
      Height          =   5325
      Left            =   90
      TabIndex        =   8
      Top             =   1575
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9393
      _Version        =   393217
      Indentation     =   794
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   11865
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Traer"
         Height          =   315
         Left            =   6480
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   405
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   2430
         TabIndex        =   0
         Top             =   255
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM/yyyy"
         Format          =   23461891
         UpDown          =   -1  'True
         CurrentDate     =   38940
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   330
         Left            =   2430
         TabIndex        =   1
         Top             =   675
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
         Left            =   270
         TabIndex        =   7
         Top             =   750
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
         Left            =   1575
         TabIndex        =   6
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   10665
      TabIndex        =   4
      Top             =   7425
      Width           =   1230
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   9315
      TabIndex        =   3
      Top             =   7425
      Width           =   1230
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desvío"
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
      Left            =   9720
      TabIndex        =   15
      Top             =   1305
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desvío %"
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
      Left            =   10890
      TabIndex        =   14
      Top             =   1305
      Width           =   825
   End
   Begin VB.Label LBReal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Real: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6345
      TabIndex        =   13
      Top             =   7020
      Width           =   1125
   End
   Begin VB.Label LBPres 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Presupuestado: $"
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
      Left            =   9000
      TabIndex        =   12
      Top             =   7020
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Presup."
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
      Left            =   8190
      TabIndex        =   11
      Top             =   1305
      Width           =   1470
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Real"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6705
      TabIndex        =   10
      Top             =   1305
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuenta Contable - Centro de Costo"
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
      TabIndex        =   9
      Top             =   1305
      Width           =   6375
   End
End
Attribute VB_Name = "A01_5200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoCuenta
    CuentaContable As String
    DescCta As String
    R_Total As Double
    P_Total As Double
End Type

Private Type TipoCentro
    CuentaContable As String
    DescCentro As String
    Total As Double
End Type

Private Cuenta() As TipoCuenta
Private CentoDeCosto() As TipoCentro

Private Nivel As Integer
Private AlterColor As Boolean

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5200.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    RsListado.Fields.Append "CodCuenta", adVarChar, 4
    RsListado.Fields.Append "Cuenta", adVarChar, 50
    RsListado.Fields.Append "Total", adDouble
    RsListado.Fields.Append "TotalPres", adDouble

    RsListado.Open
    i = 1
    While i <= UBound(Cuenta)
        RsListado.AddNew
      With Cuenta(i)
        RsListado!CodCuenta = .CuentaContable
        RsListado!Cuenta = .DescCta
        RsListado!Total = .R_Total
        RsListado!TotalPres = .P_Total
      End With
        i = i + 1
    Wend
    
    For i = 1 To UBound(CentoDeCosto)
        RsListado.AddNew
        With CentoDeCosto(i)
            RsListado!CodCuenta = .CuentaContable
            RsListado!Cuenta = "        " & .DescCentro
            RsListado!Total = .Total
        End With
    Next
    
    RsListado.MoveFirst
    RsListado.Sort = "CodCuenta"
    
    ListA01_5200.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5200.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5200.DataControl1.Recordset = RsListado
    ListA01_5200.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
    Call CargarTreeView(CalPeriodo.Value)
    Call CalcularTotales
    CmdImprimir.Enabled = TVListado.Nodes.Count > 0
End Sub

Private Sub CargarTreeView(Periodo As Date)
Dim Item As String
Dim Sql As String
Dim i As Integer

Dim Ctro As String
Dim Desvio As Double
Dim Total As String
Dim TotalPres As String

Dim RsCargar As ADODB.Recordset
Set RsCargar = New ADODB.Recordset

'On Error GoTo Error
    
    TVListado.Nodes.Clear
    'pone el príodo en el primer día del mes
    Periodo = "01/" + CStr(Format(Periodo, "MM/yyyy"))
    
    Sql = "SpOCConsultaPorCuentaTraerCuentas " + _
                   "@PeriodoCta = " + FechaSQL(CStr(Periodo), "SQL") + _
                 ", @PeriodoPres = '" + CStr(Format(Periodo, "MM/yyyy")) + _
                "', @CentroDeCostoEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
    'trae los artículo de ese período
    RsCargar.Open Sql, Conec
  With RsCargar
      ReDim Cuenta(.RecordCount)
        i = 1
    While Not .EOF
    'carga las cuentas en un vertor para una eventual impresión
        Cuenta(i).CuentaContable = !CuentaContable
        Cuenta(i).DescCta = BuscarDescCta(!CuentaContable)
        Cuenta(i).R_Total = VerificarNulo(!O_Total, "N")
        Cuenta(i).P_Total = VerificarNulo(!P_Total, "N")
        
     'agrega el artículo en el tv sin totales
        Desvio = Cuenta(i).R_Total - Cuenta(i).P_Total
        Item = Cuenta(i).DescCta + Space(85 - Len(Cuenta(i).DescCta) - Len(Format(Cuenta(i).R_Total, "0.00"))) + Format(Cuenta(i).R_Total, "0.00")
        Item = Item + Space(102 - Len(Item) - Len(Format(Cuenta(i).P_Total, "0.00"))) + Format(Cuenta(i).P_Total, "0.00")
        Item = Item + Space(115 - Len(Item) - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00")
        If Cuenta(i).P_Total = 0 Then
           Desvio = 1
        Else
            Desvio = Desvio / Cuenta(i).P_Total
        End If
        Item = Item + Space(124 - Len(Item) - Len(Format(Desvio, "0.00%"))) + Format(Desvio, "0.00%")

        TVListado.Nodes.Add , , CStr(!CuentaContable) + "C", Item, 1
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = &HFFC0C0
        i = i + 1
        .MoveNext
    Wend
        .Close
        
    Sql = "SpOCConsultaPorCuentaTraerCentro " + _
            "@PeriodoCta =" + FechaSQL(CStr(Periodo), "SQL") + "," + _
            "@PeriodoPres = '" + Format(Periodo, "MM/yyyy") + "'," + _
            "@CentroDeCostoEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
   'trae los Centros de costos que participan del período
     .Open Sql, Conec
      ReDim CentoDeCosto(.RecordCount)
     i = 1
    While Not .EOF
        If !CentroDeCosto = 0 Then
           Ctro = "(Presupuestado Sin Asignar)"
        Else
           Ctro = BuscarDescCentro(!CentroDeCosto)
        End If
        
        Total = VerificarNulo(!O_Total, "N")
       
        CentoDeCosto(i).CuentaContable = !CuentaContable
        CentoDeCosto(i).DescCentro = Ctro
        CentoDeCosto(i).Total = Total
        i = i + 1
        'TotalPres = VerificarNulo(!P_Total, "N")
        'Desvio = Total - TotalPres
        Item = Ctro + Space(80 - Len(Ctro) - Len(Format(Total, "0.00"))) + Format(Total, "0.00")
        'Item = Item + Space(97 - Len(Item) - Len(Format(TotalPres, "0.00"))) + Format(TotalPres, "0.00")
        'Item = Item + Space(110 - Len(Item) - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00")
        
        'If Val(TotalPres) = 0 Then
        '    Item = Item + Space(119 - Len(Item) - Len(Format("1", "0.00%"))) + Format("1", "0.00%")
        'Else
        '    Item = Item + Space(119 - Len(Item) - Len(Format(Desvio / TotalPres, "0.00%"))) + Format(Desvio / TotalPres, "0.00%")
        'End If
      'agraga el Tv el nodo con la Cta - centro que componen el artículo
        TVListado.Nodes.Add CStr(!CuentaContable) + "C", tvwChild, , Item, 2
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
        AlterColor = Not AlterColor
      .MoveNext
    Wend

        .Close
  End With
   
   TVListado.Sorted = True
    
   For i = 1 To TVListado.Nodes.Count
        TVListado.Nodes(i).Expanded = True
        TVListado.Nodes(i).Sorted = True
   Next
   
   Dim Nodo
   i = 1
   If TVListado.Nodes.Count > 0 Then
    While i < TVListado.Nodes.Count
         Set Nodo = TVListado.Nodes(i).Child
         AlterColor = True
         While Not Nodo Is Nothing
         'pone la alteración de colores
            Nodo.BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
            AlterColor = Not AlterColor
            Set Nodo = Nodo.Next
         Wend
         
         i = i + 1
     Wend
    End If
Error:
    Call ManipularError(Err.Number, Err.Description)

End Sub

Private Sub Form_Load()
   
    CalPeriodo.Value = Date
        
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor, "Todos")
    
    Nivel = TraerNivel("A015200")
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
  
    For i = 1 To UBound(Cuenta)
        Total = Total + Cuenta(i).R_Total
        TotalPres = TotalPres + Cuenta(i).P_Total
    Next
        LBReal.Caption = "Total Real: $" + Format(Total, "0.00")
        LBPres.Caption = "Total Presupuestado: $" + Format(TotalPres, "0.00")
End Sub

