VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Begin VB.Form A01_5300 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por Centro de Costo"
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
      Top             =   7245
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
            Picture         =   "A01_5300.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5300.frx":27B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "A01_5300.frx":4F64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   11955
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H80000003&
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
         Format          =   54394883
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
   Begin MSComctlLib.TreeView TVListado 
      Height          =   5280
      Left            =   0
      TabIndex        =   8
      Top             =   1755
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   9313
      _Version        =   393217
      Indentation     =   661
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
      Height          =   400
      Left            =   10980
      TabIndex        =   19
      Top             =   1305
      Width           =   750
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
      Height          =   405
      Left            =   10080
      TabIndex        =   18
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      ForeColor       =   &H00800000&
      Height          =   400
      Left            =   3825
      TabIndex        =   17
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precio Unit. Promedio"
      ForeColor       =   &H00800000&
      Height          =   400
      Left            =   4590
      TabIndex        =   16
      Top             =   1305
      Width           =   990
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
      Height          =   400
      Left            =   5625
      TabIndex        =   15
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad Presup."
      Height          =   400
      Left            =   6570
      TabIndex        =   14
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precio Unit. Promedio Pres."
      Height          =   405
      Left            =   7470
      TabIndex        =   13
      Top             =   1305
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Presupuestado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8685
      TabIndex        =   12
      Top             =   1305
      Width           =   1350
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
      Left            =   5400
      TabIndex        =   11
      Top             =   7110
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
      Left            =   8145
      TabIndex        =   10
      Top             =   7110
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Centro de Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   45
      TabIndex        =   9
      Top             =   1305
      Width           =   3735
   End
End
Attribute VB_Name = "A01_5300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoReporte
    Centro As String
    SubCentro As String
    Articulo As String
    R_Cant As String
    R_PUnit As String
    R_Total As Double
    P_Cant As String
    P_PUnit As String
    P_Total As Double
End Type

Private Type TipoCentro
    CodCentro As String
    DescCentro As String
    R_Total As Double
    P_Total As Double
    Padre As Integer
End Type

Private Centros() As TipoCentro
Private CentrosPadres() As TipoCentro
Private Reporte() As TipoReporte
Private Nivel As Integer
Private AlterColor As Boolean

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
   Call ConfImpresionDeConsulta
   ListA01_5300.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "Centro", adVarChar, 50
    RsListado.Fields.Append "SubCentro", adVarChar, 50
    RsListado.Fields.Append "Articulo", adVarChar, 50
    RsListado.Fields.Append "R_Cant", adVarChar, 25
    RsListado.Fields.Append "R_PUnit", adVarChar, 25
    RsListado.Fields.Append "R_Total", adDouble
    RsListado.Fields.Append "P_Cant", adVarChar, 25
    RsListado.Fields.Append "P_PUnit", adVarChar, 25
    RsListado.Fields.Append "P_Total", adDouble
    
    RsListado.Open
    i = 1
    While i <= UBound(Reporte)
        RsListado.AddNew
      With Reporte(i)
            RsListado!Centro = .Centro
            RsListado!SubCentro = .SubCentro
            RsListado!Articulo = .Articulo
            RsListado!R_Cant = .R_Cant
            RsListado!R_PUnit = .R_PUnit
            RsListado!R_Total = .R_Total
            RsListado!P_Cant = .P_Cant
            RsListado!P_PUnit = .P_PUnit
            RsListado!P_Total = .P_Total

      End With
        i = i + 1
    Wend
    RsListado.MoveFirst
    RsListado.Sort = "Centro"
    RsListado.Sort = "SubCentro"
    RsListado.Sort = "Articulo"
    
    ListA01_5300.TxtPeriodo.Text = Format(CalPeriodo.Value, "MMMM/yyyy")
    ListA01_5300.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    ListA01_5300.DataControl1.Recordset = RsListado

End Sub

Private Function BuscarDescCentroPadre(C_Padre As Integer) As String
    Dim i  As Integer
    For i = 1 To UBound(CentrosPadres)
        If CentrosPadres(i).CodCentro = C_Padre Then
            BuscarDescCentroPadre = CentrosPadres(i).DescCentro
            Exit Function
        End If
    Next
End Function

Private Sub CmdTraer_Click()
   Call CargarTreeView(CalPeriodo.Value)
   Call CalcularTotales
  'si se carga algún nodo se Habilita la impresión
   CmdImprimir.Enabled = TVListado.Nodes.Count > 0
    
End Sub

Private Sub CargarTreeView(Periodo As Date)
Dim Item As String
Dim Sql As String
Dim i As Integer

Dim Art As String
Dim Ctro As String

Dim TotalSup As Double
Dim TotalSupPres As Double

Dim Desvio As Double
Dim DesvioPoc As Double
Dim Total As String
Dim cant As String
Dim PrecioU As String

Dim TotalPres As String
Dim CantPres As String
Dim PrecioUPres As String

Dim RsCargar As ADODB.Recordset
Set RsCargar = New ADODB.Recordset

On Error GoTo Error
   
    TVListado.Nodes.Clear
    ReDim Reporte(0)

    'inserta todos los centros padres
    Sql = "SpTA_CentrosDeCostosPadres"
    RsCargar.Open Sql, Conec
    
    ReDim CentrosPadres(RsCargar.RecordCount)
    i = 1
    While Not RsCargar.EOF
        CentrosPadres(i).CodCentro = RsCargar!C_Codigo
        CentrosPadres(i).DescCentro = RsCargar!C_Descripcion
        
        i = i + 1
        
        TVListado.Nodes.Add , , CStr(RsCargar!C_Codigo) + "CC", RsCargar!C_Descripcion + " ", 1
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = &HFFC0C0
        RsCargar.MoveNext
    Wend
     RsCargar.Close
        
    'pone el príodo en el primer día del mes
    Periodo = "01/" + CStr(Format(Periodo, "MM/yyyy"))
        
    Sql = "SpOCConsultaPorCentroTraerCentros " + _
                  "@PeriodoReal = " + FechaSQL(CStr(Periodo), "SQL") + _
                ", @PeriodoPres = '" + CStr(Format(Periodo, "MM/yyyy")) + _
                "', @CentroDeCostoEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
    'trae los artículo de ese período
    RsCargar.Open Sql, Conec
  With RsCargar
      ReDim Centros(.RecordCount)
        i = 1
    While Not .EOF
    'carga las cuentas en un vertor para una eventual impresión
        Centros(i).CodCentro = !CentroDeCosto
        Centros(i).DescCentro = BuscarDescCentro(!CentroDeCosto)
        Centros(i).R_Total = VerificarNulo(!O_Total, "N")
        Centros(i).P_Total = VerificarNulo(!P_Total, "N")
        Centros(i).Padre = BuscarCentroPadre(!CentroDeCosto)
        
        TotalSup = TotalSup + Centros(i).R_Total
        TotalSupPres = TotalSupPres + Centros(i).P_Total
        
        Desvio = Centros(i).R_Total - Centros(i).P_Total
     'agrega el sub-centro en el tv sin totales
     'segundo nivel
        Item = Centros(i).DescCentro + Space(64 - Len(Centros(i).DescCentro) - Len(Format(Centros(i).R_Total, "0.00"))) + Format(Centros(i).R_Total, "0.00")
        Item = Item + Space(103 - Len(Item) - Len(Format(Centros(i).P_Total, "0.00"))) + Format(Centros(i).P_Total, "0.00")
        Item = Item + Space(114 - Len(Item) - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00")
        
        If Centros(i).P_Total = 0 Then
            Item = Item + Space(122 - Len(Item) - Len(Format("1", "0.00%"))) + Format("1", "0.00%")
        Else
            Item = Item + Space(122 - Len(Item) - Len(Format(Desvio / Centros(i).P_Total, "0.00%"))) + Format(Desvio / Centros(i).P_Total, "0.00%")
        End If
        
        TVListado.Nodes.Add BuscarCentroPadre(!CentroDeCosto) + "CC", tvwChild, CStr(!CentroDeCosto) + "C", Item, 2
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = &HC0C0FF
 
        .MoveNext
        
        If Not .EOF Then
            Desvio = TotalSup - TotalSupPres
            If TotalSupPres = 0 Then
                DesvioPoc = 1
            Else
                DesvioPoc = Desvio / TotalSupPres
            End If

            If BuscarCentroPadre(!CentroDeCosto) <> BuscarCentroPadre(Centros(i).CodCentro) Then
               With TVListado.Nodes(TVListado.Nodes.Count).Parent
                   .Text = .Text + Space(68 - Len(.Text) - Len(Format(TotalSup, "0.00"))) + Format(TotalSup, "0.00") + _
                    Space(39 - Len(Format(TotalSupPres, "0.00"))) + Format(TotalSupPres, "0.00") + _
                    Space(11 - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00") + _
                    Space(8 - Len(Format(DesvioPoc, "0.00%"))) + Format(DesvioPoc, "0.00%")

               End With
                TotalSup = 0
                TotalSupPres = 0
            End If
        End If
         i = i + 1
    Wend
      If Not TVListado.Nodes(TVListado.Nodes.Count).Parent Is Nothing Then
        Desvio = TotalSup - TotalSupPres
        With TVListado.Nodes(TVListado.Nodes.Count).Parent
            If TotalSupPres = 0 Then
             .Text = .Text + Space(68 - Len(.Text) - Len(Format(TotalSup, "0.00"))) + Format(TotalSup, "0.00") + _
                     Space(39 - Len(Format(TotalSupPres, "0.00"))) + Format(TotalSupPres, "0.00") + _
                     Space(11 - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00") + _
                     Space(8 - Len(Format("1", "0.00%"))) + Format("1", "0.00%")
            Else
             .Text = .Text + Space(68 - Len(.Text) - Len(Format(TotalSup, "0.00"))) + Format(TotalSup, "0.00") + _
                     Space(39 - Len(Format(TotalSupPres, "0.00"))) + Format(TotalSupPres, "0.00") + _
                     Space(11 - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00") + _
                     Space(8 - Len(Format(Desvio / TotalSupPres, "0.00%"))) + Format(Desvio / TotalSupPres, "0.00%")
            End If
        End With
      End If
        .Close
        
    Sql = "SpOCConsultaPorCentroTraerArticulos " + _
            "@PeriodoReal =" + FechaSQL(CStr(Periodo), "SQL") + "," + _
            "@PeriodoPres = '" + Format(Periodo, "MM/yyyy") + "'," + _
            "@CentroDeCostoEmisor = '" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & "'"
   'trae los Art que participan del período
     .Open Sql, Conec
     i = 1
    While Not .EOF
        ReDim Preserve Reporte(UBound(Reporte) + 1)
        
        Art = IIf(!CodigoArticulo = 0, "Servicio", BuscarDescArt(!CodigoArticulo, BuscarTablaCentroEmisor(!CentroDeCostoEmisor)))
        Total = VerificarNulo(!O_Total, "N")
        cant = VerificarNulo(!O_CantidadRecibida, "N")
        If cant > 0 Then
            PrecioU = Total / cant
        Else
            PrecioU = ""
        End If
        
        TotalPres = VerificarNulo(!P_Total, "N")
        CantPres = VerificarNulo(!P_Cantidad, "N")
        If CantPres > 0 Then
            PrecioUPres = TotalPres / CantPres
        Else
            PrecioUPres = ""
        End If
   'datos para el reporte
        Reporte(UBound(Reporte)).Articulo = Art
        Reporte(UBound(Reporte)).R_Cant = cant
        Reporte(UBound(Reporte)).R_PUnit = PrecioU
        Reporte(UBound(Reporte)).R_Total = Total
        Reporte(UBound(Reporte)).P_Cant = CantPres
        Reporte(UBound(Reporte)).P_PUnit = PrecioUPres
        Reporte(UBound(Reporte)).P_Total = TotalPres
        Reporte(UBound(Reporte)).Centro = BuscarDescCentroEmisor(BuscarCentroPadre(!CentroDeCosto))
        Reporte(UBound(Reporte)).SubCentro = BuscarDescCentro(!CentroDeCosto)
        
        Art = Mid(Art, 1, 30)
        Desvio = Total - TotalPres
        
        Item = Art + Space(38 - Len(Art) - Len(cant)) + cant
        Item = Item + Space(49 - Len(Item) - Len(Format(PrecioU, "0.00"))) + Format(PrecioU, "0.00")
        Item = Item + Space(60 - Len(Item) - Len(Format(Total, "0.00"))) + Format(Total, "0.00")
        Item = Item + Space(69 - Len(Item) - Len(CantPres)) + CantPres
        Item = Item + Space(84 - Len(Item) - Len(Format(PrecioUPres, "0.00"))) + Format(PrecioUPres, "0.00")
        Item = Item + Space(99 - Len(Item) - Len(Format(TotalPres, "0.00"))) + Format(TotalPres, "0.00")
        Item = Item + Space(110 - Len(Item) - Len(Format(Desvio, "0.00"))) + Format(Desvio, "0.00")
        
        If TotalPres <> 0 Then
            Desvio = Desvio / TotalPres
            Item = Item + Space(118 - Len(Item) - Len(Format(Desvio, "0.00%"))) + Format(Desvio, "0.00%")
        End If
      'agraga el Tv el nodo con la Cta - centro que componen el artículo
        TVListado.Nodes.Add CStr(!CentroDeCosto) + "C", tvwChild, , Item, 3
        TVListado.Nodes(TVListado.Nodes.Count).BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
        
        AlterColor = Not AlterColor

      .MoveNext
    Wend

      .Close
  End With
   
   For i = 1 To TVListado.Nodes.Count
        TVListado.Nodes(i).Expanded = True
   Next
    i = 1
    'borra los centros de contos padres que no tienen movimientos
   While Not TVListado.Nodes(i).Next Is Nothing
      If TVListado.Nodes(i).Parent Is Nothing And _
         TVListado.Nodes(i).Child Is Nothing Then
           TVListado.Nodes.Remove (i)
           i = 1
      Else
          i = TVListado.Nodes(i).Next.Index
      End If
        
   Wend
   'trata el último elemento del tv
    If TVListado.Nodes(i).Parent Is Nothing And _
       TVListado.Nodes(i).Child Is Nothing Then
         TVListado.Nodes.Remove (i)
    End If
    TVListado.Sorted = True
    ' ordena el tree view
    For i = 1 To TVListado.Nodes.Count
        TVListado.Nodes(i).Sorted = True
    Next
    i = 1
    Dim Nodo
    If TVListado.Nodes.Count > 0 Then
        While Not TVListado.Nodes(i).Child Is Nothing
            Set Nodo = TVListado.Nodes(i).Child
            AlterColor = True
            While Not Nodo Is Nothing
                If Nodo.Child Is Nothing Then
                'pone la alteración de colores
                    Nodo.BackColor = IIf(AlterColor, &HFFFFFF, &HE0E0E0)
                    AlterColor = Not AlterColor
                End If
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
    
    Nivel = TraerNivel("A015300")
    If Nivel = 2 Then
        CmbCentroDeCostoEmisor.ListIndex = 0
    Else
        Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
        CmbCentroDeCostoEmisor.Enabled = False
    End If

End Sub

Private Sub CalcularTotales()
  Dim i As Integer
  Dim Total As Double
  Dim TotalPres As Double
  
    For i = 1 To UBound(Centros)
        Total = Total + Centros(i).R_Total
        TotalPres = TotalPres + Centros(i).P_Total
    Next
        LBReal.Caption = "Total Real: $" + Format(Total, "0.00")
        LBPres.Caption = "Total Presupuestado: $" + Format(TotalPres, "0.00")
End Sub

