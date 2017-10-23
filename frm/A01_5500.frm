VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form A01_5500 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Certificaciones de Servicio Pendientes"
   ClientHeight    =   7485
   ClientLeft      =   4590
   ClientTop       =   2220
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   10320
      Begin VB.CommandButton CmdTraer 
         BackColor       =   &H80000003&
         Caption         =   "Traer"
         Height          =   315
         Left            =   2160
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   173
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalPeriodo 
         Height          =   330
         Left            =   1035
         TabIndex        =   5
         Top             =   165
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   54984707
         UpDown          =   -1  'True
         CurrentDate     =   38940
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
         Left            =   180
         TabIndex        =   6
         Top             =   233
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   400
      Left            =   3465
      TabIndex        =   2
      Top             =   6975
      Width           =   1635
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerra&r"
      Height          =   400
      Left            =   5277
      TabIndex        =   1
      Top             =   6975
      Width           =   1635
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   6255
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   11033
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
Attribute VB_Name = "A01_5500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub CrearEncabezado()

    LVListado.ColumnHeaders.Add , , "Centro de Costos Emisos", 2000
    LVListado.ColumnHeaders.Add , , "Proveedor", 2500
    LVListado.ColumnHeaders.Add , , "Orden Nº", 1100, 1
    LVListado.ColumnHeaders.Add , , "Fecha", 1100
    LVListado.ColumnHeaders.Add , , "Observaciones", 2500
    
    LVListado.ColumnHeaders.Add , , "Cuenta Constable", 2000
    LVListado.ColumnHeaders.Add , , "Sub-Centro de Costo", 2000
    LVListado.ColumnHeaders.Add , , "Precio", 1000, 1
    
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeConsulta
    ListA01_5500.Show
End Sub

Private Sub ConfImpresionDeConsulta()
  Dim i As Integer
  Dim RsListado As New ADODB.Recordset
    
    RsListado.Fields.Append "CentroDeCosto", adVarChar, 100
    RsListado.Fields.Append "Proveedor", adVarChar, 100
    RsListado.Fields.Append "NroOrden", adVarChar, 10
    RsListado.Fields.Append "Fecha", adVarChar, 10
    RsListado.Fields.Append "Observaciones", adVarChar, 500
    RsListado.Fields.Append "Cuenta", adVarChar, 100
    RsListado.Fields.Append "SubCentro", adVarChar, 100
    RsListado.Fields.Append "PrecioPactado", adDouble
    
    
    RsListado.Open
    i = 1
    While i <= LVListado.ListItems.Count
        RsListado.AddNew
      With LVListado.ListItems(i)
            RsListado!CentroDeCosto = .Text
            RsListado!Proveedor = .SubItems(1)
            RsListado!NroOrden = .SubItems(2)
            RsListado!Fecha = .SubItems(3)
            RsListado!Observaciones = .SubItems(4)
            RsListado!Cuenta = .SubItems(5)
            RsListado!SubCentro = .SubItems(6)
            RsListado!PrecioPactado = ValN(.SubItems(7))
      End With
        i = i + 1
    Wend
    RsListado.MoveFirst
    
    ListA01_5500.DataControl1.Recordset = RsListado
    ListA01_5500.Zoom = -1
End Sub

Private Sub CmdTraer_Click()
       Call CargarLV
End Sub

Private Sub Form_Load()
   Call CrearEncabezado
   CalPeriodo.Value = Date
End Sub

Private Sub CargarLV()
Dim RsCargar As New ADODB.Recordset
Dim Sql As String
Dim CentroOrden As String
Dim i As Integer
    
    With RsCargar
        Sql = "SpOcConsultaCertificacionDeServiciosPendientes @Periodo= " & FechaSQL(CalPeriodo, "SQL")
        .Open Sql, Conec
        While Not .EOF
            i = i + 1
            LVListado.ListItems.Add

            If CentroOrden <> !C_Descripcion & !O_NumeroOrdenDeContratacion Then
               LVListado.ListItems(i).Text = !C_Descripcion
               LVListado.ListItems(i).SubItems(1) = BuscarDescProv(!O_CodigoProveedor)
               LVListado.ListItems(i).SubItems(2) = Format(!O_NumeroOrdenDeContratacion, "00000000")
               LVListado.ListItems(i).SubItems(3) = !O_Fecha
               LVListado.ListItems(i).SubItems(4) = !O_Observaciones
               
               CentroOrden = !C_Descripcion & !O_NumeroOrdenDeContratacion
            End If
               LVListado.ListItems(i).SubItems(6) = BuscarDescCentro(!O_CentroDeCosto)
               LVListado.ListItems(i).SubItems(5) = BuscarDescCta(!O_CuentaContable)
            LVListado.ListItems(i).SubItems(7) = Format(!O_PrecioPactado, "0.00")
            .MoveNext
        Wend
    End With
End Sub

