VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_2110 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTraer 
      Caption         =   "Traer"
      Height          =   315
      Left            =   5490
      TabIndex        =   8
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox TxtCant 
      Height          =   280
      Left            =   6885
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdCrearRecepcion 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3990
      TabIndex        =   2
      Top             =   6840
      Width           =   1320
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5430
      TabIndex        =   1
      Top             =   6840
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6285
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   11086
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
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker CalFechaDesde 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   90
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Format          =   111804417
      CurrentDate     =   38940
   End
   Begin MSComCtl2.DTPicker CalFechaHasta 
      Height          =   315
      Left            =   4140
      TabIndex        =   5
      Top             =   90
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Format          =   111804417
      CurrentDate     =   38940
   End
   Begin VB.Label LBFechaHasta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Hasta:"
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
      Left            =   2970
      TabIndex        =   7
      Top             =   135
      Width           =   1155
   End
   Begin VB.Label LbFechaDesde 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Desde:"
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
      Left            =   315
      TabIndex        =   6
      Top             =   135
      Width           =   1200
   End
End
Attribute VB_Name = "A01_2110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CodTaller As Long
Private ItemsChequeados As Integer

Private Sub CmdCrearRecepcion_Click()
  Dim i As Integer
  Dim R As Integer
    
    ReDim Requerimientos(0)
    
    For i = 1 To LvListado.ListItems.Count
      If LvListado.ListItems(i).Checked Then
        ReDim Preserve Requerimientos(UBound(Requerimientos) + 1)
        R = UBound(Requerimientos)
        Requerimientos(R).CodArticulo = ValN(LvListado.ListItems(i).SubItems(3))
        Requerimientos(R).DescArticulo = LvListado.ListItems(i).SubItems(4)
            'Cant Pedida                              Cant. Pendiente
        If ValN(LvListado.ListItems(i).SubItems(6)) > ValN(LvListado.ListItems(i).SubItems(7)) Then
             Requerimientos(R).Cantidad = ValN(LvListado.ListItems(i).SubItems(7))
             Requerimientos(R).CantidadExtra = ValN(LvListado.ListItems(i).SubItems(6)) - ValN(LvListado.ListItems(i).SubItems(7))
        Else
             Requerimientos(R).Cantidad = ValN(LvListado.ListItems(i).SubItems(6))
             Requerimientos(R).CantidadExtra = 0
        End If
        
        Requerimientos(R).Taller = ValN(LvListado.ListItems(i).SubItems(8))
        Requerimientos(R).Numero = ValN(LvListado.ListItems(i).SubItems(1))
        Requerimientos(R).Marca = LvListado.ListItems(i).SubItems(5)
      End If
    Next
   
 'cierra el formulario actual antes de abrir el de creacion de recepciones
    Unload Me
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTraer_Click()
    Call TraerArticulosPendientes
End Sub

Private Sub Form_Load()
    Call CreaEncabezado
    CalFechaDesde.Value = DateAdd("M", -1, Date)
    CalFechaHasta.Value = Date
    Call TraerArticulosPendientes
End Sub

Private Sub TraerArticulosPendientes()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    
    Sql = "SpOcOrdenesDeCompraRequerimientosPendientes @FechaDesde =" & FechaSQL(CalFechaDesde, "SQL") & _
                                                   ",  @FechaHasta =" & FechaSQL(CalFechaHasta, "SQL")
          
    LvListado.ListItems.Clear
    LvListado.Sorted = False
    With RsCargar
        .Open Sql, Conec
        i = 1
       While Not .EOF
          LvListado.ListItems.Add
          LvListado.ListItems(i).Text = BuscarDescTaller(!R_Taller)
          LvListado.ListItems(i).SubItems(1) = !R_Numero
          LvListado.ListItems(i).SubItems(2) = !R_Fecha
          LvListado.ListItems(i).SubItems(3) = !R_Articulo
          LvListado.ListItems(i).SubItems(4) = Trim(!A_Descripcion) 'BuscarDescArt(!R_Articulo, Trim(!C_TablaArticulos))
          LvListado.ListItems(i).SubItems(5) = !R_Marca
          LvListado.ListItems(i).SubItems(6) = Format(!R_RestaOC, "0.00")
          LvListado.ListItems(i).SubItems(7) = Format(!R_RestaOC, "0.00")
          
          LvListado.ListItems(i).SubItems(8) = !R_Taller
          i = i + 1
          .MoveNext
       Wend
    End With
    LvListado.Sorted = True
End Sub

Private Sub CreaEncabezado()
    LvListado.ColumnHeaders.Add , , "Taller", 1500
    LvListado.ColumnHeaders.Add , , "Requerimiento Nº", 1000, 1
    LvListado.ColumnHeaders.Add , , "Fecha", 1100
    LvListado.ColumnHeaders.Add , , "Codigo", 800
    LvListado.ColumnHeaders.Add , , "Artículo", LvListado.Width - 7800
    LvListado.ColumnHeaders.Add , , "Marca", 1000
    LvListado.ColumnHeaders.Add , , "Cant. a Comprar", 1000, 1
    LvListado.ColumnHeaders.Add , , "Cant. Pendiente", 1300, 1

    LvListado.ColumnHeaders.Add , , "CodTaller", 0
    LvListado.ColumnHeaders.Add , , "Prioridad", 1000
End Sub

Private Sub LVListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvListado.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LvListado_ItemCheck(ByVal Item As MSComctlLib.ListItem)
      If Item.Checked Then
         ItemsChequeados = ItemsChequeados + 1
      Else
         ItemsChequeados = ItemsChequeados - 1
      End If
      
      If ItemsChequeados > 0 Then
         CmdCrearRecepcion.Enabled = True
      Else
         CmdCrearRecepcion.Enabled = False
      End If
    
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
   With LvListado
      .SelectedItem.EnsureVisible
      TxtCant.Text = Replace(.ListItems(Item.Index).SubItems(6), ",", ".")
      TxtCant.Move .ColumnHeaders.Item(7).Left + 30, .SelectedItem.Top + 450, .ColumnHeaders.Item(6).Width, 200
      TxtCant.Visible = True
      TxtCant.SetFocus
   End With
End Sub

Private Sub TxtCant_GotFocus()
    TxtCant.SelStart = 0
    TxtCant.SelLength = Len(TxtCant)
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LvListado.SetFocus
        Call ValidateTxtCant
    Else
        Call TxtNumerico(TxtCant, KeyAscii)
        LvListado.SelectedItem.Checked = True
        Call LvListado_ItemCheck(LvListado.SelectedItem)
    End If
End Sub

Private Sub TxtCant_KeyUp(KeyCode As Integer, Shift As Integer)
  With LvListado
    If KeyCode = 38 And .SelectedItem.Index > 1 Then
        Call ValidateTxtCant
       .ListItems(.SelectedItem.Index - 1).Selected = True
        Call LvListado_ItemClick(.SelectedItem)
    End If
    If KeyCode = 40 And .SelectedItem.Index < .ListItems.Count Then
        Call ValidateTxtCant
       .ListItems(.SelectedItem.Index + 1).Selected = True
        Call LvListado_ItemClick(.SelectedItem)
    End If
  End With

End Sub

Private Sub ValidateTxtCant()
    If Val(TxtCant) = 0 Then
        MsgBox "La Cantidad debe ser mayor a 0", vbInformation
        Exit Sub
    End If
    'If Val(TxtCant) <= ValN(LvListado.SelectedItem.SubItems(6)) Then
       LvListado.SelectedItem.SubItems(6) = Format(Val(TxtCant), "0.00")
    '   If TxtCant.Visible Then
    '      TxtCant.Visible = False
    '   End If
    'Else
    '    MsgBox "La Cantidad supera a requisito", vbInformation
    'End If
End Sub

Private Sub TxtCant_Validate(Cancel As Boolean)
    Call ValidateTxtCant
End Sub
