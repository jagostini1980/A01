VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_2120 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de artículos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCant 
      Height          =   280
      Left            =   180
      TabIndex        =   3
      Top             =   6795
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdCrearRecepcion 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2482
      TabIndex        =   2
      Top             =   6840
      Width           =   1320
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3922
      TabIndex        =   1
      Top             =   6840
      Width           =   1320
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   6690
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   7665
      _ExtentX        =   13520
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
End
Attribute VB_Name = "A01_2120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CodTaller As Long

Private Sub CmdCrearRecepcion_Click()
    Dim i As Integer
    ReDim Precios(LvListado.ListItems.Count)
    For i = 1 To UBound(Precios)
        Precios(i) = Val(Replace(LvListado.ListItems(i).SubItems(2), ",", "."))
    Next
    Unload Me
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CreaEncabezado
End Sub

Private Sub CreaEncabezado()
    LvListado.ColumnHeaders.Add , , "Artículo", LvListado.Width - 3900
    LvListado.ColumnHeaders.Add , , "Cant. a Comprar", 1000, 1
    LvListado.ColumnHeaders.Add , , "Precio U. sin IVA", 1300, 1
    LvListado.ColumnHeaders.Add , , "Total", 1300, 1
    LvListado.ColumnHeaders.Add , , "A_Codigo", 0
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
   With LvListado
      .SelectedItem.EnsureVisible
      TxtCant.Text = Replace(.ListItems(Item.Index).SubItems(2), ",", ".")
      TxtCant.Move .ColumnHeaders.Item(3).Left + 60, .SelectedItem.Top + 45, .ColumnHeaders.Item(3).Width, 200
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
    If Val(TxtCant) > 0 Then
       LvListado.SelectedItem.SubItems(2) = Format(Val(TxtCant), "0.00")
       LvListado.SelectedItem.SubItems(3) = Format(ValN(LvListado.SelectedItem.SubItems(2)) * ValN(LvListado.SelectedItem.SubItems(1)), "0.00")
       If TxtCant.Visible Then
          TxtCant.Visible = False
       End If
       LvListado.SetFocus
       
    Else
        MsgBox "El Precio debe ser Mayor a 0", vbInformation
    End If
    Dim i As Integer
    CmdCrearRecepcion.Enabled = True
    
    For i = 1 To LvListado.ListItems.Count
        If ValN(LvListado.ListItems(i).SubItems(2)) = 0 Then
            CmdCrearRecepcion.Enabled = False
            Exit For
        End If
    Next
End Sub

Private Sub TxtCant_Validate(Cancel As Boolean)
    Call ValidateTxtCant
End Sub
