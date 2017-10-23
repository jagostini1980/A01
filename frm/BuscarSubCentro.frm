VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarSubCentro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Sub-Centros de costo"
   ClientHeight    =   6825
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2685
      TabIndex        =   2
      Top             =   6300
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   6165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BuscarSubCentro.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BuscarSubCentro.frx":27B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BuscarSubCentro.frx":4F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BuscarSubCentro.frx":527E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Centro de costo"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   6330
      Begin MSComctlLib.TreeView TvCentros 
         Height          =   5775
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   10186
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "BuscarSubCentro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodigoSubCentro As String

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Call CargarTvCentros
End Sub

Private Sub CargarTvCentros()
 Dim Nodo As Node
  Dim i As Integer
 TvCentros.Nodes.Clear
     
     'carga el primer nivel
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        With VecCentroDeCostoEmisor(i)
           Set Nodo = TvCentros.Nodes.Add(, , .C_Jerarquia & "C", .C_Descripcion, 1)
        End With
    Next
    
    For i = 1 To UBound(VecCentroDeCostoNivel2)
        'carga el segundo nivel
        With VecCentroDeCostoNivel2(i)
             Set Nodo = TvCentros.Nodes.Add(.C_Padre & "C", tvwChild, .C_Jerarquia + "C", .C_Descripcion, 4)
        End With
    Next

    For i = 1 To UBound(VecCentroDeCosto)
        'carga el segundo nivel
        With VecCentroDeCosto(i)
             Set Nodo = TvCentros.Nodes.Add(.C_Padre + "C", tvwChild, .C_Codigo + "C", .C_Descripcion, 2)
        End With
    Next
    
    'For i = 1 To TvCentros.Nodes.Count
    '    TvCentros.Nodes(i).Expanded = True
    'Next
    TvCentros.Nodes(1).Selected = True
    End Sub

'Private Sub TvCentros_Collapse(ByVal Node As MSComctlLib.Node)
'    Node.Expanded = True
'End Sub

Private Sub TvCentros_DblClick()
    If TvCentros.SelectedItem.Child Is Nothing Then
        CodigoSubCentro = Mid(TvCentros.SelectedItem.Key, 1, 4)
        Unload Me
    'Else
        'MsgBox "Debe Seleccionar un Sub-Centro de Costo", vbInformation
    End If
End Sub

