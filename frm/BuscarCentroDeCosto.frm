VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarCentroDeCosto 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Centro de Costo"
   ClientHeight    =   7755
   ClientLeft      =   4590
   ClientTop       =   2220
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUbicar 
      Caption         =   "&Ubicar"
      Height          =   315
      Left            =   4005
      TabIndex        =   5
      Top             =   675
      Width           =   1230
   End
   Begin VB.TextBox TxtCodigo 
      Height          =   315
      Left            =   2025
      TabIndex        =   4
      Top             =   675
      Width           =   1680
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerra&r"
      Height          =   510
      Left            =   2070
      TabIndex        =   2
      Top             =   7020
      Width           =   1635
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   7110
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   5805
      Left            =   180
      TabIndex        =   1
      Top             =   1080
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   10239
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Centros de Costo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   3
      Top             =   720
      Width           =   1530
   End
   Begin VB.Label LBTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Centros de Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   5715
   End
End
Attribute VB_Name = "BuscarCentroDeCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub CmdUbicar_Click()
 Dim ItemX  As ListItem
    Set ItemX = LVListado.FindItem(TxtCodigo.Text)
   If Not (ItemX Is Nothing) Then
        ItemX.Selected = True
   Else
        MsgBox "El Artículo no existe", vbCritical
   End If
End Sub

Private Sub CrearEncabezado()
    LVListado.ColumnHeaders.Add , , "Artículo de Compras", LVListado.Width - 250
End Sub

Private Sub Form_Load()
   Call CrearEncabezado
   
End Sub

Private Sub LVListado_DblClick()
    If LVListado.ListItems.Count > 0 Then
         CentroDeCostoActual = VecCentroDeCosto(Me.LVListado.SelectedItem.Index)
    End If
    Unload Me
End Sub

Private Sub Timer1_Timer()
On Error GoTo errores
Dim i As Integer
    i = 1
    'With v.TbListado
    'cargo de a 25 renglones para no perder tanto tiempo y luego vuelve a arrancar el timer
    
        While i < UBound(VecCentroDeCosto) + 1
            LVListado.ListItems.Add
         'apareo LVListado con el vactor de modelos otros
            LVListado.ListItems(i).Text = VecCentroDeCosto(i).C_Descripcion

            i = i + 1
        Wend
  '  If .EOF Then
        'cuando terminé de calcular todos los renglones deshabilito el timer
  '      .Close
        Timer1.Enabled = False
   ' End If
  '  End With
      Call AlterColor(LVListado)

errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LvListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo errores
   ' Cuando se hace clic en un objeto ColumnHeader, el
   ' control ListView se ordena por los subelementos de
   ' esa columna.
   ' Establece el SortKey como el Index del ColumnHeader - 1
   ' Asigna a Sorted el valor True para ordenar la lista.
   LVListado.SortKey = ColumnHeader.Index - 1
   If LVListado.SortOrder = lvwAscending Then
        LVListado.SortOrder = lvwDescending
   Else
       LVListado.SortOrder = lvwAscending
   End If
   LVListado.Sorted = True
errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdUbicar_Click
    End If
End Sub
