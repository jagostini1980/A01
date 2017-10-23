VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form A01_5810 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Periodos"
   ClientHeight    =   5070
   ClientLeft      =   3420
   ClientTop       =   3405
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4590
      Width           =   1200
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   4485
      Left            =   83
      TabIndex        =   0
      Top             =   45
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   7911
      View            =   3
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
End
Attribute VB_Name = "A01_5810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PeriodoInicio As Date
Public PeriodoFin As Date

Private Sub CmdAceptar_Click()
    Dim i As Integer
    Dim Contar As Integer
    
    For i = 1 To LvListado.ListItems.Count
        If LvListado.ListItems(i).Checked Then
            Contar = Contar + 1
            Select Case Contar
            Case 1
                A01_5800.Per1 = i - 1
            Case 2
                A01_5800.Per2 = i - 1
            Case 3
                A01_5800.Per3 = i - 1
            End Select
        End If
    Next
    If Contar = 0 Then
        MsgBox "Debe Seleccionar algún período", vbInformation
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call CrearEmcabezados
    For i = 0 To DateDiff("M", PeriodoInicio, PeriodoFin)
        LvListado.ListItems.Add
        LvListado.ListItems(i + 1).Text = Format(DateAdd("M", i, PeriodoInicio), "MMMM/yyyy")
    Next
End Sub

Private Sub CrearEmcabezados()
   LvListado.ColumnHeaders.Add , , "Períodos", LvListado.Width - 250
End Sub

Private Sub LvListado_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim Contar As Integer
    
    For i = 1 To LvListado.ListItems.Count
        If LvListado.ListItems(i).Checked Then
            Contar = Contar + 1
        End If
    Next
    If Contar > 3 Then
        MsgBox "Solo Puede seleccionar 3 Períodos", vbInformation
        Item.Checked = False
    End If
End Sub
