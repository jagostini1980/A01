VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmModificacionesParaUsuario 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificaciones del sistema"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   4492
      TabIndex        =   2
      Top             =   7110
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   10140
      Begin VB.TextBox TxtUltimaVersion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4410
         TabIndex        =   8
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox TxtVercionActual 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1485
         TabIndex        =   6
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox TxtNombreSist 
         Enabled         =   0   'False
         Height          =   285
         Left            =   945
         TabIndex        =   5
         Top             =   180
         Width           =   4965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Última Versión:"
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
         Left            =   3060
         TabIndex        =   7
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión actual:"
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
         Left            =   135
         TabIndex        =   4
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label LbSistema 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema:"
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
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView LvObs 
      Height          =   6000
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   10583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmModificacionesParaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerra_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Call CrearEmcabezado
End Sub

Private Sub CargarLV(Sistema As String)
   Dim Sql As String
   Dim RsCargar As New ADODB.Recordset
   Dim i As Integer
   
   RsCargar.CursorLocation = adUseClient
   RsCargar.CursorType = adOpenKeyset
   
   Sql = "SpSisModificacionesRealizadasSistema @Sistema ='" & Sistema & "'"
   
   RsCargar.Open Sql, Conec
   With RsCargar
     'limpia el LV
      LvObs.ListItems.Clear

      If .RecordCount > 0 Then
        For i = 1 To .RecordCount
           
           LvObs.ListItems.Add
           LvObs.ListItems(i).Text = VerificarNulo(!M_Fecha)
           LvObs.ListItems(i).SubItems(1) = VerificarNulo(!M_VersionSistema)
           LvObs.ListItems(i).SubItems(2) = VerificarNulo(!M_Observacion)
                             
           .MoveNext
         Next
      End If
      .Close
       
       Sql = "SpSisSistemasTraer @Sistema ='" & Sistema & "'"
      .Open Sql, Conec
       TxtNombreSist = VerificarNulo(!S_NombreCompleto)
       TxtUltimaVersion.Text = VerificarNulo(!UltimaVersion)
      .Close
   End With
   
   Set RsCargar = Nothing
End Sub

Private Sub CrearEmcabezado()
    LvObs.ColumnHeaders.Add , , "Frecha", 1300
    LvObs.ColumnHeaders.Add , , "Versión", 900
    LvObs.ColumnHeaders.Add , , "Observaciones", LvObs.Width - 2500
End Sub

