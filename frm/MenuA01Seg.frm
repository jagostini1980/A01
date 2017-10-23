VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MenuA01Seg 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sistema de Gestión De Pólizas"
   ClientHeight    =   8865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10230
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cuadros 
      Left            =   5130
      Top             =   3825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu A011000 
      Caption         =   "Actualizaciones"
      Begin VB.Menu A011700 
         Caption         =   "ABM Unidades"
      End
      Begin VB.Menu A011800 
         Caption         =   "Páramatros Seguro"
      End
      Begin VB.Menu A01Linea 
         Caption         =   "-"
      End
      Begin VB.Menu A01Clave 
         Caption         =   "Cambio de Clave"
      End
   End
   Begin VB.Menu A017000 
      Caption         =   "Movimientos"
      Begin VB.Menu A017100 
         Caption         =   "Ingreso de Pólizas"
      End
      Begin VB.Menu A017200 
         Caption         =   "Generación de Pagos"
      End
   End
   Begin VB.Menu A01Sistema 
      Caption         =   "Sistema"
   End
   Begin VB.Menu A019999 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MenuA01Seg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub A011700_Click()
    Man_1900.Show
End Sub

Private Sub A011800_Click()
    A01_1800.Show
End Sub

Private Sub A017100_Click()
    A01_7100.Show
End Sub

Private Sub A017200_Click()
    A01_7200.Show
End Sub

Private Sub A01Clave_Click()
    CambioDeClave.Show vbModal
End Sub

Private Sub A01Sistema_Click()
    Call FrmModificacionesParaUsuario.CargarLV("A01Seg")
    Dim Version As String
    Version = IngresoSeg.LBVersion.Caption
    FrmModificacionesParaUsuario.TxtVercionActual.Text = Trim(Mid(Version, InStr(1, Version, ":") + 1, Len(Version)))

    FrmModificacionesParaUsuario.Show
End Sub

Private Sub MDIForm_Load()
    Call HabilitarMenus(Usuario, Conec.ConnectionString, Me)
    Caption = Caption & " - Usuario: " & Usuario & " - Versión: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub A019999_Click()
    End
End Sub
