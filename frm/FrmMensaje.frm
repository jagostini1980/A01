VERSION 5.00
Begin VB.Form FrmMensaje 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkImp2 
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   1380
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.CheckBox ChkImp1 
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   1150
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   3982
      TabIndex        =   4
      Top             =   1620
      Width           =   1185
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   2707
      TabIndex        =   3
      Top             =   1620
      Width           =   1185
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "&Exportar a PDF"
      Height          =   350
      Left            =   1342
      TabIndex        =   2
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   350
      Left            =   67
      TabIndex        =   1
      Top             =   1620
      Width           =   1185
   End
   Begin VB.Label LbMensaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   870
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   5100
   End
End
Attribute VB_Name = "FrmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Retorno As Opciones
Public Imp1 As Boolean
Public Imp2 As Boolean

Private Sub CmdCerrar_Click()
    Retorno = vbCerrar
    Unload Me
End Sub

Private Sub CmdExportar_Click()
    Retorno = vbExportesPDF
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Retorno = vbimprimir
    Imp1 = ChkImp1.Value = 1
    Imp2 = ChkImp2.Value = 1
    Unload Me
End Sub

Private Sub CmdNuevo_Click()
    Retorno = vbNuevo
    Unload Me
End Sub
