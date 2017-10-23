VERSION 5.00
Begin VB.Form FrmImpPresFinanciero 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Egresos"
      Height          =   870
      Left            =   45
      TabIndex        =   11
      Top             =   630
      Width           =   5685
      Begin VB.CheckBox ChkFinanciero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fianciero"
         Height          =   240
         Left            =   3555
         TabIndex        =   5
         Top             =   225
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox ChkContable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contable"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox ChkPres 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Presupuestado"
         Height          =   240
         Left            =   1710
         TabIndex        =   4
         Top             =   225
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ChkDesvioContablePres 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desvio Contable/Presupuestado"
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   540
         Value           =   1  'Checked
         Width           =   2670
      End
      Begin VB.CheckBox ChkDesvioFinancieroPres 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desvio Financiero/Presupuestado"
         Height          =   240
         Left            =   2880
         TabIndex        =   7
         Top             =   540
         Value           =   1  'Checked
         Width           =   2760
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingresos"
      Height          =   555
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   5685
      Begin VB.CheckBox ChkReal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Real"
         Height          =   240
         Left            =   1485
         TabIndex        =   1
         Top             =   225
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox ChkProyeccion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proyección"
         Height          =   240
         Left            =   135
         TabIndex        =   0
         Top             =   225
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox ChkDesvioProyecciónReal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desvio Proyección/Real"
         Height          =   240
         Left            =   2790
         TabIndex        =   2
         Top             =   225
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   2978
      TabIndex        =   9
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   350
      Left            =   1628
      TabIndex        =   8
      Top             =   1575
      Width           =   1185
   End
End
Attribute VB_Name = "FrmImpPresFinanciero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Retorno As Opciones
Public Proyeccion As Boolean
Public Real As Boolean
Public DesvioProyeccionReal As Boolean
Public Contable As Boolean
Public Pres As Boolean
Public Financiero As Boolean
Public DesvioContablePres As Boolean
Public DesvioFinancieroPres As Boolean

Private Sub CmdCerrar_Click()
    Retorno = vbCerrar
    Unload Me
End Sub


Private Sub CmdImprimir_Click()
    Retorno = vbimprimir
    Proyeccion = ChkProyeccion.Value = 1
    Real = ChkReal.Value = 1
    DesvioProyeccionReal = ChkDesvioProyecciónReal.Value = 1
    Contable = ChkContable.Value = 1
    Pres = ChkPres.Value = 1
    Financiero = ChkFinanciero.Value = 1
    DesvioContablePres = chkDesvioContablePres.Value = 1
    DesvioFinancieroPres = ChkDesvioFinancieroPres.Value = 1
    Unload Me
End Sub

Private Sub CmdNuevo_Click()
    Retorno = vbNuevo
    Unload Me
End Sub

