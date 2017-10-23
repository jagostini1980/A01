VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MenuEmisionOrdenCompra 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sistema de Gestión Presupuestaria"
   ClientHeight    =   8865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10170
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
      Begin VB.Menu A011100 
         Caption         =   "Cuentas"
         Visible         =   0   'False
      End
      Begin VB.Menu A011B100 
         Caption         =   "Formas De Pago"
      End
      Begin VB.Menu A011200 
         Caption         =   "Proveedores"
         Visible         =   0   'False
      End
      Begin VB.Menu A011300 
         Caption         =   "Artículos de Compra"
      End
      Begin VB.Menu A011400 
         Caption         =   "Centros de Costos"
      End
      Begin VB.Menu A011410 
         Caption         =   "Cierre de período"
      End
      Begin VB.Menu A011420 
         Caption         =   "Relacionar artículos"
      End
      Begin VB.Menu A011500 
         Caption         =   "Clasificación Financiera de Cuentas"
      End
      Begin VB.Menu A011600 
         Caption         =   "Cuentas No Utilizadas en Consultas"
      End
      Begin VB.Menu A011900 
         Caption         =   "Cuentas no Utilizadas en Financiero"
      End
      Begin VB.Menu A011700 
         Caption         =   "Lugares de Entrega"
      End
      Begin VB.Menu A011800 
         Caption         =   "Artículos Taller"
      End
      Begin VB.Menu A011B200 
         Caption         =   "Agrupación Rubros Contables"
      End
      Begin VB.Menu A01Linea2 
         Caption         =   "-"
      End
      Begin VB.Menu A01Clave 
         Caption         =   "Cambio de Clave"
      End
   End
   Begin VB.Menu A014000 
      Caption         =   "Presupuestos"
      Begin VB.Menu A014100 
         Caption         =   "Ingreso de Presupuestos"
      End
      Begin VB.Menu A014200 
         Caption         =   "Aprobación de Presupuestos"
      End
      Begin VB.Menu A014300 
         Caption         =   "Comparación Presupuesto/Real"
         Visible         =   0   'False
      End
      Begin VB.Menu A012400 
         Caption         =   "Estado de Presupuestos"
      End
   End
   Begin VB.Menu A012000 
      Caption         =   "Órdenes"
      Begin VB.Menu A013600 
         Caption         =   "Requerimientos de Compra"
      End
      Begin VB.Menu A01Line31 
         Caption         =   "-"
      End
      Begin VB.Menu A012100 
         Caption         =   "Crear Órdenes de Compra"
      End
      Begin VB.Menu A016100 
         Caption         =   "Crear Ordenes de Contratación"
      End
      Begin VB.Menu A012300 
         Caption         =   "-"
      End
      Begin VB.Menu A013400 
         Caption         =   "Crear Ordenes de Compra Especial"
      End
      Begin VB.Menu A01Linea 
         Caption         =   "-"
      End
      Begin VB.Menu A013500 
         Caption         =   "Autorizar Órdenes"
      End
      Begin VB.Menu A013700 
         Caption         =   "Ver Requerimiento"
      End
   End
   Begin VB.Menu A013000 
      Caption         =   "Recepcion - Certificaciones"
      Begin VB.Menu A013200 
         Caption         =   "Crear Recepción de mercadería"
      End
      Begin VB.Menu A016200 
         Caption         =   "Crear Certificación de Servicios"
      End
      Begin VB.Menu A013300 
         Caption         =   "-"
      End
      Begin VB.Menu A014400 
         Caption         =   "Crear Certificación de Servicio Especiales"
      End
      Begin VB.Menu A014500 
         Caption         =   "Crear Centificación de Fondo Fijo"
      End
      Begin VB.Menu A014600 
         Caption         =   "Crear Centificación Turismo"
      End
   End
   Begin VB.Menu A015000 
      Caption         =   "Consultas"
      Begin VB.Menu A015100 
         Caption         =   "Por Cuenta Contable - Artículo"
      End
      Begin VB.Menu A015200 
         Caption         =   "Por Cuenta Contable - Sub Centro de Costo"
      End
      Begin VB.Menu A015300 
         Caption         =   "Por Centro de Costo"
         Visible         =   0   'False
      End
      Begin VB.Menu A015400 
         Caption         =   "Recepciones de Mercaderías Pendientes"
      End
      Begin VB.Menu A015500 
         Caption         =   "Certificaciones de Servicio Pendientes"
      End
      Begin VB.Menu A015600 
         Caption         =   "Desvio Presupuestado/Contable"
      End
      Begin VB.Menu A015700 
         Caption         =   "Desvio Presupuestado/Contable/SGP"
      End
      Begin VB.Menu A015800 
         Caption         =   "Evolución Presupuestado/Real"
      End
      Begin VB.Menu A015B300 
         Caption         =   "Evolución Movilidades"
      End
      Begin VB.Menu A015900 
         Caption         =   "Presupuesto Financiero"
      End
      Begin VB.Menu A015B200 
         Caption         =   "Rubros Contables"
      End
      Begin VB.Menu A015B900 
         Caption         =   "Rubros Contables Acumulado"
      End
      Begin VB.Menu A015B100 
         Caption         =   "Estado de Presupuesto"
      End
      Begin VB.Menu A015B400 
         Caption         =   "Totales por Empresa"
      End
      Begin VB.Menu A015B500 
         Caption         =   "Egresos Financieros por Empresa"
      End
      Begin VB.Menu A015B600 
         Caption         =   "Gastos Por Centro de Costo Emisor"
      End
      Begin VB.Menu A015B700 
         Caption         =   "Gastos por U de Negocio"
      End
      Begin VB.Menu A015B800 
         Caption         =   "Evolución Pres/Real por Cuenta"
      End
      Begin VB.Menu A015C100 
         Caption         =   "Evolución Pres/SGP por Cuenta"
      End
      Begin VB.Menu A015C200 
         Caption         =   "Comparativo MC Presupuestado - Real"
      End
   End
   Begin VB.Menu A017000 
      Caption         =   "Proveedores"
      Begin VB.Menu A017100 
         Caption         =   "Evaluar Proveerores"
      End
      Begin VB.Menu A017200 
         Caption         =   "Consulta de Proveedores Evaluados"
      End
      Begin VB.Menu A017300 
         Caption         =   "Estado de Comprobantes"
      End
      Begin VB.Menu A017400 
         Caption         =   "OP Pendientes de Devolución"
      End
   End
   Begin VB.Menu A01Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu A01Sistema 
         Caption         =   "Sistema"
      End
      Begin VB.Menu A01Forzar 
         Caption         =   "Forzar Actualización"
      End
      Begin VB.Menu A01Actualizador 
         Caption         =   "Actualizar Actualizador"
      End
   End
   Begin VB.Menu A019999 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MenuEmisionOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub A011B100_Click()
    A01_1B100.Show
End Sub

Private Sub A011410_Click()
    A01_1410.Show
End Sub

Private Sub A011420_Click()
    A01_1420.Show
End Sub

Private Sub A011500_Click()
    A01_1500.Show
End Sub

Private Sub A011600_Click()
    A01_1600.Show
End Sub

Private Sub A011700_Click()
    A01_1700.Show
End Sub

Private Sub A011800_Click()
    Man_1700.Show
End Sub

Private Sub A011900_Click()
    A01_1900.Show
End Sub

Private Sub A011B200_Click()
    A01_1B200.Show
End Sub

Private Sub A012400_Click()
    A01_2400.Show
End Sub

Private Sub A013200_Click()
    ReDim Articulos(0)
    A01_3200.Show
End Sub

Private Sub A013400_Click()
    A01_3400.Show
End Sub

Private Sub A013500_Click()
    A01_3500.Show
End Sub

Private Sub A013600_Click()
    A01_3600.Show
End Sub

Private Sub A013700_Click()
    A01_Requerimientos.Show
End Sub

Private Sub A014100_Click()
    A01_4100.Show
End Sub

Private Sub A014200_Click()
    A01_4200.Show vbModal
End Sub

Private Sub A014400_Click()
    A01_4400.Show
End Sub

Private Sub A014500_Click()
    A01_4500.Show
End Sub

Private Sub A014600_Click()
  MousePointer = vbHourglass
    A01_4600.Show
  MousePointer = vbNormal
End Sub

Private Sub A015100_Click()
    A01_5100.Show
End Sub

Private Sub A015200_Click()
    A01_5200.Show
End Sub

Private Sub A015300_Click()
    A01_5300.Show
End Sub

Private Sub A015400_Click()
    A01_5400.Show
End Sub

Private Sub A015500_Click()
    A01_5500.Show
End Sub

Private Sub A015600_Click()
    A01_5600.Show
End Sub

Private Sub A015700_Click()
    A01_5700.Show
End Sub

Private Sub A015800_Click()
    A01_5800.Show
End Sub

Private Sub A015900_Click()
    A01_5900.Show
End Sub

Private Sub A015B100_Click()
    A01_5B100.Show
End Sub

Private Sub A015B200_Click()
    A01_5B200.Show
End Sub

Private Sub A015B300_Click()
    A01_5B300.Show
End Sub

Private Sub A015B400_Click()
    A01_5B400.Show
End Sub

Private Sub A015B500_Click()
    A01_5B500.Show
End Sub

Private Sub A015B600_Click()
    A01_5B600.Show
End Sub

Private Sub A015B700_Click()
    A01_5B700.Show
End Sub

Private Sub A015B800_Click()
    A01_5B800.Show
End Sub

Private Sub A015B900_Click()
    A01_5B900.Show
End Sub

Private Sub A015C100_Click()
    A01_5C100.Show
End Sub

Private Sub A015C200_Click()
    A01_5C200.Show
End Sub

Private Sub A016100_Click()
    A01_6100.Show
End Sub

Private Sub A016200_Click()
    A01_6200.Show
End Sub

Private Sub A017100_Click()
    A01_7100.Show
End Sub

Private Sub A017200_Click()
    A01_7200.Show
End Sub

Private Sub A017300_Click()
    A01_7300.Show
End Sub

Private Sub A017400_Click()
    A01_7400.Show
End Sub

Private Sub A01Actualizador_Click()
On Error GoTo ErrorFtp
  Dim mFTP As cFTP
  Set mFTP = New cFTP
  
  mFTP.SetModeActive
  mFTP.SetTransferBinary
  
  MousePointer = vbHourglass
  
    If FileExist(App.Path & "\Actualizador.zip") Then
        Call Kill(App.Path & "\Actualizador.zip")
    End If
    
     If mFTP.OpenConnection("svrppack.dyndns.org", "", "") Then
        Call mFTP.SetFTPDirectory("A01")
        If mFTP.FTPDownloadFile(App.Path & "\Actualizador.zip", "Actualizador.zip") Then
            Call UnZip(App.Path & "\Actualizador.zip", App.Path)
        Else
           MsgBox "Error al Descargar el Actualizados", vbCritical, "Error de Actualizacion"
        End If
     Else
        MsgBox "Error al intentar conectarse al Servidor", vbCritical, "Error de Actualizacion"
     End If

      MsgBox "El Actualizador se actualizó correctamente", vbInformation
ErrorFtp:
    Call ManipularError(Err.Number, Err.Description)
    MousePointer = vbNormal
End Sub

Private Sub A01Clave_Click()
    CambioDeClave.Show vbModal
End Sub

Private Sub A01Forzar_Click()
On Error Resume Next
    Call Kill(App.Path & "\A01.zip")
    MsgBox "El sistema se cerrará y se actualizará en el proximo ingreso", vbInformation
    End
End Sub

Private Sub A01Sistema_Click()
    Call FrmModificacionesParaUsuario.CargarLV("A01")
    Dim Version As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
    FrmModificacionesParaUsuario.TxtVercionActual.Text = Trim(Mid(Version, InStr(1, Version, ":") + 1, Len(Version)))

    FrmModificacionesParaUsuario.Show
End Sub

Private Sub MDIForm_Activate()
    If Tag = "R" And A013700.Enabled Then
      A01_Requerimientos.Show
      A01_Requerimientos.SetFocus
      Tag = ""
    End If
End Sub

Private Sub MDIForm_Load()
    Call HabilitarMenus(Usuario, Me)
    Caption = Caption & " - Usuario: " & Usuario & " - Versión: " & App.Major & "." & App.Minor & "." & App.Revision
    Tag = "R"
End Sub

Private Sub A011100_Click()
    Cuentas.Show
End Sub

Private Sub A011300_Click()
    A01_1300.Show
End Sub

Private Sub A011400_Click()
    A01_1400.Show
End Sub

Private Sub A012100_Click()
    ReDim VecRequerimientoCompra(0)
    A01_2100.Show
End Sub

Private Sub A019999_Click()
    End
End Sub
