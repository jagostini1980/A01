VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarProveedores 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "El Pulqui"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por: "
      Height          =   615
      Left            =   12555
      TabIndex        =   13
      Top             =   6660
      Width           =   2580
      Begin VB.OptionButton OptOrden 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.U.I.T."
         Height          =   255
         Index           =   2
         Left            =   1515
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton OptOrden 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alfabetico"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerra&r"
      Height          =   510
      Left            =   7208
      TabIndex        =   4
      Top             =   6870
      Width           =   1635
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   510
      Left            =   5408
      TabIndex        =   3
      Top             =   6885
      Width           =   1635
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   270
      Top             =   6660
   End
   Begin MSComctlLib.ListView LVListado 
      Height          =   5805
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   14745
      _ExtentX        =   26009
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
   Begin VB.Label CUIT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   12
      Top             =   7110
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Direccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1485
      TabIndex        =   11
      Top             =   6705
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Localidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1485
      TabIndex        =   10
      Top             =   7110
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Telefono 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3735
      TabIndex        =   9
      Top             =   7110
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Reducida 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      TabIndex        =   8
      Top             =   6705
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label CodigoPostal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3735
      TabIndex        =   7
      Top             =   6705
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Descripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   7110
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6705
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LBTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedores"
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
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   13695
   End
   Begin VB.Label LbConexion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "Lbconexion"
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   315
      TabIndex        =   2
      Top             =   7035
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "BuscarProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Listado
    Rubro As Integer
    UnidadDeMedida As String
    Marca As String
End Type

Private Type VariablesGlobales
    Servidor As String
    Env As rdoEnvironment
    db As rdoConnection
    TbListado As rdoResultset
    VecBuscar() As Integer
    VecListado() As Listado
    Cargar As Boolean
End Type

Private Type VariablesImpresion
    TamanioLetra As Single
    SeparacionConceptos As Integer
    CantidadLetras As Integer
    ancho As Integer
    VecPosiciones() As Integer        'es la posicion donde se imprime cada encabezado
End Type


Const VgNumero = "#0.00" 'esta constante es el formato de los numeros
Dim v As VariablesGlobales
Dim Vi As VariablesImpresion

Public Sub CargarParametros(Conexion As String)
    LbConexion.Caption = Conexion
    'OptOrden(1).Value = True
'    CargarListado
End Sub


Private Sub CmdCerrar_Click()
    Hide
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Errores
    TeclaPresionada ActiveControl, KeyAscii
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LbConexion_Change()
On Error GoTo Errores
    Dim NombreConexion As String
    Dim Usuario As String
    Dim Clave As String
    Dim Opciones As String
    MousePointer = vbHourglass
    Set v.Env = rdoEnvironments(0)
    NombreConexion = BuscarString(LbConexion.Caption, "dsn=")
    Usuario = BuscarString(LbConexion.Caption, "UID=")
    Clave = BuscarString(LbConexion.Caption, "PWD=")
    If Trim(Clave) = "" Then
        Opciones = "UID=" & Usuario
    Else
        Opciones = "UID=" & Usuario & ";PWD=" & Clave
    End If
    
    v.Servidor = BuscarString(LbConexion.Caption, "servidor=")
    Set v.db = v.Env.OpenConnection(NombreConexion, rdDriverNoPrompt, False, Opciones)
    InicializarTodo
    MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub CargarListado()
On Error GoTo Errores
Dim sSQL As String
        MousePointer = vbHourglass
        If OptOrden(1).Value = False And OptOrden(2).Value = False Then
            OptOrden(1).Value = True
        Else
            sSQL = "SpTAProveedores"
            Set v.TbListado = v.db.OpenResultset(sSQL)
            LvListado.ListItems.Clear
            LvListado.Sorted = False
            ' esto activa el timer para empezar a cargar los renglones
            
            Timer1.Enabled = True
            MousePointer = vbNormal
        End If

Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub InicializarTodo()
On Error GoTo Errores
   ' InicializarTags
    CargarEncabezados
    Vi.TamanioLetra = 7
    Vi.CantidadLetras = 10
    Vi.SeparacionConceptos = 10
    CargarListado
    
    ReDim Vi.VecPosiciones(8)
    Vi.VecPosiciones(0) = 1    ' Descripcion
    Vi.VecPosiciones(1) = 42   ' Codigo
    Vi.VecPosiciones(2) = 47   ' Reducida
    Vi.VecPosiciones(3) = 64   ' Direccion
    Vi.VecPosiciones(4) = 106  ' Localidad
    Vi.VecPosiciones(5) = 128   ' Codigo Postal
    Vi.VecPosiciones(6) = 138  ' Telefono
    Vi.VecPosiciones(7) = 180  ' CUIT
    Vi.ancho = 200
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub CargarEncabezados()
On Error GoTo Errores
Dim TamanioTotal As Integer
Dim TamanioColumna As Integer
    'tamaniocolumna es el tamaño que va a tener la columna indicada
    
    TamanioColumna = 3000
    TamanioTotal = TamanioColumna
    LvListado.ColumnHeaders.Add 1, "Descripción", "Descripción", TamanioColumna
    
    TamanioColumna = 500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 2, "Cód.", "Cód.", TamanioColumna
    
    TamanioColumna = 1500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 3, "Reducida", "Reducida", TamanioColumna
    
    TamanioColumna = 3000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 4, "Dirección", "Dirección", TamanioColumna
    
    TamanioColumna = 1500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 5, "Localidad", "Localidad", TamanioColumna
    
    TamanioColumna = 1000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 6, "C.P.", "C.P.", TamanioColumna
    
    TamanioColumna = 3000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 7, "Teléfono", "Teléfono", TamanioColumna
    
    'el tamaño de la última columna se actualiza automáticamente para acomodarse al listview
    LvListado.ColumnHeaders.Add 8, "CUIT", "CUIT", LvListado.Width - TamanioTotal - 100

Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub
Private Sub CargarEncabezadosXCuit()
On Error GoTo Errores
Dim TamanioTotal As Integer
Dim TamanioColumna As Integer
    'tamaniocolumna es el tamaño que va a tener la columna indicada
    
    TamanioColumna = 1200
    TamanioTotal = TamanioColumna
    LvListado.ColumnHeaders.Add 1, "CUIT", "CUIT", TamanioColumna
    
    TamanioColumna = 3000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 2, "Descripción", "Descripción", TamanioColumna
    
    TamanioColumna = 500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 3, "Cód.", "Cód.", TamanioColumna
    
    TamanioColumna = 1500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 4, "Reducida", "Reducida", TamanioColumna
    
    TamanioColumna = 3000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 5, "Dirección", "Dirección", TamanioColumna
    
    TamanioColumna = 1500
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 6, "Localidad", "Localidad", TamanioColumna
    
    TamanioColumna = 1000
    TamanioTotal = TamanioTotal + TamanioColumna
    LvListado.ColumnHeaders.Add 7, "C.P.", "C.P.", TamanioColumna
    
    'el tamaño de la última columna se actualiza automáticamente para acomodarse al listview
    LvListado.ColumnHeaders.Add 8, "Teléfono", "Teléfono", LvListado.Width - TamanioTotal - 100

Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub




Private Sub LVListado_DblClick()
    If LvListado.ListItems.Count > 0 Then
        Codigo.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(1))
        Descripcion.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).Text)
        Reducida.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(2))
        Direccion.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(3))
        Localidad.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(4))
        CodigoPostal.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(5))
        Telefono.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(6))
        CUIT.Caption = Trim(LvListado.ListItems(LvListado.SelectedItem.Index).SubItems(7))
        
    End If
    Me.Visible = False

End Sub

Private Sub OptOrden_Click(Index As Integer)
        If OptOrden(1).Value = True Then
            LvListado.ColumnHeaders.Clear
            CargarEncabezados
            CargarListado
        Else
            LvListado.ColumnHeaders.Clear
            CargarEncabezadosXCuit
            CargarListadoxCuit
        End If
End Sub
Private Sub CargarListadoxCuit()
On Error GoTo Errores
Dim sSQL As String
        MousePointer = vbHourglass
        sSQL = "SpTAProveedoresXCuit"
        Set v.TbListado = v.db.OpenResultset(sSQL)
        LvListado.ListItems.Clear
        LvListado.Sorted = False
        ' esto activa el timer para empezar a cargar los renglones
        Timer1.Enabled = True
        MousePointer = vbNormal

Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errores
Dim i As Integer
    i = 1
    With v.TbListado
    'cargo de a 25 renglones para no perder tanto tiempo y luego vuelve a arrancar el timer
    
        While Not .EOF And i < 25
            LvListado.ListItems.Add
            If OptOrden(1).Value = True Then
                LvListado.ListItems(LvListado.ListItems.Count).Text = VerificarNulo(!P_Descripcion)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(1) = VerificarNulo(!P_Codigo)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(2) = VerificarNulo(!P_Reducida)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(3) = VerificarNulo(!P_Direccion)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(4) = VerificarNulo(!P_Localidad)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(5) = VerificarNulo(!P_CodigoPostal)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(6) = VerificarNulo(!P_Telefono)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(7) = VerificarNulo(!P_Cuit)
                OptOrden(1).Value = True
            Else
                LvListado.ListItems(LvListado.ListItems.Count).Text = VerificarNulo(!P_Cuit)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(1) = VerificarNulo(!P_Descripcion)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(2) = VerificarNulo(!P_Codigo)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(3) = VerificarNulo(!P_Reducida)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(4) = VerificarNulo(!P_Direccion)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(5) = VerificarNulo(!P_Localidad)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(6) = VerificarNulo(!P_CodigoPostal)
                LvListado.ListItems(LvListado.ListItems.Count).SubItems(7) = VerificarNulo(!P_Telefono)
            End If
            'el vector vectotales es un acumulador para calcular los totales
            i = i + 1
            .MoveNext
        Wend
    If .EOF Then
        'cuando terminé de calcular todos los renglones deshabilito el timer
        .Close
        Timer1.Enabled = False
    End If
    End With
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error GoTo Errores
    TeclaPresionada ActiveControl, KeyAscii
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LvListado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Errores
   ' Cuando se hace clic en un objeto ColumnHeader, el
   ' control ListView se ordena por los subelementos de
   ' esa columna.
   ' Establece el SortKey como el Index del ColumnHeader - 1
   ' Asigna a Sorted el valor True para ordenar la lista.
   LvListado.SortKey = ColumnHeader.Index - 1
   If LvListado.SortOrder = lvwAscending Then
        LvListado.SortOrder = lvwDescending
   Else
       LvListado.SortOrder = lvwAscending
   End If
   LvListado.Sorted = True
Errores:
    ManipularError Err.Number, Err.Description, Timer1
End Sub

