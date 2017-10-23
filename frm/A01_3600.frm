VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{242A80DB-94C3-4BA9-BA6B-EC6D66393472}#13.0#0"; "ComboEspecial.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form A01_3600 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos de Compras"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   720
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   1485
      Top             =   7785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton CmdExpPdf 
      Caption         =   "Exportar a PDF"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2610
      TabIndex        =   29
      Top             =   5190
      Width           =   1230
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular"
      Height          =   330
      Left            =   1335
      TabIndex        =   28
      Top             =   5190
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdNueva 
      Caption         =   "&Nueva"
      Height          =   350
      Left            =   3885
      TabIndex        =   14
      Top             =   5180
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   75
      Top             =   5085
   End
   Begin VB.Frame FraArt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Artículo"
      Height          =   2190
      Left            =   5910
      TabIndex        =   23
      Top             =   2250
      Width           =   3120
      Begin VB.TextBox TxtCodArticulo 
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   160
         Width           =   1005
      End
      Begin VB.CommandButton CmdModif 
         Caption         =   "Modificar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   195
         TabIndex        =   11
         Top             =   1320
         Width           =   1300
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   350
         Left            =   195
         TabIndex        =   13
         Top             =   1740
         Width           =   1300
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Item"
         Height          =   350
         Left            =   1635
         TabIndex        =   12
         Top             =   1320
         Width           =   1300
      End
      Begin VB.TextBox TxtCant 
         Height          =   315
         Left            =   2025
         TabIndex        =   10
         Top             =   945
         Width           =   1005
      End
      Begin Controles.ComboEsp CmbArtCompra 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   540
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cantidad:"
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
         Left            =   1140
         TabIndex        =   25
         Top             =   1005
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Artículo:"
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
         Left            =   90
         TabIndex        =   24
         Top             =   220
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5160
      TabIndex        =   15
      Top             =   5180
      Width           =   1230
   End
   Begin VB.CommandButton CmdCerra 
      Caption         =   "&Cerrar"
      Height          =   350
      Left            =   7785
      TabIndex        =   18
      Top             =   5180
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Requerimiento"
      Height          =   2130
      Left            =   45
      TabIndex        =   19
      Top             =   45
      Width           =   8970
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prioridad"
         Height          =   840
         Left            =   7470
         TabIndex        =   33
         Top             =   870
         Width           =   1380
         Begin VB.OptionButton OptPrioridad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Urgente"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   510
            Width           =   1215
         End
         Begin VB.OptionButton OptPrioridad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Normal"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1755
         Width           =   7440
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   3210
         TabIndex        =   1
         Top             =   225
         Width           =   900
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4170
         TabIndex        =   2
         Top             =   225
         Width           =   900
      End
      Begin VB.TextBox TxtSolicitante 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   4
         Top             =   675
         Width           =   4005
      End
      Begin VB.TextBox TxtNroOrden 
         Height          =   315
         Left            =   1935
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker CalFecha 
         Height          =   315
         Left            =   7455
         TabIndex        =   3
         Top             =   495
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   " "
         Format          =   49741825
         UpDown          =   -1  'True
         CurrentDate     =   38993
      End
      Begin Controles.ComboEsp CmbCentroDeCostoEmisor 
         Height          =   315
         Left            =   2355
         TabIndex        =   5
         Top             =   1050
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin Controles.ComboEsp CmbCentroDeCostoDestino 
         Height          =   315
         Left            =   2340
         TabIndex        =   31
         Top             =   1410
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ListIndex       =   -1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo Destino:"
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
         Left            =   150
         TabIndex        =   32
         Top             =   1470
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha:"
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
         Left            =   6765
         TabIndex        =   20
         Top             =   555
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Observaciones:"
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
         Left            =   90
         TabIndex        =   30
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label LBAnulada 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anulada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7005
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costo Emisor:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1110
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº de Requerimiento:"
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
         Left            =   90
         TabIndex        =   22
         Top             =   285
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitado Por:"
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
         Left            =   645
         TabIndex        =   21
         Top             =   735
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView LvListado 
      Height          =   2850
      Left            =   45
      TabIndex        =   7
      Top             =   2250
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5027
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdConfirnar 
      Caption         =   "&Guardar Nueva"
      Height          =   350
      Left            =   6440
      TabIndex        =   17
      Top             =   5190
      Width           =   1300
   End
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   6440
      TabIndex        =   16
      Top             =   5180
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label LbFechaProbableDeEntrega 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Prob. Entrega: "
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
      Left            =   5925
      TabIndex        =   36
      Top             =   4575
      Width           =   1890
   End
End
Attribute VB_Name = "A01_3600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VecOrdenDeCompra() As TipoOrdenDeCompra
'este vector no se aparea con el Lv
Private VecCentroCta() As TipoCentroCta

Private MontoSinPres As Double
Private Modificado As Boolean
Private A_Codigo As Long
Public NroOrden As Integer
Private Nivel As Integer
Private CantNoAsignada As Long
Private VecArtCompra() As TipoArticuloCompras
Private VecRequerimientoCompra() As TipoRequerimientoCompra
Public CentroEmisorActual As String
Public CentroEmisorDestino As String
Public TablaArticulos As String
Public TablaRequerimientos As String
Private RsRenglones As ADODB.Recordset
Private Requerimiento As Boolean

Private Sub CalFecha_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Modificado = True
End Sub

Private Function Existe(Articulo As Long) As Integer
    Dim i As Integer
    Existe = 0
    For i = 1 To UBound(VecOrdenDeCompra)
        If VecOrdenDeCompra(i).A_Codigo = Articulo Then
            Existe = i
            Exit Function
        End If
    Next
End Function

Private Sub CmbCentroDeCostoDestino_Validate(Cancel As Boolean)
    'Modificado = True
    TablaArticulos = VecCentroDeCostoEmisor(CmbCentroDeCostoDestino.ListIndex).C_TablaArticulos
    TxtCodArticulo.Visible = TablaArticulos <> ""
    Call CargarVecCentro(VecCentroDeCostoEmisor(CmbCentroDeCostoDestino.ListIndex).C_Codigo)
    CentroEmisorDestino = VecCentroDeCostoEmisor(CmbCentroDeCostoDestino.ListIndex).C_Codigo
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    ReDim VecRequerimientoCompra(0)
    CmbCentroDeCostoDestino.Enabled = False
End Sub


Private Sub TxtCodArticulo_KeyPress(KeyAscii As Integer)
    'controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 _
       And KeyAscii <> Asc(".") Then
          Beep
          KeyAscii = 0
    End If
End Sub

Private Sub TxtCodArticulo_Validate(Cancel As Boolean)
    If TxtCodArticulo.Text <> "" Then
       Call BuscarArt(Val(TxtCodArticulo), CmbArtCompra)
    End If
End Sub

Private Sub CmdAnular_Click()
  On Error GoTo ErrorAnulacion
    Dim Sql As String
    Dim Rta As Integer

    Dim RsAnular As New ADODB.Recordset
       
    Rta = MsgBox("¿Confirma que desea anular el Requerimiento?", vbYesNo)
    
    If Rta = vbYes Then
        
        Sql = "SpOcRequerimientosDeCompraCabeceraAnular @R_Numero = " & NroOrden & _
                                                     ", @R_UsuarioAnulacion = '" & Usuario & "'"
        RsAnular.Open Sql, Conec
        'If RsAnular!OK = "OK" Then
            MsgBox RsAnular!Mensaje
        'End If
    End If
ErrorAnulacion:
 If Err.Number <> 0 Then
    Call ManipularError(Err.Number, Err.Description)
 End If

End Sub

Private Sub CMDBuscar_Click()
    NroOrden = 0
    Unload A01_3610
    A01_3610.CentroEmisor = CentroEmisorActual
    A01_3610.Show vbModal
    Timer1.Enabled = True
End Sub

Private Sub CmdCambiar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea Modificar El Raquerimineto de Compra?", vbYesNo)
    If Rta = vbYes Then
        Call ModificarOrden
    End If
End Sub

Private Sub ModificarOrden()
  Dim Sql As String
  Dim i As Integer
  Dim Rta As Integer
  
On Error GoTo ErrorUpdate

  NroOrden = Val(TxtNroOrden.Text)
   If ValidarEncabezado Then
     Conec.BeginTrans
     Sql = "SpOcRequerimientosDeCompraCabeceraModificar @R_Numero =" & NroOrden & _
                                                    " , @R_Prioridad =" & IIf(OptPrioridad(0).Value, 0, 1) & _
                                                    " , @R_Usuario ='" & Usuario & _
                                                    "', @R_Observaciones ='" & TxtObs.Text & "'"
           
    Conec.Execute Sql
    Sql = "SpOcRequerimientosDeCompraRenglonesBorrar @R_Numero=" & NroOrden
    Conec.Execute Sql
    
    For i = 1 To UBound(VecRequerimientoCompra)
      With VecRequerimientoCompra(i)
        Sql = "SpOcRequerimientosDeCompraRenglonesAgregar @R_Numero =" & NroOrden & _
                                                       ", @R_Articulo =" & .CodArticulo & _
                                                       ", @R_Cantidad =" & Replace(.Cantidad, ",", ".") & _
                                                       ", @R_CantidadPendiente= " & Replace(.CantidadPendiente, ",", ".")
      End With
        Conec.Execute Sql
    Next

    Conec.CommitTrans

ErrorUpdate:
    If Err.Number = 0 Then
       MsgBox "El Requerimiento de compra se Grabó correctamente"
       Modificado = False
       Call CmdCargar_Click
    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Sub

Private Sub CmdCargar_Click()
    Call CargarOrden(Val(TxtNroOrden))
    Modificado = False
End Sub

Private Sub CmdCerra_Click()
    Unload Me
End Sub

Private Sub CmdConfirnar_Click()
Dim Rta As Integer
    Rta = MsgBox("¿Desea confirmar El Requerimiento de Compra?", vbYesNo)
    If Rta = vbYes Then
       Call GrabarOrden
    End If
End Sub

Private Sub GrabarOrden()
  Dim Sql As String
  Dim RsGrabar As New ADODB.Recordset
  Dim Rta As Integer
  Dim i As Integer

On Error GoTo ErrorInsert
  
  If ValidarEncabezado Then
    Conec.BeginTrans
     Sql = "SpOcRequerimientosDeCompraCabeceraAgragar @R_CentroEmisor ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoEmisor.ListIndex).C_Codigo & _
                                                  "', @R_CentroDestino ='" & VecCentroDeCostoEmisor(CmbCentroDeCostoDestino.ListIndex).C_Codigo & _
                                                  "', @R_Fecha = " & FechaSQL(CalFecha, "SQL") & _
                                                  " , @R_Prioridad =" & IIf(OptPrioridad(0).Value, 0, 1) & _
                                                  " , @R_Usuario ='" & Usuario & _
                                                  "', @R_Observaciones ='" & TxtObs.Text & "'"
           
     'graba el encabezado y retorna el Nro de Orden
        RsGrabar.Open Sql, Conec
        NroOrden = RsGrabar!R_Numero
        
    For i = 1 To UBound(VecRequerimientoCompra)
      With VecRequerimientoCompra(i)
        Sql = "SpOcRequerimientosDeCompraRenglonesAgregar @R_Numero =" & NroOrden & _
                                                       ", @R_Articulo =" & .CodArticulo & _
                                                       ", @R_Cantidad =" & Replace(.Cantidad, ",", ".") & _
                                                       ", @R_CantidadPendiente= " & Replace(.CantidadPendiente, ",", ".")
      End With
        Conec.Execute Sql
    Next

    Conec.CommitTrans
    
ErrorInsert:
    If Err.Number = 0 Then
       CmdConfirnar.Visible = False
       CmdCambiar.Visible = True
       CmdImprimir.Enabled = True
       CmdExpPdf.Enabled = True
       
       Modificado = False
       Call ConfImpresionDeOrden
       RepRequerimiento.Show vbModal

    Else
       Conec.RollbackTrans
       Call ManipularError(Err.Number, Err.Description)
    End If
  End If
End Sub

Private Function ValidarEncabezado() As Boolean
Dim i As Integer

    ValidarEncabezado = True
 
     If CmbCentroDeCostoDestino.ListIndex = 0 Then
        MsgBox "Debe Seleccionar un Centro de Costo Destino"
        CmbCentroDeCostoDestino.SetFocus
        ValidarEncabezado = False
        Exit Function
    End If
      
    If LvListado.ListItems.Count <= 1 Then
       MsgBox "Debe ingresar artículos al Requerimiento"
       LvListado.SetFocus
       ValidarEncabezado = False
       Exit Function
    End If
        
End Function

Private Sub CmdEliminar_Click()
    Dim IndexBorrar As Integer
    
    IndexBorrar = LvListado.SelectedItem.Index
    
    'borra del LV
 If VecRequerimientoCompra(IndexBorrar).Cantidad = VecRequerimientoCompra(IndexBorrar).CantidadPendiente Then
      LvListado.ListItems.Remove (IndexBorrar)
      'borrar del vector haciento un corrimiento
      While IndexBorrar < UBound(VecRequerimientoCompra)
          VecRequerimientoCompra(IndexBorrar) = VecRequerimientoCompra(IndexBorrar + 1)
          IndexBorrar = IndexBorrar + 1
      Wend
      

      ReDim Preserve VecRequerimientoCompra(UBound(VecRequerimientoCompra) - 1)

      Modificado = True
      
      If LvListado.ListItems.Count = LvListado.SelectedItem.Index Then
          CmdEliminar.Enabled = False
          CmdModif.Enabled = False
      End If
  Else
      MsgBox "El renglón no puede ser borrado por Haber artículos recibidos", , "Borrar"
  End If
End Sub

Private Sub ConfImpresionDeOrden()
  Dim i As Integer
  Dim RsListado As ADODB.Recordset
  Set RsListado = New ADODB.Recordset
    
    RsListado.Fields.Append "Articulo", adVarChar, 100
    RsListado.Fields.Append "Cantidad", adDouble
    RsListado.Open
    i = 1
    While i < LvListado.ListItems.Count
        RsListado.AddNew
      With LvListado.ListItems(i)
        RsListado!Articulo = .Text
        RsListado!Cantidad = ValN(.SubItems(1))
      End With
        i = i + 1
    Wend
    If Not RsListado.EOF Then
        RsListado.MoveFirst
    End If
    TxtNroOrden.Text = Format(NroOrden, "0000000000")
    RepRequerimiento.TxtObservaciones = TxtObs
    RepRequerimiento.TxtCentroEmisor.Text = CmbCentroDeCostoEmisor.Text
    RepRequerimiento.TxtCentroDestino.Text = CmbCentroDeCostoDestino.Text
    RepRequerimiento.TxtFecha = Format(CalFecha.Value, "dd/MM/yyyy")
    RepRequerimiento.TxtNroOrden.Text = TxtNroOrden.Text
    RepRequerimiento.TxtResp.Text = TxtSolicitante.Text
    RepRequerimiento.TxtAnulada.Visible = LBAnulada.Visible
    RepRequerimiento.TxtAnulada.Text = LBAnulada.Caption
    RepRequerimiento.TxtPrioridad = IIf(OptPrioridad(0).Value, OptPrioridad(0).Caption, OptPrioridad(1).Caption)
    RepRequerimiento.DataControl1.Recordset = RsListado
    RepRequerimiento.Zoom = -1
End Sub

Private Sub CmdImprimir_Click()
    Call ConfImpresionDeOrden
    RepRequerimiento.Show
End Sub

Private Sub CmdModif_Click()
On Error GoTo Errores
Dim i As Integer
    i = LvListado.SelectedItem.Index
    
 If ValidarCargaOC Then

   If VecRequerimientoCompra(i).CantidadPendiente + (Val(TxtCant.Text) _
       - VecRequerimientoCompra(i).Cantidad) >= 0 Then
        Modificado = True
        A_Codigo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
       'agrega al vector
       With VecRequerimientoCompra(i)
            .CodArticulo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
            .DescArticulo = CmbArtCompra.Text
            .CantidadPendiente = .CantidadPendiente + (Val(TxtCant.Text) - .Cantidad)
            .Cantidad = Val(TxtCant.Text)
            'lo pone en el LV
           LvListado.ListItems(i).Text = Trim(.DescArticulo)
           LvListado.ListItems(i).SubItems(1) = Format(.Cantidad, "0.00##")
        End With

    Else
        MsgBox "La cantidad pedida de artículos debe ser mayor a la entregada para permitir su modificación"
    End If
 End If

Errores:
  Call ManipularError(Err.Number, Err.Description, Timer1)
End Sub

Private Sub CmdNueva_Click()
    TxtNroOrden.Text = ""
    
    CmdConfirnar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
    'esto es por que cuando se trae una orden anulada los
    'valores están invertidos
     LBAnulada.Visible = False
     FraArt.Enabled = True
    CmbCentroDeCostoDestino.ListIndex = 0
    'busca y establece el centro de costo emisor del usuario actual del sistema
     'Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoDestino)
     CalFecha.Enabled = True
     Modificado = False
     Requerimiento = False
End Sub

Private Sub Form_Load()
    CentroEmisorActual = CentroEmisor
    ReDim VecRequerimientoCompra(0)
    Call CrearEncabezados
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoEmisor)
    Call CargarCmbCentrosDeCostosEmisor(CmbCentroDeCostoDestino)
    CmbCentroDeCostoDestino.ListIndex = 0
    Call BuscarCentroEmisor(CentroEmisor, CmbCentroDeCostoEmisor)
    Nivel = TraerNivel("A013600")
    TxtNroOrden.Text = ""
    CalFecha.Value = Date
    TxtSolicitante.Text = NombreUsuario
    Modificado = False
End Sub

Private Sub CmdExpPdf_Click()
On Error GoTo Error
     MenuEmisionOrdenCompra.Cuadros.Filter = "*.pdf"
     MenuEmisionOrdenCompra.Cuadros.ShowSave
     
  If MenuEmisionOrdenCompra.Cuadros.Filename <> "" Then
         Call ConfImpresionDeOrden
         RepRequerimiento.Run
        'guarda la orden de compra como un PDF
         Dim myPDFExport As ActiveReportsPDFExport.ARExportPDF
         Set myPDFExport = New ActiveReportsPDFExport.ARExportPDF
         myPDFExport.AcrobatVersion = DDACR40
         
         myPDFExport.Filename = NombreArchivoPDF(MenuEmisionOrdenCompra.Cuadros.Filename)
         myPDFExport.JPGQuality = 100
         myPDFExport.SemiDelimitedNeverEmbedFonts = ""
         myPDFExport.Export RepRequerimiento.Pages
         Unload RepRequerimiento
         MsgBox "La Exportación se ralizó correctamente", vbInformation, "Exportación"
  End If
Error:
     Call ManipularError(Err.Number, Err.Description)
 

End Sub


Private Sub ValidarCantArticulos(A_Codigo As Long)
   Dim i As Integer
   Dim cant As Double
   
   For i = 1 To UBound(VecCentroCta)
      If VecCentroCta(i).O_CodigoArticulo = A_Codigo Then
         cant = cant + VecCentroCta(i).O_CantidadPedida
      End If
   Next
   
   i = IIf(LvListado.SelectedItem.Index < LvListado.ListItems.Count, LvListado.SelectedItem.Index, LvListado.SelectedItem.Index - 1)
   
   If VecOrdenDeCompra(i).Cantidad = cant Then
     LvListado.ListItems(i).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(1).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(2).ForeColor = vbBlack
     LvListado.ListItems(i).ListSubItems(3).ForeColor = vbBlack

   Else
     LvListado.ListItems(i).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(1).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(2).ForeColor = vbRed
     LvListado.ListItems(i).ListSubItems(3).ForeColor = vbRed
   End If
End Sub

Private Sub CmdAgregar_Click()
'On Error GoTo errores
Dim i As Integer
    
 If ValidarCargaOC Then
    Modificado = True
        
    If LvListado.SelectedItem.Index = LvListado.ListItems(LvListado.ListItems.Count).Index Then
       'agrega al vector
        i = LvListado.SelectedItem.Index
        ReDim Preserve VecRequerimientoCompra(UBound(VecRequerimientoCompra) + 1)
          
        'lo pone en el LV
        With VecRequerimientoCompra(i)
            A_Codigo = .CodArticulo
            .Cantidad = Val(TxtCant.Text)
            .CantidadPendiente = Val(TxtCant.Text)
            .CodArticulo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo
            .DescArticulo = CmbArtCompra.Text

           LvListado.ListItems(i).Text = Trim(.DescArticulo)
           LvListado.ListItems(i).SubItems(1) = Format(.Cantidad, "0.00##")
          ' LvListado.ListItems(i).SubItems(2) = Format(.PrecioUnit, "0.00##")
           'LvListado.ListItems(i).SubItems(3) = Format(.Cantidad * .PrecioUnit, "0.00##")
           
        End With
    'es el último registro, por lo tanto quería agregar uno nuevo
            LvListado.ListItems.Add
         'pocisiona en el último
            LvListado.ListItems(LvListado.ListItems.Count).Selected = True
            Call LimpiarOC
     End If
     'le da el foco al combo de artículoa
      CmbArtCompra.SetFocus
  End If
Errores:
  Call ManipularError(Err.Number, Err.Description)
End Sub

Private Function ValidarCargaOC() As Boolean
On Error GoTo Error
    ValidarCargaOC = True
    Dim i As Integer
    Dim Rta As Integer
    
    If CmbCentroDeCostoEmisor.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Centro de Costo Emisor"
       CmbCentroDeCostoEmisor.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
    
    For i = 1 To UBound(VecRequerimientoCompra)
      If i <> LvListado.SelectedItem.Index Then
       If VecRequerimientoCompra(i).CodArticulo = VecArtCompra(CmbArtCompra.ListIndex).A_Codigo Then
          MsgBox "Ese artículo ya existe en esta Orden de compra"
          ValidarCargaOC = False
          Exit Function
       End If
      End If
    Next
    
    If CmbArtCompra.ListIndex = 0 Then
       MsgBox "Debe Seleccionar un Artículo"
       CmbArtCompra.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
       
    If Val(TxtCant.Text) = 0 Then
       MsgBox "Debe Ingresar una cantidad mayor que 0"
       TxtCant.SetFocus
       ValidarCargaOC = False
       Exit Function
    End If
           
     
    Exit Function
Error:
    ValidarCargaOC = False
    Call ManipularError(Err.Number, Err.Description)
End Function

Private Sub CrearEncabezados()
    LvListado.ColumnHeaders.Add , , "Descripción Artículo", LvListado.Width - 1050
    LvListado.ColumnHeaders.Add , , "Cant.", 750, 1
End Sub

Private Sub CalcularTotal()
        'acumular el total
Dim i As Integer
Dim Total As Double
    For i = 1 To LvListado.ListItems.Count
       ' Total = Total + Val(Replace(LvListado.ListItems(i).SubItems(3), ",", "."))
    Next
       ' txtTotal.Text = Format(Total, "$ 0.00##")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Rta As Integer
    If Modificado Then
       Rta = MsgBox("Ha efectuado cambio ¿Desea Guardarlos?", vbYesNoCancel)
       If Rta = vbCancel Then
         Cancel = 1
         Exit Sub
       Else
         If Rta = vbYes Then
            Call ModificarOrden
         End If
       End If
    End If
End Sub

Private Sub Timer1_Timer()
   If NroOrden <> 0 Then
      TxtNroOrden.Text = CStr(NroOrden)
      Call CmdCargar_Click
   End If
   
    Timer1.Enabled = False
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
    Call TxtNumerico2(TxtCant, KeyAscii)
End Sub

Private Sub LvListado_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo errores
'NO SE TOCA
   If Item.Index < LvListado.ListItems.Count Then
        Call CargarEnModificar(Item.Index)
        CmdModif.Enabled = True
        CmdEliminar.Enabled = True
        CmdAgregar.Enabled = False
    Else
        CmdAgregar.Enabled = True
        CmdModif.Enabled = False
        CmdEliminar.Enabled = False
        TxtCant.Enabled = True
        CmbArtCompra.Enabled = True
   End If
'errores:
 '   ManipularError Err.Number, Err.Description, Timer1
End Sub

Private Sub LimpiarOC()
    CmbArtCompra.ListIndex = 0
    TxtCant.Text = ""
    TxtCodArticulo.Text = ""
End Sub

Public Sub BuscarArt(Codigo As Long, CmbArt As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = Codigo Then
           Exit For
        End If
   Next
   If i <= UBound(VecArtCompra) Then
      CmbArt.ListIndex = i
   End If
End Sub

Public Function BuscarCtaPorDefectoArt(Codigo As Long) As String
 Dim i As Integer
   For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = Codigo Then
           BuscarCtaPorDefectoArt = VecArtCompra(i).A_CuentaPorDefecto
           Exit Function
        End If
   Next
End Function

Private Sub CargarEnModificar(Index As Integer)
   Call BuscarArt(VecRequerimientoCompra(Index).CodArticulo, CmbArtCompra)
  'esta variable se usa para cargar el lvCentroCta
   A_Codigo = VecRequerimientoCompra(Index).CodArticulo
    If VecRequerimientoCompra(Index).FechaProbableDeEntrega <> "" Then
        LbFechaProbableDeEntrega.Caption = "Fecha Prob. De Entrega: " & VecRequerimientoCompra(Index).FechaProbableDeEntrega
    Else
        LbFechaProbableDeEntrega.Caption = ""
    End If

   TxtCant.Text = Replace(VecRequerimientoCompra(Index).Cantidad, ",", ".")
End Sub

Private Sub LimpiarOrden()
    TxtCant.Text = ""
    TxtSolicitante.Text = NombreUsuario

    TxtObs.Text = ""
    TxtObs.Enabled = True
    CmbCentroDeCostoDestino.Enabled = True
    'CalFecha.Enabled = False
    NroOrden = 0
    
    LvListado.ListItems.Clear
    LvListado.ListItems.Add
    LvListado.ListItems(1).Selected = True
    Call LvListado_ItemClick(LvListado.SelectedItem)
    
    ReDim VecRequerimientoCompra(0)

    CmbArtCompra.ListIndex = 0
   ' CmbCentroDeCostoEmisor.ListIndex = 0
    CalFecha.Value = Date
 
    CmdConfirnar.Visible = True
    CmdAnular.Visible = False
    CmdCambiar.Visible = False
    CmdImprimir.Enabled = False
    CmdExpPdf.Enabled = False
    
    TxtCodArticulo.Visible = False
End Sub

Private Sub TxtNroOrden_KeyPress(KeyAscii As Integer)
 ' controla que solo se ingresen números
    If KeyAscii = 13 Then
       Call CmdCargar_Click
    Else
       If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 Then
          Beep
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub TxtNroOrden_LostFocus()
  If Val(TxtNroOrden.Text) <> NroOrden Then
    CmdConfirnar.Visible = TxtNroOrden.Text = ""
    CmdCambiar.Visible = TxtNroOrden.Text <> ""
    Call LimpiarOrden
  End If
End Sub

Private Sub CargarOrden(NroOrden As Integer)
    Dim Sql As String
    Dim i As Integer

    Dim RsCargar As New ADODB.Recordset
    Dim PeriodoCerrado As Boolean
    Dim RsValidarPeriodo As New ADODB.Recordset
    Dim Autorizado As Boolean
    LBAnulada.Visible = False
    
    If RsRenglones Is Nothing Then
        Set RsRenglones = New ADODB.Recordset
        RsRenglones.CursorType = adOpenKeyset
        RsRenglones.LockType = adLockBatchOptimistic
    Else
        If RsRenglones.State = adStateOpen Then
            RsRenglones.Close
        End If
    End If
    
  With RsCargar
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        '.LockType = adLockPessimistic
        
        Sql = "SpOcRequerimientosDeCompraCabeceraTraer @R_Numero= " & NroOrden
        .Open Sql, Conec
        
      If .EOF Then
          MsgBox "No existe Requerimineto con esa numeración", vbInformation
          Exit Sub
      Else
        If !R_CentroEmisor <> CentroEmisor Then
          MsgBox "El Requerimineto Pertenese a Otro Centro de Costo", vbInformation
          Exit Sub
        End If
          CalFecha.Enabled = False
          'CalFecha.Format = dtpShortDate
          CmbCentroDeCostoEmisor.Enabled = False
           
        If Not IsNull(!R_FechaAnulacion) Then
            If Not IsNull(!R_FechaAnulacion) Then
                LBAnulada.Caption = "Anulada " + Mid(CStr(!R_FechaAnulacion), 1, 10)
                LBAnulada.Visible = True
            End If
            CmdAnular.Visible = False
            TxtObs.Enabled = False
        Else
            CmdAnular.Visible = True
            LBAnulada.Visible = False
            TxtObs.Enabled = True
        End If
        
        TxtNroOrden.Text = Format(!R_Numero, "0000000000")
        Me.NroOrden = !R_Numero
        
        CmdConfirnar.Visible = False
        CmdCambiar.Visible = True
        CmdImprimir.Enabled = True
        CmdExpPdf.Enabled = True
        CalFecha.Value = !R_Fecha
        'TxtResp = RsCargar!O_Responsable
        Call BuscarCentroEmisor(!R_CentroDestino, CmbCentroDeCostoDestino)
        Call BuscarCentroEmisor(!R_CentroEmisor, CmbCentroDeCostoEmisor)
        Call CmbCentroDeCostoDestino_Validate(False)
        TxtObs.Text = VerificarNulo(!R_Observaciones)
        OptPrioridad(!R_Prioridad).Value = True
        'If Not IsNull(!R_FechaProbableDeEntrega) Then
        '    LbFechaProbableDeEntrega.Caption = "Fecha Prob. De Entrega: " & !R_FechaProbableDeEntrega
        'Else
        '    LbFechaProbableDeEntrega.Caption = ""
        'End If
        .Close
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        
        Sql = "SpOcRequerimientosDeCompraRenglonesTraer @R_Numero=" & NroOrden
        .Open Sql, Conec
        ReDim VecRequerimientoCompra(.RecordCount)
        i = 1
        LvListado.ListItems.Clear
        
        While Not .EOF
            LvListado.ListItems.Add
            VecRequerimientoCompra(i).CodArticulo = !R_Articulo
            VecRequerimientoCompra(i).DescArticulo = BuscarDescArt(!R_Articulo)
            VecRequerimientoCompra(i).Cantidad = Format(!R_Cantidad, "0.00##")
            VecRequerimientoCompra(i).CantidadPendiente = !R_CantidadPendiente
            VecRequerimientoCompra(i).FechaProbableDeEntrega = VerificarNulo(!R_FechaProbableDeEntrega)
            
            LvListado.ListItems(i).Text = VecRequerimientoCompra(i).DescArticulo
            LvListado.ListItems(i).SubItems(1) = Format(VecRequerimientoCompra(i).Cantidad, "0.00##")
            .MoveNext
            i = i + 1
        Wend
        
        LvListado.ListItems.Add
        LvListado.ListItems(LvListado.ListItems.Count).Selected = True

      End If
  End With
  
  LvListado.ListItems(1).Selected = True
  Call LvListado_ItemClick(LvListado.ListItems(1))

End Sub

Private Sub CargarVecCentro(Centro As String)
 Dim Sql As String
 Dim RsCargar As New ADODB.Recordset
 Dim i As Integer
 'dependiendo del centro de costo emisor carga las cuentas y artículos correspondientes
    Sql = "SpOCRelacionCentroDeCostoArticulosTraer @R_CentroDeCosto='" & Centro & "'"
    With RsCargar
       ReDim VecArtCompra(0)
       If TablaArticulos = "" Then
             .Open Sql, Conec, adOpenStatic, adLockReadOnly
            'en esta sección carga los art
              For i = 1 To UBound(VariablesYFunciones.VecArtCompra)
                  .Find "R_Articulo = " & VariablesYFunciones.VecArtCompra(i).A_Codigo, , , 1
                 If Not .EOF Then
                    ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                    VecArtCompra(UBound(VecArtCompra)) = VariablesYFunciones.VecArtCompra(i)
                 End If
              Next
             .Close
        Else
            For i = 1 To UBound(VecArtTaller)
                ReDim Preserve VecArtCompra(UBound(VecArtCompra) + 1)
                VecArtCompra(UBound(VecArtCompra)) = VecArtTaller(i)
            Next
        End If
    End With
    'carga los combos con los valores de los vectores locales
    Call CargarCmbArtCompra(CmbArtCompra)
End Sub

Public Sub CargarCmbArtCompra(CmbAtrCompra As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbAtrCompra.Clear
    For i = 0 To UBound(VecArtCompra)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbAtrCompra.AddItem "Seleccione un Artículo"
           Else
              CmbAtrCompra.AddItem "Todos los Artículos"
           End If
        Else
            CmbAtrCompra.AddItem VecArtCompra(i).A_Descripcion
        End If
    Next
        
    CmbAtrCompra.ListIndex = 0
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarCmbCuentasContables(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    
    If Tipo = "Elegir" Then
       Cmb.AddItem "Seleccione una Cuenta Contable"
    Else
       Cmb.AddItem "Todas las Cuentas Contables"
    End If

    For i = 1 To UBound(VecCuentasContables)
        Cmb.AddItem Trim(VecCuentasContables(i).Descripcion)
    Next
        
    Cmb.ListIndex = 0
Errores:
   Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub UbicarCuentaContable(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
   
    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = C_Codigo Then
           Cmb.ListIndex = i
           Exit For
        End If
    Next
    
End Sub

Public Function BuscarDescArt(A_Codigo As Long) As String
    Dim i As Integer
    For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = A_Codigo Then
            BuscarDescArt = VecArtCompra(i).A_Descripcion
            Exit Function
        End If
    Next
End Function
