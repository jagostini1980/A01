Attribute VB_Name = "VariablesYFunciones"
Option Explicit

Public Enum Opciones
    vbimprimir
    vbExportesPDF
    vbCerrar
    vbNuevo
End Enum

Public Type TipoArticuloCompras
    A_Codigo As Long
    A_Descripcion As String
    A_CuentaPorDefecto As String
End Type

Public Type TipoLugarDeEntrega
    L_Codigo As Long
    L_Descripcion As String
    L_EMail As String
    Estado As String
End Type

Public Type TipoFormaDePago
    F_Codigo As Integer
    F_Descripcion As String
    Estado As String
End Type

Public Type TipoRubroContable
    Codigo As String
    Descripcion As String
End Type

Public Type TipoCentroDeCosto
    C_Codigo As String
    C_Descripcion As String
    C_Padre As String
    C_UsuarioResponsable As String
    C_TablaArticulos As String
    C_MontoSinPresupuestar As Double
    C_Jerarquia As String
    C_Nivel As Integer
    C_TablaRequerimientosDeCompra As String
    C_ImporteAutorizar As Double
    C_EmailAutorizacion As String
End Type

Public Type TipoOrdenDeCompra
    A_Codigo As Long
    A_Descripcion As String
    Cantidad As Double
    CantPendiente As Long
    PrecioUnit As Double
    MontoSinPres As Double
    Requerimiento As Boolean
End Type

Public Type TipoCentroCta
   O_CodigoArticulo As Long
   O_CuentaContable As String
   Cta_Descripcion As String
   O_CentroDeCosto As String
   Centro_Descripcion As String
   O_CantidadPedida As Double
   O_CantidadPendiente As Double
   O_SinPresupuestar As Boolean
   O_MontoSinPresupuestar As Double
End Type

Public Type TipoCentroCtaRecepcion
   NroOrden As Integer
   FechaOrden As String
   O_CentroDeCostoEmisor As String
   O_CodigoArticulo As Long
   O_CuentaContable As String
   Cta_Descripcion As String
   O_CentroDeCosto As String
   Centro_Descripcion As String
   O_CantidadPedida As Double
   O_CantidadPendiente As Double
   R_Precio As Double
   O_Precio As Double
   CantRecibida As Double
   O_FormaDePagoPactada As String
End Type

Public Type TipoRecepcion
    O_NumeroOrdenDeCompra As Integer
    O_CentroDeCostoEmisor As String
    O_CantidadPedida As Double
    O_CantidadPendiente As Double
    A_Codigo As Long
    A_Descripcion As String
    R_Precio As Double
    O_PrecioPactado As Double
    O_FormaDePagoPactada As String
End Type

Public Type TipoAutorizacionDePago
    O_NumeroOrdenDeContratacion As Integer
    O_CuentaContable As String
    O_CentroDeCosto As String
    O_PrecioPactado As Double
    PrecioReal As Double
    O_CentroDeCostoEmisor As String
    O_Fecha As String
    MontoSinPresupuestar As Double
End Type

Public Type TiposDeCoche
    Codigo As String
    Descripcion As String
End Type

Public Type TipoParametros
    P_ControDeCostoEmisorSeguro As String
    P_PorcentajeIVA As Double
    P_CuentaPorDefecto As String
    P_ProveedorPorDefecto As Integer
End Type

Public Type TipoRequerimiento
    CodArticulo As Long
    DescArticulo As String
    Cantidad As Double
    CantidadExtra As Double
    Numero As Integer
    Taller As Integer
    Marca As String
End Type

Public Type TipoRequerimientoCompra
    CodArticulo As Long
    DescArticulo As String
    Cantidad As Double
    CantidadPendiente As Double
    Numero As Integer
    FechaProbableDeEntrega As String
End Type

Public Type TipoFondoFijoPendiente
    R_NumeroFondoFijo As Long
    R_TotalARendir As Double
    R_Numero As Long
    R_Fecha As String
End Type

Public Type TipoFiles
    NumeroFile As Long
    Nombre As String
    Fecha As String
End Type

Public Type TipoAgrupacionRubroContrable
    A_Codigo As Integer
    A_Descripcion As String
    A_Padre As Integer
    A_Nivel As Integer
    A_Rubro As String
    ImpContable As Double
    ImpPresupuestado As Double
    ImpFinanciero As Double
    DesvioContable As Double
    DesvioFinanciero As Double
    TotELP As Double
    TotLTO As Double
    TotPPK As Double
    TotVyM As Double
    Estado As String
End Type

Public Type TipoUnidadDeNegocio
    U_Codigo As Integer
    U_Descripcion As String
    U_PorcenjateDeIva As Double
End Type

Public Type TipoDistribucionPresupuesto
    P_NumeroPresupuesto As Long
    P_Importe As Double
    P_CentroDeCostosEmisor As String
    P_SubCentroDeCosto As String
    P_CuentaContable As String
    P_Periodo As String
End Type

Public Type TipoPaquetes
    P_Descripcion As String
    P_Codigo As Integer
End Type

Public Type MarcasXArticulo
    Articulo As Integer
    Marca As String
    Descripcion As String
End Type


'variables de tipo registro
Public VecArtCompra() As TipoArticuloCompras
Public ArtCompraActual As TipoArticuloCompras
Public VecArtTaller() As TipoArticuloCompras
Public VecArtMotoVan() As TipoArticuloCompras
Public VecCentroDeCosto() As TipoCentroDeCosto
Public VecCentroDeCostoPorPadre() As TipoCentroDeCosto
Public VecCentroDeCostoNivel2() As TipoCentroDeCosto
Public CentroDeCostoActual As TipoCentroDeCosto
Public VecCentroDeCostoEmisor() As TipoCentroDeCosto
Public VecTiposDeCoche() As TiposDeCoche
Public VecLugaresDeEntrega() As TipoLugarDeEntrega
Public VecFormasDePago() As TipoFormaDePago
Public VecRubrosContables() As TipoRubroContable
Public VecUnidadesDeNegocio() As TipoUnidadDeNegocio

Public Articulos() As Long
Public Requerimientos() As TipoRequerimiento
Public Precios() As Double

Public VecUsuarios() As String
Public VecAutorizacionDePago() As TipoAutorizacionDePago
Public VecDistribucionPresupuesto() As TipoDistribucionPresupuesto
Public VecRequerimientoCompra() As TipoRequerimientoCompra
Public VecPaquetes()  As TipoPaquetes

Public ParametrosSeguro As TipoParametros
'variable de ADO
Public Conec As ADODB.Connection
Public ConecSvr As ADODB.Connection

'variables que definen al usuario
Public Usuario As String
Public NombreUsuario As String
Public N_Nivel As Integer
Public CentroEmisor As String
Public MaxSinAutorizacion As Double
Public EmailAutorizacion As String
Public VecTodasLasMarcas() As Marcas
Public VecMarcasXArticulo() As MarcasXArticulo
Public VecAutorizacionesAnticiposApli() As Long
Public VecProveedoresPorCentroDeCosto() As Proveedores

'funciones blobales
Public Sub ColorObligatorio(Componente As Control, ParamArray Controls())
   Dim C
    If Componente.Text = "" Then
    'pone el control en color crema
        Componente.BackColor = &HC0FFFF
        For Each C In Controls
            C.Enabled = False
        Next
    Else
    'pone el contro en color blanco
        Componente.BackColor = &H80000005
        For Each C In Controls
            C.Enabled = True
        Next
    End If
End Sub

Public Sub LimpiarTXT(ByVal Frm As Form)
    Dim C As Control
    For Each C In Frm
       If TypeOf C Is TextBox Then
            C.Text = ""
        End If
    Next
End Sub

Public Sub InicializarConexionADO()
    If Conec Is Nothing Then
        Dim StrConec As String
        
        Set Conec = New ADODB.Connection
        Conec.CursorLocation = adUseClient
        Conec.CommandTimeout = 1500
        Conec.Open "servidor=sql;dsn=ElPulqui1;uid=todos;PWD=todos"
        '"Provider=sqloledb; Data Source=Encomienda;Initial Catalog=ElPulquiPrueba;User Id=Todos;Password=Todos; "
        '"servidor=sql;dsn=ElPulqui1;uid=todos;PWD=todos"
        'Dim i As Integer
        'For i = 0 To Conec.Properties.Count - 1
        '   Debug.Print Conec.Properties(i).Name & " - " & Conec.Properties(i).Value
        'Next
        Set ConecSvr = New ADODB.Connection
        StrConec = "Provider=sqloledb; Data Source=svrppack.dyndns.org" & _
                    "; Initial Catalog=" & Conec.Properties("Current Catalog").Value & _
                    ";User Id=Todos;Password=todos; "
        
        ConecSvr.Open StrConec
    End If
End Sub

Public Sub AlterColor(Lv As ListView)
   Dim i As Integer
   Dim j As Integer
   For i = 1 To Lv.ListItems.Count
     If Lv.ListItems.Item(i).Index Mod 2 = 1 Then
        Lv.ListItems(i).ForeColor = vbBlue
        Lv.ListItems(i).Bold = True
        For j = 1 To Lv.ListItems.Item(i).ListSubItems.Count
            Lv.ListItems(i).ListSubItems(j).ForeColor = vbBlue
            Lv.ListItems(i).ListSubItems(j).Bold = True
        Next
        
     End If
   Next
End Sub

Public Sub CargarVecArtCompras()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpTA_ArticulosCompras"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecArtCompra(.RecordCount)
    
   While Not .EOF
        VecArtCompra(i).A_Codigo = !A_Codigo
        VecArtCompra(i).A_Descripcion = Trim(!A_Descripcion)
        VecArtCompra(i).A_CuentaPorDefecto = VerificarNulo(!A_CuentaPorDefecto)
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecArtTaller()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    
    i = 1
    Sql = "SpTaArticulos"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecArtTaller(.RecordCount)
    ReDim VecArticulos(0)
   
   While Not .EOF
        VecArtTaller(i).A_Codigo = !A_Codigo
        VecArtTaller(i).A_Descripcion = Trim(!A_Descripcion) & " (Cod. " & !A_Codigo & ")"
        VecArtTaller(i).A_CuentaPorDefecto = VerificarNulo(!A_CuentaContable)
        
        ReDim Preserve VecArticulos(UBound(VecArticulos) + 1)
        VecArticulos(UBound(VecArticulos)).Descripcion = !A_Descripcion
        VecArticulos(UBound(VecArticulos)).Codigo = !A_Codigo
        VecArticulos(UBound(VecArticulos)).UnidadDeMedida = !A_UnidadDeMedida
        VecArticulos(UBound(VecArticulos)).ImprimeEtiquetas = VerificarNulo(!A_ImprimeEtiquetas, "B")
        VecArticulos(UBound(VecArticulos)).Servicio = VerificarNulo(!A_Servicio, "B")
        VecArticulos(UBound(VecArticulos)).CuentaContable = VerificarNulo(!A_CuentaContable)
        VecArticulos(UBound(VecArticulos)).DescripcionRubro = VerificarNulo(!DescripcionRubro)
        VecArticulos(UBound(VecArticulos)).DescripcionUnidad = VerificarNulo(!DescripcionUnidad)
        VecArticulos(UBound(VecArticulos)).LlevaStock = VerificarNulo(!A_LlevaStock, "B")
        VecArticulos(UBound(VecArticulos)).Proyeccion = Val(VerificarNulo(!A_Proyeccion, "N"))
        VecArticulos(UBound(VecArticulos)).ProyectaIngresos = VerificarNulo(!A_ProyectaIngresos, "B")
        VecArticulos(UBound(VecArticulos)).AcumulaEnRendimiento = VerificarNulo(!A_AcumulaEnRendimiento, "B")
        VecArticulos(UBound(VecArticulos)).Grupo = VerificarNulo(!R_Grupo, "N")
        VecArticulos(UBound(VecArticulos)).Rubro = VerificarNulo(!A_Rubro, "N")
        VecArticulos(UBound(VecArticulos)).Reducida = VerificarNulo(!A_Reducida)
        VecArticulos(UBound(VecArticulos)).Ubicacion = VerificarNulo(!A_Ubicacion)
        VecArticulos(UBound(VecArticulos)).ServicioRecapado = VerificarNulo(!A_ServicioRecapado, "B")
        VecArticulos(UBound(VecArticulos)).GasOil = VerificarNulo(!A_GasOil, "B")
        VecArticulos(UBound(VecArticulos)).Cubierta = VerificarNulo(!A_Cubierta, "B")

        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecArtMotoVan()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    
    i = 1
    Sql = "SpOcArticulosMotoVan"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecArtMotoVan(.RecordCount)
    
   While Not .EOF
        VecArtMotoVan(i).A_Codigo = !Codigo
        VecArtMotoVan(i).A_Descripcion = Trim(!Descripcion)
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecUsuarios()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpAC_Usuarios"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecUsuarios(.RecordCount)
    
   While Not .EOF
        VecUsuarios(i) = !U_Usuario
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecCentrosDeCostos()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    'Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpTaCentrosDeCostos"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    '.CursorType = adOpenKeyset
    '.CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecCentroDeCosto(.RecordCount)
    
   While Not .EOF
        VecCentroDeCosto(i).C_Codigo = !C_Codigo
        VecCentroDeCosto(i).C_Descripcion = Convertir(!C_Descripcion)
        VecCentroDeCosto(i).C_Padre = VerificarNulo(!C_Padre)
        VecCentroDeCosto(i).C_Jerarquia = VerificarNulo(!C_Jerarquia)
        VecCentroDeCosto(i).C_Nivel = VerificarNulo(!C_Nivel, "N")
        'Debug.Print Convertir(!C_Descripcion)
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecCentrosDeCostosNivel2()
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    
    i = 1
    Sql = "SpTaCentrosDeCostosNivel2"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecCentroDeCostoNivel2(.RecordCount)
    
   While Not .EOF
        VecCentroDeCostoNivel2(i).C_Codigo = !C_Codigo
        VecCentroDeCostoNivel2(i).C_Descripcion = Convertir(!C_Descripcion)
        VecCentroDeCostoNivel2(i).C_Padre = VerificarNulo(!C_Padre)
        VecCentroDeCostoNivel2(i).C_Jerarquia = VerificarNulo(!C_Jerarquia)
        VecCentroDeCostoNivel2(i).C_Nivel = VerificarNulo(!C_Nivel, "N")

        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarCmbCentrosDeCostos(CmbCentrosDeCostos As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbCentrosDeCostos.Clear
    For i = 0 To UBound(VecCentroDeCosto)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbCentrosDeCostos.AddItem "Seleccione un Centro de Costo"
           Else
              CmbCentrosDeCostos.AddItem "Todos los Centros de Costos"
           End If
        Else
            CmbCentrosDeCostos.AddItem VecCentroDeCosto(i).C_Descripcion
        End If
    Next
        
    CmbCentrosDeCostos.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCmbCentrosDeCostosNivel2(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    For i = 0 To UBound(VecCentroDeCostoNivel2)
        If i = 0 Then
           If Tipo = "Elegir" Then
              Cmb.AddItem "Seleccione un Centro de Costo"
           Else
              Cmb.AddItem "Todos los Centros de Costos"
           End If
        Else
            Cmb.AddItem VecCentroDeCostoNivel2(i).C_Descripcion
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarCmbCentrosDeCostosEmisor(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer
    
    Cmb.Clear
    
    If Tipo = "Elegir" Then
       Cmb.AddItem "Seleccione un Centro de Costo"
    Else
       Cmb.AddItem "Todos los Centros de Costos"
    End If
    
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        Cmb.AddItem VecCentroDeCostoEmisor(i).C_Descripcion
    Next
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub BuscarCentroEmisor(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
    
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
    
End Sub

Public Function BuscarIndexCentroEmisor(C_Codigo As String) As Integer
    Dim i As Integer
    
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            BuscarIndexCentroEmisor = i
            Exit Function
        End If
    Next
    
End Function

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
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCmbUsuarios(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    For i = 0 To UBound(VecUsuarios)
        If i = 0 Then
           If Tipo = "Elegir" Then
              Cmb.AddItem "Seleccione un Usuario"
           Else
              Cmb.AddItem "Todos los Usuarios"
           End If
        Else
            Cmb.AddItem VecUsuarios(i)
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarVecCentrosDeCostosEmisor()
On Error GoTo Errores
Dim i As Integer
 Dim RsCargar As ADODB.Recordset
 Dim Sql As String
 Set RsCargar = New ADODB.Recordset
 
    Sql = "SpTaCentrosDeCostosPadres"

    With RsCargar
        .Open Sql, Conec
       ReDim VecCentroDeCostoEmisor(.RecordCount)
       VecCentroDeCostoEmisor(i).C_Descripcion = "Todos Los Centros De Costos"
       i = 1
       While Not .EOF
       
           VecCentroDeCostoEmisor(i).C_Codigo = !C_Codigo
           VecCentroDeCostoEmisor(i).C_Descripcion = Convertir(!C_Descripcion)
           VecCentroDeCostoEmisor(i).C_UsuarioResponsable = VerificarNulo(!C_UsuarioResponsable)
           VecCentroDeCostoEmisor(i).C_TablaArticulos = Trim(VerificarNulo(!C_TablaArticulos))
           VecCentroDeCostoEmisor(i).C_MontoSinPresupuestar = VerificarNulo(!C_MontoSinPresupuestar, "N")
           VecCentroDeCostoEmisor(i).C_Jerarquia = VerificarNulo(!C_Jerarquia)
           VecCentroDeCostoEmisor(i).C_Nivel = VerificarNulo(!C_Nivel, "N")
           VecCentroDeCostoEmisor(i).C_TablaRequerimientosDeCompra = VerificarNulo(!C_TablaRequerimientosDeCompra)
           VecCentroDeCostoEmisor(i).C_ImporteAutorizar = VerificarNulo(!C_ImporteAutorizar, "N")
           VecCentroDeCostoEmisor(i).C_EmailAutorizacion = VerificarNulo(!C_EmailAutorizacion)
           If LCase(VerificarNulo(!C_UsuarioResponsable)) = LCase(Usuario) Then
                CentroEmisor = VecCentroDeCostoEmisor(i).C_Codigo
           End If
            i = i + 1
          .MoveNext
       Wend
       .Close
    End With
    
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub BuscarCuentaContable(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
    i = 1
    
    While VecCuentasContables(i).Codigo <> C_Codigo
        i = i + 1
    Wend
    
    Cmb.ListIndex = i
End Sub

Public Sub BuscarCentro(C_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
    
    Cmb.ListIndex = 0
    For i = 1 To UBound(VecCentroDeCosto)
       If VecCentroDeCosto(i).C_Codigo = C_Codigo Then
            Cmb.ListIndex = i
            Exit For
       End If
    Next
End Sub

Public Sub UbicarEmpresa(E_Codigo As String, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecEmpresas)
      If Trim(VecEmpresas(i).Codigo) = Trim(E_Codigo) Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Sub UbicarCmbCentroDeCostoNivel2(Codigo As String, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecCentroDeCostoNivel2)
      If Trim(VecCentroDeCostoNivel2(i).C_Codigo) = Trim(Codigo) Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Sub UbicarUsuario(Usuario As String, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecUsuarios)
      If Trim(VecUsuarios(i)) = Trim(Usuario) Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Function BuscarDescArt(A_Codigo As Long, Tabla As String) As String
  Dim i As Integer
  If Tabla = "" Then
    For i = 1 To UBound(VecArtCompra)
        If VecArtCompra(i).A_Codigo = A_Codigo Then
            BuscarDescArt = VecArtCompra(i).A_Descripcion
            Exit Function
        End If
    Next
  Else
    If Tabla = "MotoVan" Then
        For i = 1 To UBound(VecArtMotoVan)
            If VecArtMotoVan(i).A_Codigo = A_Codigo Then
                BuscarDescArt = VecArtMotoVan(i).A_Descripcion
                Exit Function
            End If
        Next
    Else
        For i = 1 To UBound(VecArtTaller)
            If VecArtTaller(i).A_Codigo = A_Codigo Then
                BuscarDescArt = VecArtTaller(i).A_Descripcion
                Exit Function
            End If
        Next
    End If
  End If
End Function

Public Function BuscarDescCta(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCuentasContables)
        If VecCuentasContables(i).Codigo = C_Codigo Then
            BuscarDescCta = Trim(VecCuentasContables(i).Descripcion)
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescCentro(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCosto)
        If VecCentroDeCosto(i).C_Codigo = C_Codigo Then
            BuscarDescCentro = VecCentroDeCosto(i).C_Descripcion
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescProv(Codigo As Integer) As String
    Dim i As Integer
    For i = 1 To UBound(VecProveedores)
        If VecProveedores(i).Codigo = Codigo Then
            BuscarDescProv = VecProveedores(i).Descripcion
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescCentroEmisor(C_Codigo As String) As String
    Dim i As Integer
    For i = 0 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            BuscarDescCentroEmisor = VecCentroDeCostoEmisor(i).C_Descripcion
            Exit Function
        End If
    Next
    
    For i = 1 To UBound(VecCentroDeCostoNivel2)
        If VecCentroDeCostoNivel2(i).C_Codigo = C_Codigo Then
            BuscarDescCentroEmisor = BuscarDescCentroEmisor(VecCentroDeCostoNivel2(i).C_Padre)
            Exit Function
        End If
    Next

End Function

Public Function BuscarEMailAutorizacionCentroEmisor(C_Codigo As String) As String
    Dim i As Integer
    For i = 0 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            BuscarEMailAutorizacionCentroEmisor = VecCentroDeCostoEmisor(i).C_EmailAutorizacion
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescCentroEmisorPorJerarquia(C_Jerarquia As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Jerarquia = C_Jerarquia Then
            BuscarDescCentroEmisorPorJerarquia = VecCentroDeCostoEmisor(i).C_Descripcion
            Exit Function
        End If
    Next
    
    For i = 1 To UBound(VecCentroDeCostoNivel2)
        If VecCentroDeCostoNivel2(i).C_Jerarquia = C_Jerarquia Then
            BuscarDescCentroEmisorPorJerarquia = BuscarDescCentroEmisor(VecCentroDeCostoNivel2(i).C_Padre)
            Exit Function
        End If
    Next

End Function

Public Function BuscarTablaCentroEmisor(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            BuscarTablaCentroEmisor = VecCentroDeCostoEmisor(i).C_TablaArticulos
            Exit Function
        End If
    Next
End Function

Public Function BuscarCentroPadre(C_Codigo As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 1 To UBound(VecCentroDeCosto)
        If VecCentroDeCosto(i).C_Codigo = C_Codigo Then
            For j = 1 To UBound(VecCentroDeCostoNivel2)
                If VecCentroDeCostoNivel2(j).C_Jerarquia = VecCentroDeCosto(i).C_Padre Then
                    For k = 1 To UBound(VecCentroDeCostoEmisor)
                        If VecCentroDeCostoEmisor(k).C_Jerarquia = VecCentroDeCostoNivel2(j).C_Padre Then
                            BuscarCentroPadre = VecCentroDeCostoEmisor(k).C_Codigo
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    Next
    BuscarCentroPadre = 0
End Function

Public Sub TxtNumerico(Txt As TextBox, KeyAscii As Integer)
If KeyAscii <> 8 Then
 'esta variable indica donde está el "."
   Dim SepDecimal As Integer
    SepDecimal = InStr(1, Txt.Text, ".")
    
    If SepDecimal > 0 And KeyAscii = Asc(".") Then
       KeyAscii = 0
       Beep
       Exit Sub
    End If
    
    If SepDecimal > 0 Then
       If Txt.SelStart >= SepDecimal And _
          Len(Txt.Text) - SepDecimal >= 4 Then
          Beep
          KeyAscii = 0
          Exit Sub
       End If
   End If
     ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 _
       And KeyAscii <> Asc(".") Then
          Beep
          KeyAscii = 0
    End If
End If

End Sub

Public Function ValidarPeriodo(Periodo As Date, Optional MSG As Boolean = True) As Date
    Dim PeriodoCerrado As Boolean
    Dim RsValidarPeriodo As New ADODB.Recordset
    Dim Sql As String
      Sql = "SpOCCierrePeriodoValidarPeriodo @C_Periodo = '" & Format(Periodo, "MM/yyyy") & "'"
      RsValidarPeriodo.Open Sql, Conec
      PeriodoCerrado = RsValidarPeriodo!Cerrado > 0
      If PeriodoCerrado Then
        If MSG Then
           MsgBox "El príodo está cerrado", vbInformation, "Período Cerrado"
        End If
         ValidarPeriodo = DateAdd("M", 1, RsValidarPeriodo!UltimoCerrado)
      Else
        ValidarPeriodo = Periodo
      End If

End Function

Public Function NombreArchivoPDF(Nombre As String) As String
    If Mid(Nombre, Len(Nombre) - 3, 4) = ".pdf" Then
        NombreArchivoPDF = Nombre
    Else
        NombreArchivoPDF = Nombre & ".pdf"
    End If
End Function

Public Function TraerNivel(Modulo As String) As Integer
    Dim RsNivel As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SpTraerNivel @M_Modulo = '" & Modulo & "', @Usuario= '" & Usuario & "'"
    RsNivel.Open Sql, Conec, adOpenForwardOnly, adLockReadOnly
    If Not RsNivel.EOF Then
        TraerNivel = VerificarNulo(RsNivel!N_Nivel, "N")
    End If
    RsNivel.Close
    
End Function

Public Sub CentrarFormulario(Formulario As Form)
    Formulario.Top = (Screen.Height - Formulario.Height) / 2
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
End Sub

Public Function ValN(Numero) As Double
     If IsNull(Numero) Then
         ValN = 0
     Else
         ValN = Val(Replace(Numero, ",", "."))
     End If
End Function

Public Function BuscarAuxiliar(Usuario As String) As String
    Dim RsCargar As New ADODB.Recordset
    Dim Sql As String

    Sql = "SpTaAuxiliaresCentosDeCostosPorUsuario @A_Usuario='" & Usuario & "'"
    RsCargar.Open Sql, Conec, adOpenForwardOnly, adLockReadOnly
    If Not RsCargar.EOF Then
        BuscarAuxiliar = RsCargar!A_CentroDeCosto
    End If
    
    RsCargar.Close
End Function

Public Sub UbicarProveedor(Codigo As String, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecProveedores)
      If Trim(VecProveedores(i).Codigo) = Trim(Codigo) Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Function Convertir(Cadena As String) As String
        Cadena = Replace(Cadena, "¢", "ó")
        Cadena = Replace(Cadena, "¡", "í")
        Cadena = Replace(Cadena, "¤", "ñ")
        Cadena = Replace(Cadena, "‚", "é")
        Cadena = Replace(Cadena, " ", "á")
        Cadena = Replace(Cadena, "§", "º")
        Cadena = Replace(Cadena, "£", "ú")
        'Cadena = Replace(Cadena, "Tr afico", "Tráfico")
        
        Convertir = Cadena
End Function

Public Function BuscarCodigoCentro(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCosto)
        If VecCentroDeCosto(i).C_Codigo = C_Codigo Then
            BuscarCodigoCentro = Mid(VecCentroDeCosto(i).C_Jerarquia, 5, 3)
            Exit Function
        End If
    Next
End Function

Public Function BuscarJerarquiaCentro(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCostoEmisor)
        If VecCentroDeCostoEmisor(i).C_Codigo = C_Codigo Then
            BuscarJerarquiaCentro = VecCentroDeCostoEmisor(i).C_Jerarquia
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescCentroPorCodSecundario(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCosto)
        If Mid(VecCentroDeCosto(i).C_Jerarquia, 5, 3) = C_Codigo Then
            BuscarDescCentroPorCodSecundario = VecCentroDeCosto(i).C_Descripcion
            Exit Function
        End If
    Next
End Function

Public Function BuscarCodigoPorCodSecundario(C_Codigo As String) As String
    Dim i As Integer
    For i = 1 To UBound(VecCentroDeCosto)
        If Mid(VecCentroDeCosto(i).C_Jerarquia, 5, 3) = C_Codigo Then
            BuscarCodigoPorCodSecundario = VecCentroDeCosto(i).C_Codigo
            Exit Function
        End If
    Next
End Function

Public Sub TxtNumericoNeg(Txt As TextBox, KeyAscii As Integer)
If KeyAscii <> 8 Then
 'esta variable indica donde está el "."
   Dim SepDecimal As Integer
   Dim Menos As Integer
    Menos = InStr(1, Txt.Text, "-")
    SepDecimal = InStr(1, Txt.Text, ".")
    
    If SepDecimal > 0 And KeyAscii = Asc(".") Then
       KeyAscii = 0
       Beep
       Exit Sub
    End If
    If KeyAscii = Asc("-") Then
        If Txt.SelStart <> 0 Or Menos <> 0 Then
           KeyAscii = 0
           Beep
        End If
        
        Exit Sub
    End If
    
    If SepDecimal > 0 Then
       
       If Txt.SelStart >= SepDecimal And _
          Len(Txt.Text) - SepDecimal >= 4 Then
          Beep
          KeyAscii = 0
          Exit Sub
       End If
   End If
     ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 _
       And KeyAscii <> Asc(".") Then
          Beep
          KeyAscii = 0
    End If
End If

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

Public Sub TxtNumerico2(Txt As TextBox, KeyAscii As Integer)
If KeyAscii <> 8 Then
 'esta variable indica donde está el "."
   Dim SepDecimal As Integer
    SepDecimal = InStr(1, Txt.Text, ".")
    
    If SepDecimal > 0 And KeyAscii = Asc(".") Then
       KeyAscii = 0
       Beep
       Exit Sub
    End If
    
    If SepDecimal > 0 Then
       If Txt.SelStart >= SepDecimal And _
          Len(Txt.Text) - SepDecimal >= 2 Then
          Beep
          KeyAscii = 0
          Exit Sub
       End If
   End If
     ' controla que solo se ingresen números
    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And KeyAscii <> 8 _
       And KeyAscii <> Asc(".") Then
          Beep
          KeyAscii = 0
    End If
End If

End Sub

Public Sub CargarTiposDeCoche()
Dim Sql As String
Dim RsCargar As New ADODB.Recordset
    
    Sql = "SpTATiposDeCoche"
    RsCargar.Open Sql, Conec
        
    With RsCargar
        ReDim VecTiposDeCoche(0)
        With VecTiposDeCoche(0)
            .Codigo = ""
            .Descripcion = "Seleccione un Tipo de Coche"
        End With
        While Not .EOF
              ReDim Preserve VecTiposDeCoche(UBound(VecTiposDeCoche) + 1)
              VecTiposDeCoche(UBound(VecTiposDeCoche)).Descripcion = Trim(!T_Descripcion)
              VecTiposDeCoche(UBound(VecTiposDeCoche)).Codigo = !T_Codigo
              .MoveNext
        Wend
        .Close
    End With
    
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub BuscarTipoDeCoche(TipodeCoche As String, CmbTiposDeCoche As ComboEsp)
On Error GoTo Errores
'busca el punto de venta en el vector puntos de venta y lo asigna al combo asociado
Dim Encontro As Boolean
Dim i As Integer
    Encontro = False
    i = 0
    While Not Encontro And i <= UBound(VecTiposDeCoche)
        If VecTiposDeCoche(i).Codigo = TipodeCoche Then
            CmbTiposDeCoche.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function BuscarDescTipoDeCoche(TipodeCoche As String) As String
On Error GoTo Errores
'busca el punto de venta en el vector puntos de venta y lo asigna al combo asociado
Dim Encontro As Boolean
Dim i As Integer
    Encontro = False
    i = 0
    While Not Encontro And i <= UBound(VecTiposDeCoche)
        If VecTiposDeCoche(i).Codigo = TipodeCoche Then
            BuscarDescTipoDeCoche = VecTiposDeCoche(i).Descripcion
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarComboTiposDeCoche(CmbTiposDeCoche As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbTiposDeCoche.Clear
    For i = 0 To UBound(VecTiposDeCoche)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbTiposDeCoche.AddItem "Seleccione un Tipo de Coche"
           Else
              CmbTiposDeCoche.AddItem "Todas los Tipos de Coche"
           End If
       Else
            CmbTiposDeCoche.AddItem VecTiposDeCoche(i).Descripcion
       End If
    Next
    CmbTiposDeCoche.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCmbClasificacion(Cmb As ComboEsp)
    Cmb.Clear
    Cmb.AddItem "Seleccione una Clasificación"
    Cmb.AddItem "Camiones"
    Cmb.AddItem "Ómnibus"
    Cmb.AddItem "Flota Menor"
    Cmb.AddItem "Semiremolque"
    Cmb.ListIndex = 0
End Sub

Public Function BuscarDescClasificacion(Codigo As Integer) As String
   Select Case Codigo
   Case 1
       BuscarDescClasificacion = "Camiones"
   Case 2
       BuscarDescClasificacion = "Ómnibus"
   Case 3
       BuscarDescClasificacion = "Flota Menor"
   Case 4
       BuscarDescClasificacion = "Semiremolque"
   End Select

End Function

Public Sub CargarCmbCochesDominio(CmbCoches As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer
    CmbCoches.Clear
    For i = 0 To UBound(VecCoches)
        If i = 0 Then
           If Tipo = "Elegir" Then
               CmbCoches.AddItem "Sel. un Coche"
           Else
               CmbCoches.AddItem "Todos los Coches"
           End If
        Else
            CmbCoches.AddItem VecCoches(i).Codigo & " - Dominio " & VecCoches(i).Dominio
        End If
    Next
    CmbCoches.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub PosicionarCmbCoches(Codigo As Integer, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecCoches)
      If VecCoches(i).Codigo = Codigo Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Function BuscarDominioCoches(Codigo As Integer) As String
 Dim i As Integer
   For i = 1 To UBound(VecCoches)
      If VecCoches(i).Codigo = Codigo Then
           BuscarDominioCoches = VecCoches(i).Dominio
            Exit Function
      End If
   Next
End Function

Public Function BuscarSubCentroDeCostoCoche(Codigo As Integer) As String
 Dim i As Integer
   For i = 1 To UBound(VecCoches)
      If VecCoches(i).Codigo = Codigo Then
           BuscarSubCentroDeCostoCoche = VecCoches(i).SubCentroDeCosto
            Exit Function
      End If
   Next
End Function

Public Sub CargarParametrosSeguro()
Dim RsCargar As New ADODB.Recordset
    With RsCargar
        .Open "SpSegParametrosTraer", Conec
        If Not .EOF Then
            ParametrosSeguro.P_PorcentajeIVA = ValN(!P_PorcentajeIVA)
            ParametrosSeguro.P_ControDeCostoEmisorSeguro = VerificarNulo(!P_ControDeCostoEmisorSeguro)
            ParametrosSeguro.P_CuentaPorDefecto = VerificarNulo(!P_CuentaPorDefecto)
            ParametrosSeguro.P_ProveedorPorDefecto = ValN(!P_ProveedorPorDefecto)
        End If
    End With
End Sub

Public Sub CargarVecLugaresDeEntrega()
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpTaLugaresDeEntrega"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecLugaresDeEntrega(.RecordCount)
    
   While Not .EOF
        VecLugaresDeEntrega(i).L_Codigo = !L_Codigo
        VecLugaresDeEntrega(i).L_Descripcion = !L_Descripcion
        VecLugaresDeEntrega(i).L_EMail = VerificarNulo(!L_EMail)
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarCmbLugaresDeEntrega(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    For i = 0 To UBound(VecLugaresDeEntrega)
        If i = 0 Then
           If Tipo = "Elegir" Then
              Cmb.AddItem "Seleccione un Lugar de Entrega"
           Else
              Cmb.AddItem "Todos los Lugares de Entrega"
           End If
        Else
           Cmb.AddItem VecLugaresDeEntrega(i).L_Descripcion
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub UbicarCmbLugaresDeEntrega(L_Codigo As String, Cmb As ComboEsp)
    Dim i As Integer
   
    For i = 1 To UBound(VecLugaresDeEntrega)
        If VecLugaresDeEntrega(i).L_Codigo = L_Codigo Then
           Cmb.ListIndex = i
           Exit For
        End If
    Next
    
End Sub

Public Sub CargarVecFormasDePago()
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As ADODB.Recordset
    Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpTaFormasDePago"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecFormasDePago(.RecordCount)
    
   While Not .EOF
        VecFormasDePago(i).F_Codigo = !F_Codigo
        VecFormasDePago(i).F_Descripcion = !F_Descripcion
       
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarCmbFormasDePago(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    For i = 0 To UBound(VecFormasDePago)
        If i = 0 Then
           If Tipo = "Elegir" Then
              Cmb.AddItem "Seleccione una Forma de Pago"
           Else
              Cmb.AddItem "Todos las Formas de Pago"
           End If
        Else
           Cmb.AddItem VecFormasDePago(i).F_Descripcion
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub UbicarCmbFormasDePago(F_Codigo As Integer, Cmb As ComboEsp)
    Dim i As Integer
   
    For i = 1 To UBound(VecFormasDePago)
        If VecFormasDePago(i).F_Codigo = F_Codigo Then
           Cmb.ListIndex = i
           Exit For
        End If
    Next
End Sub

Public Sub CargarCmbTalleres(Cmb As ComboEsp)
    Dim i As Integer
    Cmb.Clear
    Cmb.AddItem "Seleccione Un Taller"
    
    For i = 1 To UBound(VecTalleres)
        Cmb.AddItem VecTalleres(i).Descripcion
    Next
    
    Cmb.ListIndex = 0
End Sub

Public Function BuscarDescTaller(Codigo As Integer) As String
    Dim i As Integer
   
    For i = 1 To UBound(VecTalleres)
        If VecTalleres(i).Codigo = Codigo Then
           BuscarDescTaller = VecTalleres(i).Descripcion
           Exit For
        End If
    Next
End Function

Public Sub CargarVecRubrosContables()
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    i = 1
    Sql = "SpOcRubrosContables"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecRubrosContables(.RecordCount)
    
   While Not .EOF
        VecRubrosContables(i).Codigo = !R_COD
        VecRubrosContables(i).Descripcion = !R_DES
       
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarCmbRubrosContables(Cmb As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    For i = 0 To UBound(VecRubrosContables)
        If i = 0 Then
           If Tipo = "Elegir" Then
              Cmb.AddItem "Seleccione un Rubro"
           Else
              Cmb.AddItem "Todos los Rubros"
           End If
        Else
           Cmb.AddItem VecRubrosContables(i).Descripcion
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Function NivelNodo(ByVal Node As MSComctlLib.Node) As Integer
    If Node.Parent Is Nothing Then
        NivelNodo = 1
    Else
         NivelNodo = 1 + NivelNodo(Node.Parent)
    End If
End Function

Public Sub UbicarCmbRubrosContables(Cmb As ComboEsp, Codigo As String)
On Error GoTo Errores
Dim i As Integer

    For i = 0 To UBound(VecRubrosContables)
        If VecRubrosContables(i).Codigo = Codigo Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
        
    Cmb.ListIndex = 0
Errores:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Sub CargarVecUnidadesDeNegocio()
On Error GoTo ErrorCarga
    Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    i = 1
    Sql = "SpTaUnidadesDeNegocio"
 With RsCargar
 'setea propiedades del rs para que permita obtener RecordCount
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .Open Sql, Conec
    ReDim VecUnidadesDeNegocio(.RecordCount)
    
   While Not .EOF
        VecUnidadesDeNegocio(i).U_Codigo = !U_Codigo
        VecUnidadesDeNegocio(i).U_Descripcion = !U_Descripcion
        VecUnidadesDeNegocio(i).U_PorcenjateDeIva = !U_PorcenjateDeIva
        .MoveNext
        i = i + 1
   Wend
   .Close
   
 End With
    Set RsCargar = Nothing
ErrorCarga:
    Call ManipularError(Err.Number, Err.Description)
End Sub

Public Function BuscarDescEmpresa(Codigo As String) As String
On Error GoTo Errores
Dim i As Integer

    For i = 0 To UBound(VecEmpresas)
        If Trim(Codigo) = "" Then
            BuscarDescEmpresa = "Otros"
            Exit For
        End If
        If Codigo = "OCE" Then
            BuscarDescEmpresa = "Orden de Contratación Especial"
            Exit For
        End If
        
        If VecEmpresas(i).Codigo = Codigo Then
           BuscarDescEmpresa = VecEmpresas(i).Descripcion
           Exit For
        End If
     Next
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarEmailAutorizacion()
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    
    Sql = "SpOcParametrosTraer"
    With RsCargar
        .Open Sql, Conec
        If Not .EOF Then
           EmailAutorizacion = !P_EMailAutorizacion
        End If
    End With
    Set RsCargar = Nothing
End Sub

Public Sub CargarVecPaquetes()
   Dim Sql As String
    Dim i As Integer
    Dim RsCargar As New ADODB.Recordset
    i = 1
    Sql = "SpTurPaquetes"
    With RsCargar
    'setea propiedades del rs para que permita obtener RecordCount
       .CursorType = adOpenKeyset
       .CursorLocation = adUseClient
       .Open Sql, Conec
       ReDim VecPaquetes(.RecordCount)
      While Not .EOF
           VecPaquetes(i).P_Codigo = !ID
           VecPaquetes(i).P_Descripcion = !Nombre
          .MoveNext
           i = i + 1
      Wend
      .Close
    End With
End Sub

Public Sub CargarCmbPaquetes(Cmb As ComboEsp)
    Dim i As Integer
    Cmb.Clear
    Cmb.AddItem "Seleccione Un Destino"
    
    For i = 1 To UBound(VecPaquetes)
        Cmb.AddItem VecPaquetes(i).P_Descripcion
    Next
    
    Cmb.ListIndex = 0
End Sub

Public Sub UbicarCmbPaquetes(Codigo As Integer, Cmb As ComboEsp)
 Dim i As Integer
   For i = 1 To UBound(VecPaquetes)
      If VecPaquetes(i).P_Codigo = Codigo Then
            Cmb.ListIndex = i
            Exit Sub
      End If
   Next
    Cmb.ListIndex = 0
End Sub

Public Function BuscarDescPaquete(Codigo As Integer) As String
    Dim i As Integer
    For i = 1 To UBound(VecPaquetes)
        If VecPaquetes(i).P_Codigo = Codigo Then
            BuscarDescPaquete = VecPaquetes(i).P_Descripcion
            Exit Function
        End If
    Next
End Function

Public Function BuscarDescArticulo(Codigo As Integer) As String
Dim i As Integer
    For i = 1 To UBound(VecArticulos)
        If Codigo = VecArticulos(i).Codigo Then
            BuscarDescArticulo = Trim(VecArticulos(i).Descripcion)
            Exit Function
        End If
    Next
End Function

Public Function CodigoArticuloActual(CmbArticulos As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista
    CodigoArticuloActual = VecArticulos(CmbArticulos.ListIndex).Codigo
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Sub CargarComboMarcasArticulo(Articulo As Integer, CmbMarcas As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    'MousePointer = vbHourglass
    
    CmbMarcas.Clear
    ReDim VecMarcas(0)
    For i = 0 To UBound(VecMarcasXArticulo)
        If i = 0 Then
           VecMarcas(0).Codigo = ""
           If Tipo = "Elegir" Then
              VecMarcas(0).Descripcion = "Seleccione una Marca"
              CmbMarcas.AddItem "Seleccione una Marca"
           Else
              VecMarcas(0).Descripcion = "Todas las Marcas"
              CmbMarcas.AddItem "Todas las Marcas"
           End If
        Else
            If VecMarcasXArticulo(i).Articulo = Articulo Then
               CmbMarcas.AddItem VecMarcasXArticulo(i).Descripcion
               ReDim Preserve VecMarcas(UBound(VecMarcas) + 1)
               VecMarcas(UBound(VecMarcas)).Descripcion = VecMarcasXArticulo(i).Descripcion
               VecMarcas(UBound(VecMarcas)).Codigo = VecMarcasXArticulo(i).Marca
            End If
        End If
    Next
   
    CmbMarcas.ListIndex = 0
   'MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub BuscarMarca(Codigo As String, CmbMarcas As ComboEsp)
Dim Encontro As Boolean
Dim i As Integer
    Encontro = False
    i = 1
    While Not Encontro And i <= UBound(VecMarcas)
        If UCase(Trim(VecMarcas(i).Codigo)) = UCase(Trim(Codigo)) Then
            Encontro = True
            CmbMarcas.ListIndex = i
        End If
        Inc i
    Wend
End Sub

Public Function BuscarDescProveedor(Proveedor As Integer) As String
On Error GoTo Errores
'busca el punto de venta en el vector puntos de venta y lo asigna al combo asociado
Dim i As Integer
    i = 0
    While i <= UBound(VecProveedores)
        If VecProveedores(i).Codigo = Proveedor Then
            BuscarDescProveedor = VecProveedores(i).Descripcion
            Exit Function
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarCmbCentrosDeCostosPorPadre(CmbCentrosDeCostos As ComboEsp, Padre As String)
On Error GoTo Errores
Dim i As Integer

    CmbCentrosDeCostos.Clear
    CmbCentrosDeCostos.AddItem "Todos los Centros de Costos"
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    'Set RsCargar = New ADODB.Recordset
    i = 1
    Sql = "SpTaCentrosDeCostosPorPadre @PAdre ='" & Padre & "'"
    
    With RsCargar
        .Open Sql, Conec
        ReDim VecCentroDeCostoPorPadre(.RecordCount)
        
       While Not .EOF
            VecCentroDeCostoPorPadre(i).C_Codigo = !C_Codigo
            VecCentroDeCostoPorPadre(i).C_Descripcion = Convertir(!C_Descripcion)
            VecCentroDeCostoPorPadre(i).C_Padre = VerificarNulo(!C_Padre)
            VecCentroDeCostoPorPadre(i).C_Jerarquia = VerificarNulo(!C_Jerarquia)
            VecCentroDeCostoPorPadre(i).C_Nivel = VerificarNulo(!C_Nivel, "N")
            CmbCentrosDeCostos.AddItem VecCentroDeCostoPorPadre(i).C_Descripcion
    
            .MoveNext
            i = i + 1
       Wend
   End With
        
    CmbCentrosDeCostos.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCmbProveedoresPorCentroDeCosto(Cmb As ComboEsp, Centro As String)
On Error GoTo Errores
Dim i As Integer

    Cmb.Clear
    Cmb.AddItem "Todos los Pendientes"
    Dim Sql As String
    Dim RsCargar As New ADODB.Recordset
    i = 1
    Sql = "SpTaProveeroresPorCentroDeCosto @CentroDeCosto='" & Centro & "'"
    'If Not IsArray(VecProveedoresPorCentroDeCosto) Then
        With RsCargar
            .Open Sql, Conec
            ReDim VecProveedoresPorCentroDeCosto(.RecordCount)
             
            While Not .EOF
                 VecProveedoresPorCentroDeCosto(i).Codigo = !P_Codigo
                 VecProveedoresPorCentroDeCosto(i).Descripcion = Convertir(!P_Descripcion)
                 .MoveNext
                 i = i + 1
            Wend
        End With
    'End If
   
   For i = 1 To UBound(VecProveedoresPorCentroDeCosto)
      Cmb.AddItem VecProveedoresPorCentroDeCosto(i).Descripcion
   Next
   Cmb.ListIndex = 0
   
Errores:
    ManipularError Err.Number, Err.Description
End Sub


