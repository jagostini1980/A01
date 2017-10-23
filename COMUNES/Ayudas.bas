Attribute VB_Name = "Ayudas"
Option Explicit

Public Type TipoArticulos
    Codigo As Integer
    Descripcion As String
    UnidadDeMedida As String
    ImprimeEtiquetas As Boolean
    Servicio As Boolean
    Reducida As String
    Rubro As Integer
    Ubicacion As String
    Proyeccion As Double
    ProyectaIngresos As Boolean
    AcumulaEnRendimiento As Boolean
    CuentaContable As String
    LlevaStock As Boolean
    DescripcionRubro As String
    DescripcionUnidad As String
    Grupo As String
    GasOil As Boolean
    ServicioRecapado As Boolean
    Cubierta As Boolean
End Type

Public Type Proveedores
    Codigo As Long
    Descripcion As String
    CUIT As String
    Calificacion As Integer
End Type

Public Type Empresas
    Codigo As String
    Descripcion As String
    CUIT As String
End Type

Public Type CuentasContables
    Codigo As String
    Descripcion As String
    P_PADRE As String
    P_NIVEL As Integer
    P_IMP As String
    P_JER As String
End Type

Public Type TipoCoche
    Codigo As String
    Descripcion As String
    Dominio As String
    NroMotor As String
    MarcaMotor As String
    NroChassis As String
    Año As Integer
    TipoCoche As String
    Clasificacion As Integer
    SubCentroDeCosto As String
    CantidadDeAsientos As Integer
End Type

Public Type Rubros
    Codigo As Integer
    Descripcion As String
End Type

Public Type Grupos
    Codigo As Integer
    Descripcion As String
End Type

Public Type UnidadesDeMedida
    Codigo As String
    Descripcion As String
End Type

Public Type Marcas
    Codigo As String
    Descripcion As String
End Type

Public Type Talleres
    Codigo As Integer
    Descripcion As String
    EMail As String
End Type

Public VecTodasLasMarcas() As Marcas
Public VecMarcas() As Marcas
Public VecArticulos() As TipoArticulos
Public VecProveedores() As Proveedores
Public VecCuentasContables() As CuentasContables
Public VecCuentasContablesArbol() As CuentasContables
Public VecEmpresas() As Empresas
Public VecCoches() As TipoCoche
Public VecRubros() As Rubros
Public VecGrupos() As Grupos
Public VecUnidadesDeMedida() As UnidadesDeMedida
Public Articulo As TipoArticulos
Public VecTalleres() As Talleres

Public Sub CargarArticulos(db As rdoConnection)
Dim sSQL As String
    Dim TbTabla As New ADODB.Recordset
'On Error GoTo Errores
    sSQL = "SpTAArticulos"
    TbTabla.Open sSQL, Conec
    With TbTabla
        ReDim VecArticulos(0)
        ReDim VecArtTaller(0)
   

        VecArticulos(0).Codigo = "0"
        VecArticulos(0).Descripcion = "Seleccione un Artículo"
        While Not .EOF
              VecArtTaller(UBound(VecArtTaller)).A_Codigo = !A_Codigo
              VecArtTaller(UBound(VecArtTaller)).A_Descripcion = Trim(!A_Descripcion)

              ReDim Preserve VecArticulos(UBound(VecArticulos) + 1)
              VecArticulos(UBound(VecArticulos)).Descripcion = !A_Descripcion
              VecArticulos(UBound(VecArticulos)).Codigo = !A_Codigo
              VecArticulos(UBound(VecArticulos)).UnidadDeMedida = !A_UnidadDeMedida
              VecArticulos(UBound(VecArticulos)).ImprimeEtiquetas = VerificarNulo(!A_ImprimeEtiquetas, "B")
              .MoveNext
        Wend
    End With
    TbTabla.Close

Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarMarcasPorArticulo(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargo Marcas Por Articulo
    
    sSQL = "SpTAArticulosMarcasCombo"
    Tabla.Open sSQL, Conec
   
    With Tabla
        ReDim VecMarcasXArticulo(0)
        VecMarcasXArticulo(0).Marca = ""
        VecMarcasXArticulo(0).Descripcion = "Seleccione una Marca"
        While Not .EOF
            ReDim Preserve VecMarcasXArticulo(UBound(VecMarcasXArticulo) + 1)
            'estos son los datos que guardo en el vector
            VecMarcasXArticulo(UBound(VecMarcasXArticulo)).Descripcion = !Descripcion
            VecMarcasXArticulo(UBound(VecMarcasXArticulo)).Marca = !Marca
            VecMarcasXArticulo(UBound(VecMarcasXArticulo)).Articulo = !Articulo
            .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarGrupos(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargo Grupos
    
    sSQL = "SpTAGrupos"
    Tabla.Open sSQL, Conec
    With Tabla
        ReDim VecGrupos(0)
        VecGrupos(0).Codigo = "0"
        VecGrupos(0).Descripcion = "Seleccione un Grupo"
        While Not .EOF
            ReDim Preserve VecGrupos(UBound(VecGrupos) + 1)
            VecGrupos(UBound(VecGrupos)).Descripcion = !G_descripcion
            VecGrupos(UBound(VecGrupos)).Codigo = !G_codigo
            .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarUnidadesDeMedida(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargo Unidades de Medida
    sSQL = "SpTAUnidadesDeMedida"
    Tabla.Open sSQL, Conec
   
    With Tabla
        ReDim VecUnidadesDeMedida(0)
        VecUnidadesDeMedida(0).Codigo = "0"
        VecUnidadesDeMedida(0).Descripcion = "Seleccione una Unidad de Medida"
        While Not .EOF
            ReDim Preserve VecUnidadesDeMedida(UBound(VecUnidadesDeMedida) + 1)
            VecUnidadesDeMedida(UBound(VecUnidadesDeMedida)).Descripcion = !U_Descripcion
            VecUnidadesDeMedida(UBound(VecUnidadesDeMedida)).Codigo = !U_Codigo
            .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarMarcas(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    Dim pos As Integer
    'Cargo Marcas
    sSQL = "SpTAMarcas"
    Tabla.Open sSQL, Conec
    
    With Tabla
        ReDim VecMarcas(0)
        ReDim VecTodasLasMarcas(0)
        VecMarcas(0).Codigo = 0
        VecTodasLasMarcas(0).Codigo = 0
        VecMarcas(0).Descripcion = "Seleccione una Marca"
        VecTodasLasMarcas(0).Descripcion = "Seleccione una Marca"
        While Not .EOF
            pos = UBound(VecMarcas) + 1
            pos = UBound(VecTodasLasMarcas) + 1
            ReDim Preserve VecMarcas(pos)
            ReDim Preserve VecTodasLasMarcas(pos)
            VecMarcas(pos).Codigo = !M_Codigo
            VecTodasLasMarcas(pos).Codigo = !M_Codigo
            VecMarcas(pos).Descripcion = !M_Descripcion
            VecTodasLasMarcas(pos).Descripcion = !M_Descripcion
            .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarTalleres(db As rdoConnection)
    Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargo Talleres
    sSQL = "SpTATalleres"
    Tabla.Open sSQL, Conec
        
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecTalleres(0)
        With VecTalleres(0)
            .Codigo = "0"
            .Descripcion = "Seleccione un Taller"
        End With
        While Not .EOF
              ReDim Preserve VecTalleres(UBound(VecTalleres) + 1)
              VecTalleres(UBound(VecTalleres)).Descripcion = !T_Descripcion
              VecTalleres(UBound(VecTalleres)).Codigo = !T_Codigo
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarProveedores(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargar Proveedores
    sSQL = "SpOcProveedores"
     Tabla.Open sSQL, Conec
        
    With Tabla
        ReDim VecProveedores(0)
        With VecProveedores(0)
            .Codigo = "0"
            .Descripcion = "Seleccione un Proveedor"
        End With
        While Not .EOF
              ReDim Preserve VecProveedores(UBound(VecProveedores) + 1)
              VecProveedores(UBound(VecProveedores)).Descripcion = Trim(!P_Descripcion)
              VecProveedores(UBound(VecProveedores)).Codigo = !P_Codigo
              VecProveedores(UBound(VecProveedores)).CUIT = !P_Cuit
              VecProveedores(UBound(VecProveedores)).Calificacion = !E_Calificacion
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarDepositos(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset

    'Cargo Depositos
    sSQL = "SpTADepositos"
    Tabla.Open sSQL, Conec
        
    With Tabla
        ReDim VecDepositosTodos(0)
        With VecDepositosTodos(0)
            .Codigo = "0"
            .Descripcion = "Seleccione un Deposito"
        End With
        While Not .EOF
              ReDim Preserve VecDepositosTodos(UBound(VecDepositosTodos) + 1)
              VecDepositosTodos(UBound(VecDepositosTodos)).Descripcion = !D_Descripcion
              VecDepositosTodos(UBound(VecDepositosTodos)).Codigo = !D_Codigo
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarDepositosPorTaller(db As rdoConnection)
Dim sSQL As String
     Dim Tabla As New ADODB.Recordset

    'Cargo Depositos x Taller
    sSQL = "SpTATalleresDepositosCombo"
    Tabla.Open sSQL, Conec
        
    With Tabla
        ReDim VecDepositosXTaller(0)
        With VecDepositosXTaller(0)
            .Deposito = "0"
            .Descripcion = "Seleccione un Deposito"
        End With
        While Not .EOF
              ReDim Preserve VecDepositosXTaller(UBound(VecDepositosXTaller) + 1)
              VecDepositosXTaller(UBound(VecDepositosXTaller)).Descripcion = !Descripcion
              VecDepositosXTaller(UBound(VecDepositosXTaller)).Deposito = !Codigo
              VecDepositosXTaller(UBound(VecDepositosXTaller)).Taller = !CodigoTaller
              VecDepositosXTaller(UBound(VecDepositosXTaller)).Stock = ValorSINO(!Stock)
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCoches()
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset

    'Cargo Coches
    sSQL = "SpTACoches"
    Tabla.Open sSQL, Conec
        
    With Tabla
        ReDim VecCoches(0)
        With VecCoches(0)
            .Codigo = "0"
            .Descripcion = "Seleccione un Coche"
        End With
        While Not .EOF
           If VerificarNulo(!C_Activo, "B") Then
                ReDim Preserve VecCoches(UBound(VecCoches) + 1)
                VecCoches(UBound(VecCoches)).Descripcion = !C_Interno
                VecCoches(UBound(VecCoches)).Codigo = !C_Interno
                VecCoches(UBound(VecCoches)).Dominio = !C_Dominio
                VecCoches(UBound(VecCoches)).NroMotor = VerificarNulo(!C_NumeroDeMotor)
                VecCoches(UBound(VecCoches)).MarcaMotor = VerificarNulo(!C_MarcaMotor)
                VecCoches(UBound(VecCoches)).NroChassis = VerificarNulo(!C_NumeroDeChasis)
                VecCoches(UBound(VecCoches)).Año = VerificarNulo(!C_Ano, "N")
                VecCoches(UBound(VecCoches)).CantidadDeAsientos = VerificarNulo(!C_CantidadDeAsientos, "N")
                VecCoches(UBound(VecCoches)).TipoCoche = VerificarNulo(!C_TipoDeCoche, "N")
                VecCoches(UBound(VecCoches)).Clasificacion = VerificarNulo(!C_Clasificacion, "N")
                VecCoches(UBound(VecCoches)).SubCentroDeCosto = VerificarNulo(!C_CentroDeCosto)
            End If
            .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarRubros(db As rdoConnection)
Dim sSQL As String
     Dim Tabla As New ADODB.Recordset

    'Cargar Rubros
    sSQL = "SpTARubros"
    Tabla.Open sSQL, Conec
        
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecRubros(0)
        With VecRubros(0)
            .Codigo = "0"
            .Descripcion = "Seleccione un Rubro"
        End With
        While Not .EOF
              ReDim Preserve VecRubros(UBound(VecRubros) + 1)
              VecRubros(UBound(VecRubros)).Descripcion = !R_Descripcion
              VecRubros(UBound(VecRubros)).Codigo = !R_Codigo
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub
Public Sub CargarLineas(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    
    'Cargar Lineas
    sSQL = "SpTALineas"
    Tabla.Open sSQL, Conec
        
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecLineas(0)
        With VecLineas(0)
            .Codigo = ""
            .Descripcion = "Seleccione una Linea"
        End With
        While Not .EOF
              ReDim Preserve VecLineas(UBound(VecLineas) + 1)
              VecLineas(UBound(VecLineas)).Descripcion = !L_Descripcion
              VecLineas(UBound(VecLineas)).Codigo = !L_Codigo
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarCuentasContables(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargar Medidas de Cubiertas
    sSQL = "SpOcCuentasContables"
    Tabla.Open sSQL, Conec
   
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecCuentasContablesArbol(0)
        While Not .EOF
            ReDim Preserve VecCuentasContablesArbol(UBound(VecCuentasContablesArbol) + 1)
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).Descripcion = Convertir(Trim(!C_Descripcion))
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).Codigo = !C_Codigo
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).P_NIVEL = !P_NIVEL
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).P_PADRE = !P_PADRE
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).P_IMP = !P_IMP
            VecCuentasContablesArbol(UBound(VecCuentasContablesArbol)).P_JER = !P_JER
            
          .MoveNext
        Wend
        
        .MoveFirst
        .Sort = "C_Descripcion"
            
        ReDim VecCuentasContables(0)
        
        VecCuentasContables(0).Codigo = "0"
        VecCuentasContables(0).Descripcion = "Seleccione una Cuenta Contable"
        While Not .EOF
            ReDim Preserve VecCuentasContables(UBound(VecCuentasContables) + 1)
            VecCuentasContables(UBound(VecCuentasContables)).Descripcion = Convertir(Trim(!C_Descripcion))
            VecCuentasContables(UBound(VecCuentasContables)).Codigo = !C_Codigo
            VecCuentasContables(UBound(VecCuentasContables)).P_NIVEL = !P_NIVEL
            VecCuentasContables(UBound(VecCuentasContables)).P_PADRE = !P_PADRE
            VecCuentasContables(UBound(VecCuentasContables)).P_IMP = !P_IMP
            VecCuentasContables(UBound(VecCuentasContables)).P_JER = !P_JER
            
          .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub BuscarProveedor(Proveedor As Integer, CmbProveedores As ComboEsp)
On Error GoTo Errores
'busca el punto de venta en el vector puntos de venta y lo asigna al combo asociado
Dim Encontro As Boolean
Dim i As Integer
    Encontro = False
    i = 0
    While Not Encontro And i <= UBound(VecProveedores)
        If VecProveedores(i).Codigo = Proveedor Then
            CmbProveedores.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function CodigoProveedorActual(CmbProveedores)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista
    CodigoProveedorActual = VecProveedores(CmbProveedores.ListIndex).Codigo
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function DescripcionProveedorActual(CmbProveedores)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista
    DescripcionProveedorActual = VecProveedores(CmbProveedores.ListIndex).Descripcion
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Sub CargarComboProveedores(CmbProveedores As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    'MousePointer = vbHourglass

    CmbProveedores.Clear
    For i = 0 To UBound(VecProveedores)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbProveedores.AddItem "Seleccione un Proveedor"
           Else
              CmbProveedores.AddItem "Todos los Proveedores"
           End If
        Else
            CmbProveedores.AddItem VecProveedores(i).Descripcion
        End If
    Next
    CmbProveedores.ListIndex = 0
   ' MousePointer = vbNormal
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function PosicionarComboCuentasContables(ValorABuscar As String, CbCuentasContables As ComboEsp)
On Error GoTo Errores
Dim i As Integer
Dim Encontro As Boolean
    i = 1
    Encontro = False
    While Not Encontro And i <= UBound(VecCuentasContables)
        If Trim(VecCuentasContables(i).Codigo) = Trim(ValorABuscar) Then
            CbCuentasContables.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
    If Not Encontro Then
       CbCuentasContables.ListIndex = 0
    End If
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarComboCuentasContables(CbCuentasContables As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CbCuentasContables.Clear
    For i = 0 To UBound(VecCuentasContables)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CbCuentasContables.AddItem "Seleccione una Cuenta Contable"
           Else
              CbCuentasContables.AddItem "Todas las Cuentas Contables"
           End If
        Else
            CbCuentasContables.AddItem Trim(VecCuentasContables(i).Descripcion) & " (Cod. " & VecCuentasContables(i).Codigo & ")"
        End If
    Next
        
    CbCuentasContables.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarComboEmpresas(CbEmpresas As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CbEmpresas.Clear
    For i = 0 To UBound(VecEmpresas)
        If i = 0 Then
            If Tipo = "Elegir" Then
                CbEmpresas.AddItem "Seleccione una Empresa"
            
            Else
                If Tipo = "Sin" Then
                    CbEmpresas.AddItem "Otros"
                Else
                    CbEmpresas.AddItem "Todos las Empresas"
                End If
            End If
        Else
            CbEmpresas.AddItem VecEmpresas(i).Descripcion
        End If
     Next
    CbEmpresas.ListIndex = 0
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarEmpresas(db As rdoConnection)
Dim sSQL As String
    Dim Tabla As New ADODB.Recordset
    'Cargar Lineas
    sSQL = "SpTAEmpresas"
    Tabla.Open sSQL, Conec
        
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecEmpresas(0)
        With VecEmpresas(0)
            .Codigo = ""
            .Descripcion = "Seleccione una Empresa"
        End With
        While Not .EOF
              ReDim Preserve VecEmpresas(UBound(VecEmpresas) + 1)
              VecEmpresas(UBound(VecEmpresas)).Descripcion = Trim(!E_Descripcion)
              VecEmpresas(UBound(VecEmpresas)).Codigo = !E_Codigo
              VecEmpresas(UBound(VecEmpresas)).CUIT = VerificarNulo(!E_CUIT)
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function CodigoCuentaContableActual(CbCuentasContables As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista

    CodigoCuentaContableActual = VecCuentasContables(CbCuentasContables.ListIndex).Codigo
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Function DescripcionCuentaContableActual(CbCuentasContables As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista

    DescripcionCuentaContableActual = VecCuentasContables(CbCuentasContables.ListIndex).Descripcion
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarComboRubros(CmbRubros As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbRubros.Clear
    For i = 0 To UBound(VecRubros)
        If i = 0 Then
           If Tipo = "Elegir" Then
               CmbRubros.AddItem "Seleccione un Rubro"
           Else
               CmbRubros.AddItem "Todos los Rubros"
            End If
        Else
            CmbRubros.AddItem VecRubros(i).Descripcion
        End If
    Next
    CmbRubros.ListIndex = 0

Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarComboGrupos(CbGrupos As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CbGrupos.Clear
    For i = 0 To UBound(VecGrupos)
        If i = 0 Then
            If Tipo = "Elegir" Then
                CbGrupos.AddItem "Seleccione un Grupo"
            Else
                CbGrupos.AddItem "Todos los Grupos"
            End If
        Else
            CbGrupos.AddItem VecGrupos(i).Descripcion
        End If
     Next
    CbGrupos.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub CargarComboUnidadesDeMedida(CbUnidadesDeMedida As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CbUnidadesDeMedida.Clear
    For i = 0 To UBound(VecUnidadesDeMedida)
        If i = 0 Then
           If Tipo = "Elegir" Then
               CbUnidadesDeMedida.AddItem "Seleccione una Unidad de Medida"
           Else
               CbUnidadesDeMedida.AddItem "Todas las Unidades de Medida"
           End If
        Else
            CbUnidadesDeMedida.AddItem VecUnidadesDeMedida(i).Descripcion
        End If
    Next
    CbUnidadesDeMedida.ListIndex = 0

Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function PosicionarComboGrupos(ValorABuscar As String, CbGrupos As ComboEsp)
On Error GoTo Errores
Dim i As Integer
Dim Encontro As Boolean
    i = 1
    Encontro = False
    While Not Encontro
        If Trim(VecGrupos(i).Codigo) = Trim(ValorABuscar) Then
            CbGrupos.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarComboRubrosDeGrupo(db As rdoConnection, CbRubros As ComboEsp, Grupo As Integer, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim Tabla As rdoResultset
Dim sSQL As String

    'es la consulta para recuperar los datos a cargar en el combo y el vecto
    sSQL = "SpTARubrosXGrupoTraer @Codigo=" & Grupo
    Set Tabla = db.OpenResultset(sSQL)
   
    With Tabla
        'y empiezo a cargar el combo  y el vector
        ReDim VecRubros(0)
        VecRubros(0).Codigo = "0"
        If Tipo = "Elegir" Then
            VecRubros(0).Descripcion = "Seleccione un Rubro"
        Else
            VecRubros(0).Descripcion = "Todos los Rubros"
        End If
        CbRubros.Clear
        CbRubros.AddItem VecRubros(0).Descripcion
        While Not .EOF
            'esto es lo que voy a mostrar en el combo
            CbRubros.AddItem .rdoColumns("R_Descripcion").Value
            'redim preserve es para que guarde los datos anteriores a la redimensión del vector
            ReDim Preserve VecRubros(UBound(VecRubros) + 1)
            'estos son los datos que guardo en el vector
            VecRubros(UBound(VecRubros)).Descripcion = .rdoColumns("R_descripcion").Value
            VecRubros(UBound(VecRubros)).Codigo = .rdoColumns("R_codigo").Value
    
            .MoveNext
        Wend
    End With
    CbRubros.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function PosicionarComboRubros(ValorABuscar As String, CbRubros As ComboEsp)
On Error GoTo Errores
Dim i As Integer
Dim Encontro As Boolean
    i = 1
    Encontro = False
    While Not Encontro And i <= UBound(VecRubros)
        If Trim(VecRubros(i).Codigo) = Trim(ValorABuscar) Then
            CbRubros.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
    If Encontro = False Then
        CbRubros.ListIndex = 0
    End If
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Function PosicionarComboUnidadesDeMedida(ValorABuscar As String, CbUnidadesDeMedida As ComboEsp)
On Error GoTo Errores
Dim i As Integer
Dim Encontro As Boolean
    i = 1
    Encontro = False
    While Not Encontro And i <= UBound(VecUnidadesDeMedida)
        If Trim(VecUnidadesDeMedida(i).Codigo) = Trim(ValorABuscar) Then
            CbUnidadesDeMedida.ListIndex = i
            Encontro = True
        End If
        i = i + 1
    Wend
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Function BuscarDescCuentaContable(Codigo As String) As String
Dim i As Integer
    For i = 1 To UBound(VecCuentasContables)
        If Codigo = VecCuentasContables(i).Codigo Then
            BuscarDescCuentaContable = Trim(VecCuentasContables(i).Descripcion)
            Exit Function
        End If
    Next
End Function

Public Function CodigoUnidadDeMedidaActual(CbUnidadesDeMedida As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista

    CodigoUnidadDeMedidaActual = VecUnidadesDeMedida(CbUnidadesDeMedida.ListIndex).Codigo
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Function DescripcionUnidadDeMedidaActual(CbUnidadesDeMedida As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista

    DescripcionUnidadDeMedidaActual = VecUnidadesDeMedida(CbUnidadesDeMedida.ListIndex).Descripcion
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Function CodigoRubroActual(CbRubros As ComboEsp)
On Error GoTo Errores
'esta funcion es para comodidad, habría que repetirla por cada combo que exista

    CodigoRubroActual = VecRubros(CbRubros.ListIndex).Codigo
    
Errores:
    ManipularError Err.Number, Err.Description
End Function

Public Sub CargarComboMarcas(CmbMarcas As ComboEsp, Optional Tipo As String = "Elegir")
Dim i As Integer

    CmbMarcas.Clear
    For i = 0 To UBound(VecTodasLasMarcas)
        If i = 0 Then
            If Tipo = "Elegir" Then
                CmbMarcas.AddItem "Seleccione una Marca"
            Else
                CmbMarcas.AddItem "Todas las Marcas"
            End If
        Else
             CmbMarcas.AddItem VecTodasLasMarcas(i).Descripcion
        End If
    Next
    CmbMarcas.ListIndex = 0
End Sub

Public Function CodigoTodasLasMarcasActual(CmbMarcas As ComboEsp)
    CodigoTodasLasMarcasActual = VecTodasLasMarcas(CmbMarcas.ListIndex).Codigo
End Function

Public Function DescripcionTodasLasMarcasActual(CmbMarcas As ComboEsp)
    DescripcionTodasLasMarcasActual = VecTodasLasMarcas(CmbMarcas.ListIndex).Descripcion
End Function

Public Sub CargarComboArticulos(CmbArticulos As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim p As Integer
Dim Tabla As rdoResultset
Dim sSQL As String

    CmbArticulos.Clear
    For p = 0 To UBound(VecArticulos)
        If p = 0 Then
           If Tipo = "Elegir" Then
              CmbArticulos.AddItem "Seleccione un Articulo"
           Else
              CmbArticulos.AddItem "Todos los Articulos"
           End If
        Else
           CmbArticulos.AddItem VecArticulos(p).Descripcion
        End If
    Next
    CmbArticulos.ListIndex = 0

Errores:
    ManipularError Err.Number, Err.Description
End Sub

