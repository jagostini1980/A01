Attribute VB_Name = "FuncionesMcLineas"
Option Explicit

'declaración de estructuras
Public Type TiposProvMinicenas
   P_Codigo As Integer
   P_Descripcion As String
   P_Precio As Double
   P_Activo As Boolean
   P_PrecioNeto As Double
   P_ImporteIva As Double
   P_Referencias As String
   P_FechaActualizacion As String
   P_BonificaChoferes As Boolean
   P_Observaciones As String
End Type

Public Type Lineas
    Codigo As String
    Descripcion As String
End Type

Public Type TipoPresServicioDeBar
    D_NumeroPresupuesto As Long
    D_Cuenta As String
    D_Linea As String
    D_Proveedor As Integer
    D_CentroDeCostosEmisor As String
    D_Cantidad As Integer
    D_PrecioUnitario As Double
    D_Periodo As String
End Type

Public VecProvMinicenas() As TiposProvMinicenas
Public VecLineas() As Lineas
Public VecPresServicioDeBar() As TipoPresServicioDeBar

Public Sub CargarVecLineas()
Dim Sql As String
Dim Tabla As New ADODB.Recordset
    
    'Cargar Lineas
    Sql = "SpTALineas"
        
    With Tabla
        .Open Sql, Conec
        'y empiezo a cargar el combo  y el vector
        ReDim VecLineas(0)
        With VecLineas(0)
            .Codigo = ""
            .Descripcion = "Seleccione una Linea"
        End With
        While Not .EOF
              ReDim Preserve VecLineas(UBound(VecLineas) + 1)
              VecLineas(UBound(VecLineas)).Descripcion = Trim(!L_Descripcion)
              VecLineas(UBound(VecLineas)).Codigo = !L_Codigo
              .MoveNext
        Wend
    End With
    Tabla.Close
    Set Tabla = Nothing
    
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Sub UbicarCmbLineas(Cmb As ComboEsp, Codigo As Integer)
    Dim i As Integer
    For i = 1 To UBound(VecLineas)
        If VecLineas(i).Codigo = Codigo Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
    Cmb.ListIndex = 0
End Sub

Public Sub CargarComboLineas(CmbLineas As ComboEsp, Optional Tipo As String = "Elegir")
On Error GoTo Errores
Dim i As Integer

    CmbLineas.Clear
    For i = 0 To UBound(VecLineas)
        If i = 0 Then
           If Tipo = "Elegir" Then
              CmbLineas.AddItem "Seleccione una Linea"
           Else
              CmbLineas.AddItem "Todas las lineas"
           End If
       Else
            CmbLineas.AddItem VecLineas(i).Descripcion
       End If
    Next
    CmbLineas.ListIndex = 0
Errores:
    ManipularError Err.Number, Err.Description
End Sub

Public Function BuscarDescLineas(Codigo As String) As String
Dim i As Integer
    For i = 1 To UBound(VecLineas)
        If Codigo = VecLineas(i).Codigo Then
            BuscarDescLineas = Trim(VecLineas(i).Descripcion)
            Exit Function
        End If
    Next
End Function

Public Sub CargarVecProvMinicenas()
    Dim RsProv As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
    
    Sql = "SpTA_ProveedoresMinicenas"
    RsProv.CursorLocation = adUseClient
    RsProv.CursorType = adOpenKeyset
    RsProv.LockType = adLockOptimistic
    RsProv.Open Sql, Conec
 ' On Error Resume Next
  i = 1
  
  ReDim VecProvMinicenas(RsProv.RecordCount)
   VecProvMinicenas(0).P_Codigo = 0
   VecProvMinicenas(0).P_Descripcion = "Seleccione un Proveedor"
   VecProvMinicenas(0).P_Precio = 0

    With RsProv
        While Not .EOF
           'carga el vector
           VecProvMinicenas(i).P_Codigo = !P_Codigo
           VecProvMinicenas(i).P_Descripcion = !P_Descripcion
           VecProvMinicenas(i).P_Precio = !P_Precio
           VecProvMinicenas(i).P_Activo = !P_Activo
           VecProvMinicenas(i).P_ImporteIva = ValN(!P_ImporteIva)
           VecProvMinicenas(i).P_Referencias = VerificarNulo(!P_Referencias)
           VecProvMinicenas(i).P_FechaActualizacion = VerificarNulo(!P_FechaActualizacion)
           VecProvMinicenas(i).P_BonificaChoferes = VerificarNulo(!P_BonificaChoferes, "B")
           VecProvMinicenas(i).P_Observaciones = VerificarNulo(!P_Observaciones)
            i = i + 1
           .MoveNext
        Wend
    End With
    RsProv.Close
     Set RsProv = Nothing

End Sub

Public Sub CargarCmbProv(CmbProv As ComboEsp)
 'este procedimiento tiene la función de cargar el combobox de proveedores
 'de minicenas
    
    Call CargarVecProvMinicenas
    Dim i As Integer
    CmbProv.Clear
    
    For i = 0 To UBound(VecProvMinicenas)
        CmbProv.AddItem VecProvMinicenas(i).P_Descripcion
    Next
    
    If CmbProv.ListCount > 0 Then
        CmbProv.ListIndex = 0
    End If
End Sub

Public Function BuscarDescProvMinicena(Codigo As Integer) As String
Dim i As Integer
    For i = 1 To UBound(VecProvMinicenas)
        If Codigo = VecProvMinicenas(i).P_Codigo Then
            BuscarDescProvMinicena = Trim(VecProvMinicenas(i).P_Descripcion)
            Exit Function
        End If
    Next
End Function

Public Sub UbicarCmbProvMinicenas(Cmb As ComboEsp, Codigo As Integer)
    Dim i As Integer
    For i = 1 To UBound(VecProvMinicenas)
        If VecProvMinicenas(i).P_Codigo = Codigo Then
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next
        Cmb.ListIndex = 0
End Sub
