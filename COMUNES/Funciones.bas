Attribute VB_Name = "Funciones"

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, _
ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Enum Const_Set_Regional
       LOCALE_ILANGUAGE = &H1         '  Id. de idioma
       LOCALE_SLANGUAGE = &H2         '  nombre traducido del idioma
       LOCALE_SENGLANGUAGE = &H1001   '  nombre del idioma en inglés
       LOCALE_SABBREVLANGNAME = &H3   '  nombre del idioma abreviado
       LOCALE_SNATIVELANGNAME = &H4   '  nombre nativo del idioma
       LOCALE_ICOUNTRY = &H5          '  código del país
       LOCALE_SCOUNTRY = &H6          '  nombre traducido del país
       LOCALE_SENGCOUNTRY = &H1002    '  nombre del país en inglés
       LOCALE_SABBREVCTRYNAME = &H7   '  nombre abreviado del país
       LOCALE_SNATIVECTRYNAME = &H8   '  nombre nativo del país
       LOCALE_IDEFAULTLANGUAGE = &H9  '  Id. predeterminado del idioma
       LOCALE_IDEFAULTCOUNTRY = &HA   '  código predeterminado del país
       LOCALE_IDEFAULTCODEPAGE = &HB  '  página de códigos predeterminada
       LOCALE_SLIST = &HC             '  separador de elementos de lista
       LOCALE_IMEASURE = &HD          '  0 = métrico, 1 = EE.UU.
       LOCALE_SDECIMAL = &HE          '  separador de decimales
       LOCALE_STHOUSAND = &HF         '  separador de miles
       LOCALE_SGROUPING = &H10        '  agrupación de dígitos
       LOCALE_IDIGITS = &H11          '  número de dígitos fraccionarios
       LOCALE_ILZERO = &H12           '  número de ceros iniciales de decimales
       LOCALE_SNATIVEDIGITS = &H13    '  ASCII 0-9 nativo
End Enum

Public Function ConfiguracionRegional(ByVal Caracteristica As Const_Set_Regional) As String
   
   Dim buffer As String * 100, dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, Caracteristica, buffer, CLng(Len(buffer) - 1))
   ConfiguracionRegional = Left(buffer, InStr(1, buffer, Chr(0)) - 1)
 
End Function

' OTRAS FUNCIONES QUE NO TIENEN QUE VER CON EL CONTROL PERO QUE SE USAN EN PROCEDIMIENTOS
Public Function FechaSQL(Fecha As String, Servidor As String) As String
On Error GoTo Errores
    If Not VerificarFecha(Fecha) Then
        FechaSQL = ""
    Else
        If UCase(Servidor) = "SQL" Then
            FechaSQL = "'" & Mid$(Fecha, 4, 2) & "/" & Mid$(Fecha, 1, 2) & "/" & Mid$(Fecha, 7) & "'"
        Else
            FechaSQL = "#" & Mid$(Fecha, 4, 2) & "/" & Mid$(Fecha, 1, 2) & "/" & Mid$(Fecha, 7) & "#"
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function ValorLv(Lv, Indice As Long, Columna As String)
Dim NumCol As Integer
    NumCol = Val(Lv.ColumnHeaders(Columna).Index)
    If NumCol > 1 Then
        ValorLv = Lv.ListItems(Indice).SubItems(NumCol - 1)
    Else
        ValorLv = Lv.ListItems(Indice).Text
    End If
End Function

Public Sub AsignarValorLv(Lv, Indice As Long, Columna As String, Valor)
Dim NumCol As Integer
    NumCol = Val(Lv.ColumnHeaders(Columna).Index)
    If NumCol > 1 Then
        Lv.ListItems(Indice).SubItems(NumCol - 1) = Valor
    Else
        Lv.ListItems(Indice).Text = Valor
    End If
End Sub

Public Sub RemoverIndice(Lv, Vec(), Indice As Integer)
Dim Aux()
    Lv.ListItems.Remove Indice
    For i = 0 To Indice - 1
        Aux(i) = Vec(i)
    Next
    For i = Indice + 1 To UBound(Vec)
        Aux(i) = Vec(i)
    Next
    
    Vec = Aux
    
End Sub

Public Sub ContadordeLinea(Linea As Integer, cTexto As String, Optional MaxCantPorLinea As Integer = 0)
Dim pos As Integer
Dim Largo As Integer
Dim TextoActual As String
    TextoActual = cTexto
    Largo = Len(TextoActual)
    
    While Trim(TextoActual) <> ""
        If Asc(Mid$(TextoActual, 1, 1)) = 10 Then
            TextoActual = Mid$(TextoActual, 2)
        End If
        pos = InStr(1, TextoActual, Chr(13))
        
        If pos > 0 Then
            If pos < MaxCantPorLinea Then  'si el enter está antes que la maxima cantidad de letras por linea
                Linea = Linea + 1
                TextoActual = Mid$(TextoActual, pos + 1)
            Else
                Linea = Linea + 1
                TextoActual = Mid$(TextoActual, MaxCantPorLinea)
            End If
            
        Else
            If Largo > MaxCantPorLinea Then
                Linea = Linea + 1
                TextoActual = Mid$(TextoActual, MaxCantPorLinea)
            Else
                If TextoActual <> "" Then
                    Linea = Linea + 1
                    Exit Sub
                End If
            End If
       End If
    Wend
End Sub

Public Function ExisteValor(Combo As Control) As Boolean
On Error GoTo Errores
'es para indicar si el texto escrito en un combo existe o no
Dim i As Integer
Dim Encontro As Boolean
    i = 0
    Encontro = False
    While i < Combo.ListCount And Not Encontro
        If Combo.List(i) = Combo.Text Then
            Encontro = True
        End If
        i = i + 1
    Wend
    If Not Encontro Then
        Combo.SetFocus
    End If
    ExisteValor = Encontro
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function ValorCheck(Valor As String) As Integer
If Valor = "SI" Then
    ValorCheck = 1
Else
    ValorCheck = 0
End If
End Function

Public Function Rellenar(Texto As String, Optional Tamanio As Integer = 10, Optional Cortar As Boolean = False, Optional Izquierda As Boolean = True, Optional rellenado As String = " ") As String
On Error GoTo Errores
'rellena un texto con el tamanio indicado
Dim LargoTexto As Integer
Dim i As Integer
Dim TextoARellenar
    TextoARellenar = Trim(Texto)
    LargoTexto = Len(TextoARellenar)
    If LargoTexto < Tamanio Then
        For i = LargoTexto To Tamanio - 1
            If Izquierda Then
                TextoARellenar = rellenado & TextoARellenar
            Else
                TextoARellenar = TextoARellenar & rellenado
            End If
        Next
        Rellenar = TextoARellenar
    Else
        If Cortar Then
            Rellenar = Mid$(TextoARellenar, 1, Tamanio)
        Else
            Rellenar = TextoARellenar
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function CortarTexto(Texto As String, Maximo As Integer) As String
On Error GoTo Errores
'agarra un texto y lo corta si es superior al maximo
    If Len(Texto) > Maximo Then
        CortarTexto = Mid$(Texto, 1, Maximo)
    Else
        CortarTexto = Texto
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function Valor(Numero As String) As Double
'pasa un numero con . o con , a un valor adecuado
Dim Posicion As Integer
Dim signo As Integer
Dim PosComa As Integer
Dim PosPunto As Integer
Dim Formato As String
Dim TipoDecimal As String

    TipoDecimal = ConfiguracionRegional(LOCALE_DECIMAL)
    
    If Not HayNumero(Numero) Then
        Valor = 0
    Else
        Numero = Trim(Numero)
        Posicion = InStr(1, Numero, "-")
        'If Posicion > 0 Then
        '    signo = -1
        'Else
            signo = 1
        'End If
        Posicion = InStr(1, Numero, "$")
        If Posicion > 0 Then
            Numero = Mid$(Numero, Posicion + 1)
        End If
        
        While Posicion > 0
            If Posicion >= 1 Then
                Numero = Mid$(Numero, 1, Posicion - 1) + Mid$(Numero, Posicion + 1)
            End If
            Posicion = InStr(1, Numero, "-")
        Wend
            
        If IsNumeric("0" & Trim(Numero)) Or IsNumeric(Trim(Numero)) Then
            'If TipoDecimal = "." Then
            Posicion = InStr(1, Numero, ".")
            'Else
            '   posicion = InStr(1, Numero, ",")
            'End If
            If Posicion > 0 Then
               'If TipoDecimal = "," Then
                   Valor = CDbl(Val(Numero))
               'Else
               '    Valor = CDbl(Val(Numero)) / 100
               'End If
            Else
                Valor = CDbl(Numero)
            End If
            Valor = signo * Valor
        Else
            Valor = 0
        End If
    End If
End Function

'Public Function Valor(Numero As String)
'pasa un numero con . o con , a un valor adecuado
'Dim posicion As Integer
'Dim signo As Integer
'Dim PosComa As Integer
'Dim PosPunto As Integer
'Dim Formato As String
'Dim cuenta As Integer
'    Numero = Trim(Numero)
'
'    posicion = InStr(1, Numero, "$")
'    If posicion > 0 Then
'        Numero = Mid$(Numero, posicion + 1)
'    End If
'
'    posicion = InStr(1, Numero, "-")
'    If posicion > 0 Then
'        signo = -1
'    Else
'        signo = 1
'    End If
'
'    While posicion > 0
'        If posicion >= 1 Then
'            Numero = Mid$(Numero, 1, posicion - 1) + Mid$(Numero, posicion + 1)
'        End If
'        posicion = InStr(1, Numero, "-")
'    Wend
'
'    If IsNumeric("0" & Trim(Numero)) Or IsNumeric(Trim(Numero)) Then
'        PosComa = InStr(1, Numero, ",")
'        PosPunto = InStr(1, Numero, ".")
'        If PosPunto > PosComa Then
'            Formato = "I" 'formato inglés
'        Else
'            Formato = "C" 'formato castellano
'        End If
'        If Formato = "I" Then
'            'PASO LAS , A >
'            posicion = InStr(1, Numero, ",")
'            While posicion > 0
'                If posicion >= 1 Then
'                    Numero = Mid$(Numero, 1, posicion - 1) + ">" + Mid$(Numero, posicion + 1)
'                End If
'                posicion = InStr(1, Numero, ",")
'            Wend
'
'            'PASO LOS . A ,
'            posicion = InStr(1, Numero, ".")
'            While posicion > 0
'                If posicion >= 1 Then
'                    Numero = Mid$(Numero, 1, posicion - 1) + "," + Mid$(Numero, posicion + 1)
'                End If
'                posicion = InStr(1, Numero, ".")
'            Wend
'
'            'PASO LOS > A .
'            posicion = InStr(1, Numero, ">")
'            While posicion > 0
'                If posicion >= 1 Then
'                    Numero = Mid$(Numero, 1, posicion - 1) + "." + Mid$(Numero, posicion + 1)
'                End If
'                posicion = InStr(1, Numero, ">")
'            Wend
'        End If
'
'        'borro todos los puntos
'
'        posicion = InStr(1, Numero, ".")
'        While posicion > 0
'            If posicion >= 1 Then
'                Numero = Mid$(Numero, 1, posicion - 1) + Mid$(Numero, posicion + 1)
'            End If
'            posicion = InStr(1, Numero, ".")
'        Wend
'
'        cuenta = 0
'        posicion = InStr(1, Numero, ",")
'        While posicion > 0
'            cuenta = cuenta + 1
'            posicion = InStr(posicion + 1, Numero, ",")
'        Wend
'        If cuenta > 1 Then
'            Valor = 0
'        Else
'            If Not IsNull(Numero) And Trim(Numero) <> "" Then
'                posicion = InStr(1, Numero, ",")
'                If posicion > 0 Then
'                    If posicion > 1 Then
'                        'Valor = CDbl(Numero)
'                        Valor = Numero
'                    Else
'                        Valor = CDbl("0," + Mid$(Numero, 2))
'                    End If
'                Else
'                    Valor = CDbl(Numero)
'                End If
'            Else
'                Valor = 0
'            End If
'            Valor = signo * Valor
'        End If
'    Else
'        Valor = 0
'    End If
'End Function

Public Function ValorSINO(Valor As Boolean) As String
    If Valor = True Or Valor = 1 Then
        ValorSINO = "SI"
    Else
        ValorSINO = "NO"
    End If
End Function

Public Function valorIngles(Numero As String) As String
On Error GoTo Errores
'Es para poner un valor en las consultas SQL, porque no se pueden usar las ,
Dim Posicion As Integer
    If Not IsNull(Numero) And Trim(Numero) <> "" Then
        Posicion = InStr(1, Numero, ",")
        If Posicion > 0 Then
            If Posicion > 1 Then
                valorIngles = Str(Val(Mid$(Numero, 1, Posicion - 1) + "." + Mid$(Numero, Posicion + 1)))
            Else
                valorIngles = Str("0." + Mid$(Numero, 2))
            End If
        Else
            valorIngles = Numero
        End If
    Else
        valorIngles = 0
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Sub BuscarTexto(Combo As ComboBox)
On Error GoTo Errores
'busca el texto escrito en el combo, si lo encuentra,
'lo asigna al valor de list index
Dim Encuentra As Boolean
Dim i As Integer
Dim Texto As String
    Encuentra = False
    i = 0
    Texto = Combo.Text
    While Not Encuentra And i < Combo.ListCount
        If UCase(Texto) = UCase(Mid$(Combo.List(i), 1, Len(Texto))) Then
            Combo.ListIndex = i
            Encuentra = True
        End If
        i = i + 1
    Wend
    If Not Encuentra Then Combo.SetFocus
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Function BuscarString(con As String, stringbusqueda As String, Optional ValorDefault) As String
On Error GoTo Errores
'es para recuperar parametros
Dim inicio As Integer
Dim Fin As Integer
    BuscarString = " "
    inicio = InStr(1, con, stringbusqueda, vbTextCompare)
    If inicio <> 0 Then
        inicio = inicio + Len(stringbusqueda)
        Fin = InStr(inicio, con, ";")
        If Fin <> 0 Then
            BuscarString = Mid$(con, inicio, Fin - inicio)
        Else
            BuscarString = Mid$(con, inicio)
        End If
    Else
        If Not IsMissing(ValorDefault) Then
            BuscarString = ValorDefault
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Sub RecuperarComprobante(ByVal Comprobante As String, _
                                TipoComprobante As String, _
                                Sucursal As String, _
                                Letra As String, _
                                Optional Numero As String)
Dim CompNuevo
    DevolverTexto Comprobante, TipoComprobante
    DevolverTexto Comprobante, Sucursal
    DevolverTexto Comprobante, Letra
    If Not IsMissing(Numero) Then
        DevolverTexto Comprobante, Numero
    End If
End Sub
                                
Public Sub DevolverTexto(ByRef TextoAnterior As String, ByRef TextoNuevo As String)
'esta funcion se usa para recuperar un comprobante (las distintas secciones)
On Error GoTo Errores
Dim Posicion As Integer
    Posicion = InStr(1, TextoAnterior, "-")
    If Posicion > 0 Then
        TextoNuevo = Mid$(TextoAnterior, 1, Posicion - 1)
        TextoAnterior = Mid$(TextoAnterior, Posicion + 1)
    Else
        TextoNuevo = TextoAnterior
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Sub ValidarFecha(Texto As TextBox)
On Error GoTo Errores
    If Len(Fecha) <> 10 Or Not IsDate(Fecha) Or _
        Val(Mid$(Fecha, 4, 2)) > 12 Then
         MsgBox "La fecha indicada no es correcta"
        Texto.SetFocus
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Function VerificarFecha(Texto As String) As Boolean
On Error GoTo Errores
    If Len(Texto) <> 10 Or Not IsDate(Texto) Or _
        Val(Mid$(Texto, 4, 2)) > 12 Or Val(Mid$(Texto, 7)) < 1901 Then
            VerificarFecha = False
        Else
            VerificarFecha = True
    End If
Errores:
    TratarError Err.Number, Err.Description

End Function

Public Function VerificarPeriodo(Texto As String) As Boolean
On Error GoTo Errores
    If Len(Texto) <> 7 Or _
        Val(Mid$(Texto, 1, 2)) > 12 Or Val(Mid$(Texto, 4)) < 1901 Then
            VerificarPeriodo = False
        Else
            VerificarPeriodo = True
    End If
Errores:
    TratarError Err.Number, Err.Description

End Function

Public Function VerificarNumero(Texto As String) As Boolean
On Error GoTo Errores
Dim i As Integer

    For i = 1 To Len(Trim(Texto))
        If Mid$(Texto, i, 1) <> "0" And _
           Mid$(Texto, i, 1) <> "1" And _
           Mid$(Texto, i, 1) <> "2" And _
           Mid$(Texto, i, 1) <> "3" And _
           Mid$(Texto, i, 1) <> "4" And _
           Mid$(Texto, i, 1) <> "5" And _
           Mid$(Texto, i, 1) <> "6" And _
           Mid$(Texto, i, 1) <> "7" And _
           Mid$(Texto, i, 1) <> "8" And _
           Mid$(Texto, i, 1) <> "9" Then
           VerificarNumero = False
           i = Len(Trim(Texto))
        Else
            VerificarNumero = True
        End If
    Next
Errores:
    TratarError Err.Number, Err.Description

End Function

Public Function LenMes(mes As Integer, Ano As Integer) As Integer
Select Case mes
    Case 1
        LenMes = 31
    Case 2
        LenMes = 28
        
        If Int(Ano / 4) - (Ano / 4) = 0 Then
            LenMes = 29
        End If
    Case 3
        LenMes = 31
    Case 4
        LenMes = 30
    Case 5
        LenMes = 31
    Case 6
        LenMes = 30
    Case 7
        LenMes = 31
    Case 8
        LenMes = 31
    Case 9
        LenMes = 30
    Case 10
        LenMes = 31
    Case 11
        LenMes = 30
    Case 12
        LenMes = 31
End Select

End Function

Public Sub TeclaPresionada(cnt As Control, ByRef KeyAscii As Integer)
On Error GoTo Errores
    
    If KeyAscii = 27 Then
        ' si apreta esc se asume que quiere cancelar
        KeyAscii = 0
'        Unload ActiveForm
        Exit Sub
    End If
    If Not (cnt Is Nothing) Then
        If TypeOf cnt Is CommandButton Then
            
            'si es un botón de comando, se lo considera "click"
            'en el botón
            Exit Sub
        Else
            If KeyAscii = 13 And Not (TypeOf cnt Is ListView) Then
                'al apretar Enter, se cambia a este por un
                'TAB para pasar al control siguiente
                KeyAscii = 0
                SendKeys "{tab}"
            Else
         '       If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is ComboBox Then
                If Not (TypeOf cnt Is ListView) And (Not TypeOf cnt Is CheckBox) And (Not TypeOf cnt Is TreeView) Then
                    Mascara KeyAscii, cnt
                End If
            End If
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Sub Mascara(KeyAscii As Integer, cnt As Control)
On Error GoTo Errores
Dim Posicion As Integer
Const digitos = "0123456789,.-"
Const Numeros = "0123456789"
'Const letras = "abcdefghijklmnopqrstuvwxyzñ"
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If Len(cnt.Text) = 0 Then
        Posicion = 1
    Else
        Posicion = Len(cnt.Text) + 1
    End If
    Select Case Mid$(cnt.Tag, Posicion, 1)
    Case "0"
        CargarMascara digitos, cnt, KeyAscii
    Case "9"
        CargarMascara Numeros, cnt, KeyAscii
    Case "X"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case "x"
        KeyAscii = Asc(Chr(KeyAscii))
    Case Else
        CargarMascara Mid$(cnt.Tag, Posicion, 1), cnt, KeyAscii
    End Select
Errores:
    TratarError Err.Number, Err.Description
End Sub

Private Sub CargarMascara(TextoMascara As String, cnt As Control, KeyAscii As Integer)
Dim Posicion As Integer
    If Len(cnt.Text) = Len(cnt.Tag) Then
        KeyAscii = 0
    Else
        If Len(cnt.Text) = 0 Then
            Posicion = 1
        Else
            Posicion = Len(cnt.Text) + 1
        End If

        If InStr(1, TextoMascara, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        Else
            If ConfiguracionRegional(LOCALE_SDECIMAL) = "." Then
                If Chr(KeyAscii) = "," And Mid$(cnt.Tag, Posicion, 1) = "0" Then
                   KeyAscii = Asc(".")
                Else
                   If Len(cnt.Text) < Len(cnt.Tag) Then
                      If Mid$(cnt.Tag, Posicion + 1, 1) <> "0" And Mid$(cnt.Tag, Posicion + 1, 1) <> "9" Then
                         cnt.Text = cnt.Text + Chr(KeyAscii) + Mid$(cnt.Tag, Posicion + 1, 1)
                         KeyAscii = 0
                         cnt.SelStart = Len(cnt.Text)
                      End If
                   End If
                End If
            Else
                If ConfiguracionRegional(LOCALE_SDECIMAL) = "," Then
                    If Chr(KeyAscii) = "." And Mid$(cnt.Tag, Posicion, 1) = "0" Then
                        KeyAscii = Asc(",")
                    End If
               Else
                   If Len(cnt.Text) < Len(cnt.Tag) Then
                       If Mid$(cnt.Tag, Posicion + 1, 1) <> "0" And Mid$(cnt.Tag, Posicion + 1, 1) <> "9" Then
                           cnt.Text = cnt.Text + Chr(KeyAscii) + Mid$(cnt.Tag, Posicion + 1, 1)
                           KeyAscii = 0
                          cnt.SelStart = Len(cnt.Text)
                       End If
                   End If
              End If
            End If
        End If
    End If
End Sub

Public Sub SoloDigitos(ByRef KeyAscii As Integer)
On Error GoTo Errores
    Const digitos = "0123456789,."
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If InStr(digitos, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Sub Numerodecimal(ByRef KeyAscii As Integer, AControl As Control)
On Error GoTo Errores
    Const digitos = "0123456789,"
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If InStr(digitos, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        If InStr(AControl.Text, Chr(44)) <> 0 And KeyAscii = 44 Then 'si puso una coma
            KeyAscii = 0
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Sub DigitosySignos(ByRef KeyAscii As Integer)
    Const digitos = "0123456789-()"
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If InStr(digitos, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Public Sub DigitosHora(ByRef KeyAscii As Integer)
    Const digitos = "0123456789:"
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If InStr(digitos, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Public Sub SelText(C As Control)
On Error GoTo Errores
    If TypeOf C Is TextBox Or TypeOf C Is MaskEdBox Then
        C.SelStart = 0
        C.SelLength = Len(C.Text)
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

'Public Sub CamposLista(ListView1 As ListView, Indice As Integer, Texto As String, Largo As Integer)
'Dim btnx As ColumnHeader
'    Set btnx = ListView1.ColumnHeaders.Add(Indice, Texto, Texto, Largo)
'End Sub

'Public Sub cargarItem(ListView1 As ListView, Primero As Variant, Buscarpor As Variant, ParamArray items())
'Dim btnx As ListItem
'Dim Elemento As Variant
'Dim i As Integer
'    i = 1
'    Set btnx = ListView1.ListItems.Add(, , Primero)
'    For Each Elemento In items
'        btnx.SubItems(i) = Elemento
'        i = i + 1
'    Next
'    btnx.Tag = Buscarpor
'End Sub


Public Function Digicuit(sCUIT As String) As Boolean
On Error GoTo Errores
    Dim coef(11) As Integer
    Dim Sum As Integer
    Dim Lon, Resultado
    Dim i, Resto
    Dim CUIT As String
    Dim Dverificador As Integer
    If Trim(sCUIT) <> "" Then
            If Len(sCUIT) <> 13 Then
                Digicuit = False
            Else
                CUIT = Mid$(sCUIT, 1, 2) + Mid$(sCUIT, 4, 8)
                Dverificador = Mid$(sCUIT, 13, 1)
                coef(0) = 0
                coef(1) = 5
                coef(2) = 4
                coef(3) = 3
                coef(4) = 2
                coef(5) = 7
                coef(6) = 6
                coef(7) = 5
                coef(8) = 4
                coef(9) = 3
                coef(10) = 2
                Sum = 0
                For i = 1 To 10
                    Sum = Sum + Val(Mid$(CUIT, i, 1)) * coef(i)
                Next
                Resto = Sum Mod 11
                If 11 - Resto <> Dverificador Then
                    Digicuit = False
                Else
                    Digicuit = True
                End If
            End If
      Else
          Digicuit = True
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Function BooleanoSQL(Valor As String, Servidor As String) As Boolean
'esta funcion es para recuperar el valor desde una consulta sql
    If UCase(Servidor) = "SQL" Then
        If Valor = "1" Or Valor = "Verdadero" Or Valor = "True" Then
            BooleanoSQL = True
        Else
            BooleanoSQL = False
        End If
    Else
        If Valor = "True" Or Valor = "Verdadero" Or Valor = "-1" Then
            BooleanoSQL = True
        Else
            BooleanoSQL = False
        End If
    End If
End Function

Function BooleanoSQL2(Valor As String, Servidor As String) As Integer
'esta funcion es para meterla en una consulta sql
    Valor = Trim(UCase(Valor))
    If UCase(Servidor) = "SQL" Then
        If Valor = "1" Or Valor = "VERDADERO" Or Valor = "TRUE" Or Valor = "SI" Then
            BooleanoSQL2 = 1
        Else
            BooleanoSQL2 = 0
        End If
    Else
        If Valor = "TRUE" Or Valor = "VERDADERO" Or Valor = "-1" Or Valor = "SI" Then
            BooleanoSQL2 = 0
        Else
            BooleanoSQL2 = -1
        End If
    End If
End Function

Function SinEspaciosSQL(Valor As String, Servidor As String) As String
    If UCase(Servidor) = "SQL" Then
        SinEspaciosSQL = "RTRIM(LTRIM(" & Valor & "))"
    Else
        SinEspaciosSQL = "TRIM(" & Valor & ")"
    End If
End Function

Function ValueDateSQL(Valor As String, Servidor As String) As String
    If UCase(Servidor) = "SQL" Then
        ValueDateSQL = "'" & Mid$(Valor, 4, 2) & "/" & Mid$(Valor, 1, 2) & "/" & Mid$(Valor, 7) & "'"
    Else
        ValueDateSQL = "#" & Mid$(Valor, 4, 2) & "/" & Mid$(Valor, 1, 2) & "/" & Mid$(Valor, 7) & "#"
    End If
End Function

Public Function Repetir(Valor As String, Cantidad As Integer) As String
On Error GoTo Errores
Dim i As Integer
Dim sValor As String
    sValor = ""
    For i = 1 To Cantidad
        sValor = sValor + Valor
    Next
    Repetir = sValor
Errores:
    TratarError Err.Number, Err.Description
End Function

Public Function VerificarNulo(ByVal Valor, Optional Tipo As String = "S", Optional Predeterminado) As String
On Error GoTo Errores
    If IsNull(Valor) Or Valor = "" Or (Tipo = "F" And Valor = "01/01/1900") Then
        If Not IsMissing(Predeterminado) Then
            VerificarNulo = Predeterminado
        Else
            Select Case UCase(Tipo)
                Case "S"
                   VerificarNulo = ""
                Case "N"
                    VerificarNulo = "0"
                Case "F"
                    VerificarNulo = ""
                Case "B"
                    VerificarNulo = "False"
            End Select
        End If
    Else
        VerificarNulo = Valor
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function
Public Function Documento(ByVal Valor As String) As String
Dim X As Integer
            
           If Len(Str(Valor)) = 7 Then
              For X = 1 To Len(Str(Valor))
                If Mid$(Valor, X, 1) = "," Then
                   Documento = Documento + "."
                Else
                   Documento = Documento + Mid$(Valor, X, 1)
                End If
              Next
              Documento = "0." + Documento
              
           Else
              For X = 1 To Len(Str(Valor)) + 2
                If Mid$(Valor, X, 1) = "," Then
                   Documento = Documento + "."
                Else
                   Documento = Documento + Mid$(Valor, X, 1)
                End If
              Next
           End If
End Function

Public Sub GuardarEncabezado(Lv, Indice As Integer, Clave As String, Texto As String, ancho As Integer, Optional Otro As String, Optional Alineacion As Integer)
On Error GoTo Errores
    If Not IsMissing(Alineacion) Then
        Lv.ColumnHeaders.Add Indice, Clave, Texto, ancho, Alineacion
    Else
        Lv.ColumnHeaders.Add Indice, Clave, Texto, ancho
    End If
    If Not IsMissing(Otro) Then
        Lv.ColumnHeaders(Indice).Tag = Otro
    End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Function VerificarHora(Hora As String) As Boolean
On Error GoTo Errores
    VerificarHora = True
    If Len(Hora) <> 5 Then
        VerificarHora = False
    Else
        If Val(Mid$(Hora, 1, 2)) > 23 Then
            VerificarHora = False
        Else
            If Val(Mid$(Hora, 4, 2)) > 59 Then
                VerificarHora = False
            End If
        End If
    End If
Errores:
    TratarError Err.Number, Err.Description
End Function

Function StringSQL(Valor As String, Servidor As String) As String
    If UCase(Servidor) = "SQL" Then
        StringSQL = "CONVERT(char(50)," & Valor & ")"
    Else
        StringSQL = "STR(" & Valor & ")"
    End If
End Function

Public Sub SetearImpresora(PuntodeVenta As String, puerto As String, db As rdoConnection)
On Error GoTo Errores
Dim p As Object
Dim Nombre As rdoResultset
Dim TraerNombreImpresora As String

    Set Nombre = db.OpenResultset("SELECT P_ImpresoraA, P_ImpresoraB, P_ImpresoraR, P_impresoraOtros FROM EN_Puntosdeventa WHERE P_codigo='" & PuntodeVenta & "'")
    If Not Nombre.EOF Then
        Select Case puerto
            Case "1"
                TraerNombreImpresora = Trim(VerificarNulo(Nombre!P_ImpresoraA, "S", "Impresora 1"))
            Case "2"
                TraerNombreImpresora = Trim(VerificarNulo(Nombre!P_ImpresoraB, "S", "Impresora 1"))
            Case "3"
                TraerNombreImpresora = Trim(VerificarNulo(Nombre!P_ImpresoraR, "S", "Impresora 1"))
            Case "4"
                TraerNombreImpresora = Trim(VerificarNulo(Nombre!P_ImpresoraC, "S", "Impresora 1"))
            Case "5"
                TraerNombreImpresora = Trim(VerificarNulo(Nombre!P_ImpresoraOtros, "S", "Impresora 1"))
                
        End Select
        For Each p In Printers
             If InStr(1, p.DeviceName, TraerNombreImpresora, vbTextCompare) > 0 Then
                Set Printer = p
                Exit For
            End If
         Next p
    Else
        MsgBox "No existe una impresora para el punto de venta " & Lbpuntodeventasesion.Caption & "", 16, "El Pulqui"
   End If
Errores:
    TratarError Err.Number, Err.Description
End Sub

Public Function HayNumero(Numero As String)
    HayNumero = InStr(1, Numero, "0") <> 0 Or _
        InStr(1, Numero, "1") <> 0 Or _
        InStr(1, Numero, "2") <> 0 Or _
        InStr(1, Numero, "3") <> 0 Or _
        InStr(1, Numero, "4") <> 0 Or _
        InStr(1, Numero, "5") <> 0 Or _
        InStr(1, Numero, "6") <> 0 Or _
        InStr(1, Numero, "7") <> 0 Or _
        InStr(1, Numero, "8") <> 0 Or _
        InStr(1, Numero, "9") <> 0
End Function

Public Sub TratarError(Numero As Long, Descripcion As String, Optional HuboError As Boolean)
    If Numero <> 0 Then
        NumErrores Numero, Descripcion
        If Not IsMissing(HuboError) Then
            HuboError = True
        End If
    End If
End Sub

Public Sub ManipularError(Numero As Long, Descripcion As String, ParamArray Controles())
Dim C
'esta funcion tiene que hacer lo mismo que tratar error
'pero no puedo borrar la anterior porque ya lo usé en algunos controles
'y a los controles que recibo como parámetros los desactivo
'es en especial para los timers.
    If Numero <> 0 Then
        NumErrores Numero, Descripcion
        For Each C In Controles
            C.Enabled = False
        Next
        MousePointer = vbNormal
    End If
End Sub

Private Sub NumErrores(Numero As Long, Descripcion As String)
    Select Case Numero
        Case 40041
            MsgBox "El campo " & Descripcion & " es inexistente", vbCritical
        Case Else
            MsgBox "Error " & Numero & Chr(13) & Descripcion, vbCritical
    End Select
End Sub

Public Sub ComenzarTransaccion(db As rdoConnection, Servidor As String)
    If UCase(Servidor) = "SQL" Then
        db.BeginTrans
    End If
End Sub

Public Sub TerminarTransaccion(db As rdoConnection, Servidor As String)
    If UCase(Servidor) = "SQL" Then
        db.CommitTrans
    End If
End Sub

Public Sub CancelarTransaccion(db As rdoConnection, Servidor As String)
    If UCase(Servidor) = "SQL" Then
        db.RollbackTrans
    End If
End Sub

Public Sub Inc(ByRef Numero, Optional Cantidad = 1)
    Numero = Numero + Cantidad
End Sub

Public Function Espacios(Cantidad As Integer) As String
Dim i As Integer
    Espacios = ""
    For i = 1 To Cantidad
        Espacios = Espacios + " "
    Next
End Function

Public Function Redondear(Numero As Double, Optional Decimales As Integer, Optional DigitosAVerificar As Integer = 1) As Double
Dim pos As Integer
Dim Redondeo As Double
Dim DigitosVerificados As Integer
Dim Encontro As Boolean
    pos = InStr(1, Numero, ",")
    If pos = 0 Then
        Redondeo = Numero
    Else
        Redondeo = 0
    End If
    DigitosVerificados = 1
    Encontro = False
    While Redondeo = 0 And DigitosVerificados <= DigitosAVerificar
        If Val(Mid$(Numero, pos + Decimales + DigitosVerificados, 1)) > 5 Then
            Redondeo = Mid$(Numero, 1, pos + Decimales) + 10 ^ (-1 * Decimales)
            Encontro = True
        End If
        DigitosVerificados = DigitosVerificados + 1
    Wend
    If Not Encontro Then
        Redondeo = Mid$(Numero, 1, pos + Decimales)
    End If
    Redondear = Mid$(Redondeo, 1, pos + Decimales)
End Function

Public Function Valcuit(sCUIT As String) As Boolean
Dim nCantdigitos As Integer
Dim nVerificador As Integer
Dim nFactor As Integer
Dim nPosicion As Integer
Dim nDigito As Integer
Dim nDigitoVerificador As Integer
Dim nPosVerif As Integer
Dim cC As String

    nCantdigitos = 0
    nVerificador = 0
    nFactor = 2
    cC = ""
    
    For nPosicion = Len(sCUIT) To 1 Step -1
        cC = Mid$(sCUIT, nPosicion, 1)
        If cC <> "-" Then
            nDigito = Val(cC)
            nCantdigitos = nCantdigitos + 1
            If nCantdigitos = 1 Then
                nDigitoVerificador = nDigito
                nPosVerif = nPosicion
            Else
                nVerificador = nVerificador + nDigito * nFactor
                nFactor = nFactor + 1
                If nFactor > 7 Then
                    nFactor = 2
                End If
            End If
        End If
     Next
     nVerificador = nVerificador Mod 11
     If nVerificador <> 0 Then
        nVerificador = 11 - nVerificador
     End If
     If nVerificador = nDigitoVerificador Then
        Valcuit = True
     Else
        Valcuit = False
     End If
End Function

Public Function ValorSQL(Valor As String, Servidor As String) As String
    If UCase(Servidor) = "SQL" Then
        ValorSQL = "CONVERT(int, " & Valor & ") "
    Else
        ValorSQL = "VAL(" & Valor & ")"
    End If
End Function

Public Sub ArmarExcel(Dialogo As CommonDialog)
    ' Establecer CancelError a True
    Dialogo.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    Dialogo.Flags = cdlOFNHideReadOnly
    ' Establecer los filtros
    Dialogo.Filter = "Excel 2000|*.xls" & _
    "|Excel 97|*.xls" & _
    "|Excel 2007|*.xlsx"
    ' Especificar el filtro predeterminado
    Dialogo.FilterIndex = 1
    ' Presentar el cuadro de diálogo Abrir
    Dialogo.ShowSave
    ' Presentar el nombre del archivo seleccionado
    MousePointer = vbHourglass
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
'    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub DatosExcel(ex As Excel.Application, listado As ListView, ByVal FilaInicial As Integer)
Dim Fila As Long
Dim col As Integer
Dim i As Integer
    Fila = FilaInicial + 1
    col = 1
    With ex
        For i = 1 To listado.ListItems.Count
            If VerificarFecha(listado.ListItems(i).Text) Then
                .Range(LetraColumna(1) & Trim(Fila)).Value = Month(listado.ListItems(i).Text) & "/" & Day(listado.ListItems(i).Text) & "/" & Year(listado.ListItems(i).Text)  'Format(listado.ListItems(i).Text, "dd/MM/yyyy")
                '.Selection.NumberFormat = "dd/MM/yyyy"
            Else
                .Range(LetraColumna(1) & Trim(Fila)).Value = listado.ListItems(i).Text
            End If
            If Trim(listado.ListItems(i).SubItems(1)) <> "" Then
               If VerificarFecha(listado.ListItems(i).SubItems(1)) Then
                  .Range(LetraColumna(2) & Trim(Fila)).Value = Format(listado.ListItems(i).SubItems(col), "dd/MM/yyyy") 'Month(listado.ListItems(i).SubItems(1)) & "/" & Day(listado.ListItems(i).SubItems(1)) & "/" & Year(listado.ListItems(i).SubItems(1))
               Else
                  .Range(LetraColumna(2) & Trim(Fila)).Value = listado.ListItems(i).SubItems(1)
               End If
            End If
            For col = 1 To listado.ColumnHeaders.Count - 2
               If Trim(listado.ListItems(i).SubItems(col)) <> "" Then
                  If VerificarFecha(listado.ListItems(i).SubItems(col)) Then
                     .Range(LetraColumna(col + 1) & Trim(Fila)).FormulaR1C1 = Format(listado.ListItems(i).SubItems(col), "dd/MM/yyyy") 'Month(listado.ListItems(i).SubItems(col)) & "/" & Day(listado.ListItems(i).SubItems(col)) & "/" & Year(listado.ListItems(i).SubItems(col))
                  Else
                    If IsNumeric(listado.ListItems(i).SubItems(col)) Then
                       .Range(LetraColumna(col + 1) & Trim(Fila)).FormulaR1C1 = Replace(FormatNumber(listado.ListItems(i).SubItems(col), 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
                       .Range(LetraColumna(col + 1) & Trim(Fila)).Select
                       .Selection.NumberFormat = "0.00"
                   Else
                       .Range(LetraColumna(col + 1) & Trim(Fila)).FormulaR1C1 = listado.ListItems(i).SubItems(col)
                    End If
                  End If
               End If
            Next
            Fila = Fila + 1
        Next
    End With
End Sub

Public Function LetraColumna(col As Integer) As String
Dim Columnas As String
    
    Columnas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If col <= 26 Then
        LetraColumna = Mid$(Columnas, col, 1)
    Else
        LetraColumna = LetraColumna((col - 1) \ 26) & Mid$(Columnas, IIf(col Mod 26 = 0, 26, col Mod 26), 1)
    End If
End Function

Public Sub GuardarPlanilla(ex As Excel.Application, NombreArchivo As String, Filtro As Integer)
    Select Case Filtro
        Case 1
            ex.ActiveWorkbook.SaveAs Filename:=NombreArchivo, FileFormat:=xlNormal, _
                Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
                CreateBackup:=False
        Case 2
            ex.ActiveWorkbook.SaveAs Filename:=NombreArchivo, FileFormat:=xlExcel9795, _
                Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
                CreateBackup:=False
        Case 3
            ex.ActiveWorkbook.SaveAs Filename:=NombreArchivo, FileFormat:=51, _
                Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
                CreateBackup:=False
    End Select
End Sub

Public Sub EncabezadoExcel(ex As Excel.Application, listado As ListView, ByVal Titulo As String, ByVal FilaInicial As Integer)
Dim col As Integer
Dim i As Long
    With ex
    
        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = Titulo
        .Range("A1:G1").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
        
        .Selection.Merge  'COMBINAR CELDAS
        
        With .Selection.Font
            .Name = "Arial"
            .Size = 20
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
        .Selection.Font.Bold = True
        For col = 1 To listado.ColumnHeaders.Count - 1
            .Range(LetraColumna(col) & Trim(FilaInicial)).Select
            With .ActiveCell
                .FormulaR1C1 = listado.ColumnHeaders(col).Text
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
            End With
        Next
        
        .Rows(Trim(Str(FilaInicial)) & ":" & Trim(Str(FilaInicial))).Select
        
        With .Selection
         
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
            .Font.Bold = True
        End With
        .Rows(Trim(Str(FilaInicial)) & ":" & Trim(Str(FilaInicial))).EntireRow.AutoFit
        
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & LetraColumna(listado.ColumnHeaders.Count - 1) & Trim(Str(FilaInicial))).Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
    End With
End Sub
