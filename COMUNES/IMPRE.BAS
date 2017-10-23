Attribute VB_Name = "impre"
Public Type impre
       nAnchoHoja      As Single
       nAnchoLetra     As Single
       nAltoHoja       As Single
       nAltoLetra      As Single
       nCharMax        As Single
       nClMax          As Single
       nEscala         As Byte
End Type
Public Impresora As impre
Const cRazonSocial = "El Pulqui"

Public Sub Setear_lis(cLetra As String, nDimension As Single, Optional cNegra As String = "N", Optional ByVal nEscala As Byte = 100)
    'Dim X As New StdFont
    'X.Name = cLetra
    'X.Size = nDimension
    'Set Printer.Font = X

    If cNegra = "" Then
       cNegra = "N"
    End If
    'Printer.Print ""

    Printer.FontName = cLetra
    Printer.FontSize = nDimension
    If cNegra = "S" Then
       Printer.FontBold = True
     Else
       Printer.FontBold = False
    End If
    '
    Impresora.nAltoHoja = Printer.Height
    Impresora.nAnchoHoja = Printer.Width - Int(Printer.Width * 0.1)
'    Impresora.nAltoLetra = Printer.TextHeight("Q") + Int(Printer.TextHeight("Q") * 0.2)
    Impresora.nAltoLetra = Printer.TextHeight("Q")
    Impresora.nAnchoLetra = Printer.TextWidth("Z")
    Impresora.nCharMax = Int(Impresora.nAnchoHoja / Impresora.nAnchoLetra)
    Impresora.nClMax = Int(Impresora.nAltoHoja / Impresora.nAltoLetra) - 5
    Impresora.nEscala = nEscala
End Sub


Public Sub Say_Lis(nfila As Single, nColumna As Single, cTexto As String, Optional nalto, Optional Centrar As Boolean = False)

 '   Printer.CurrentX = Int(Impresora.nAnchoHoja * nColumna / Impresora.nEscala)
    If Centrar Then
        Printer.CurrentX = Int(Impresora.nAnchoLetra * (nColumna)) - Int(Impresora.nAnchoLetra / 2)
    Else
        Printer.CurrentX = Int(Impresora.nAnchoLetra * (nColumna)) - Impresora.nAnchoLetra
    End If
    If IsMissing(nalto) Then
       Printer.CurrentY = nfila * Impresora.nAltoLetra ' De acuerdo a letra actual
     Else
       Printer.CurrentY = nfila * nalto ' De acuerdo al parámetro
    End If
    
    Printer.Print cTexto
    If Centrar Then
        Printer.CurrentX = Printer.CurrentX - Int(Impresora.nAnchoLetra / 2)
    End If
End Sub

Public Sub Say_MultipleLinea(nfila As Single, nColumna As Single, cTexto As String, Optional MaxCantPorLinea As Integer = 0, Optional MaxLinea As Integer = 0, Optional nalto)
Dim pos As Integer
Dim TextoActual As String
Dim Linea As Integer
    Linea = 1
    TextoActual = cTexto
    While Trim(TextoActual) <> "" And (Linea <= MaxLinea Or MaxLinea = 0)
        If Asc(Mid$(TextoActual, 1, 1)) = 10 Then
            TextoActual = Mid$(TextoActual, 2)
        End If
        pos = InStr(1, TextoActual, Chr(13))
        If pos > 0 Then
            If pos < MaxCantPorLinea Or MaxCantPorLinea = 0 Then 'si el enter está antes que la maxima cantidad de letras por linea
                Say_Lis nfila, nColumna, Mid$(TextoActual, 1, pos - 1), nalto
                nfila = nfila + 1
                TextoActual = Mid$(TextoActual, pos + 1)
            Else
                Say_Lis nfila, nColumna, Mid$(TextoActual, 1, MaxCantPorLinea - 1), nalto
                nfila = nfila + 1
                TextoActual = Mid$(TextoActual, MaxCantPorLinea)
            End If
         Else
                If TextoActual <> "" Then
                    Say_Lis nfila, nColumna, Mid$(TextoActual, 1, MaxCantPorLinea - 1), nalto
                    nfila = nfila + 1
                    TextoActual = Mid$(TextoActual, MaxCantPorLinea)
                End If
                
        End If
        Linea = Linea + 1
    Wend
End Sub

Public Sub Say_Lis2(nfila As Single, nColumna As Single, cTexto As String, Optional nalto)
Dim i As Integer
Dim Y As Integer
    Y = nfila
    For i = 1 To Len(cTexto)
        Say_Lis nfila, nColumna + i - 1, Mid$(cTexto, i, 1), nalto, True
        nfila = Y
    Next
End Sub



Public Sub Centrar_lis(cTexto As String, ByVal nfila As Single, Optional nalto)
    Dim nSay As Single

'    nAncho = Impresora.nAnchoHoja / 2
'    nocupa = Printer.TextWidth(cTexto) / 2
'    nsay = nAncho - nocupa
    nSay = (Impresora.nAnchoHoja - Printer.TextWidth(Trim(cTexto))) / 2

    Printer.CurrentX = nSay
    If IsMissing(nalto) Then
       Printer.CurrentY = nfila * Impresora.nAltoLetra
    Else
       Printer.CurrentY = nfila * nalto
    End If
    Printer.Print cTexto


End Sub

Public Sub Enca_lis(ByVal nPagina As Integer)
Dim cImprimir As String
Dim nSay As Single
'    Printer.CurrentX = 0
'    Printer.CurrentY = 0
    Printer.Print ""
'    Setear_lis "Arial", 10  'Defino cualquier letra para que abajo "registre el cambio ??!!!"
'    Printer.CurrentX = 0
'    Printer.CurrentY = 0
'    Printer.Print ""

    Setear_lis "Courier New", 10  ' Define nueva letra

    Say_Lis 0, 2, "Fecha: " & Date
    Centrar_lis cRazonSocial, 0
    'If Printer.Orientation = 1 Then
    cImprimir = "Hoja: " & Trim(Str(nPagina))
    nSay = (Impresora.nAnchoHoja - Printer.TextWidth(Trim(cImprimir)))

    Printer.CurrentX = nSay
    Printer.CurrentY = 0
    Printer.Print cImprimir

    'Say_Lis 0, Impresora.nAnchoHoja - 20, "Hoja: " & Trim(Str(nPagina))
    'Else
   '     Say_Lis 0, 105, "Hoja: " & Trim(Str(nPagina))
   ' End If
    Say_Lis 1, 2, "Hora : " & Time

    nYhasta = 2 * Impresora.nAltoLetra
    Printer.Line (8, nYhasta)-(Impresora.nAnchoHoja, nYhasta)

End Sub
Public Sub Recuadro(ByVal nFila1 As Single, ByVal nColu1 As Single, ByVal nFila2 As Single, ByVal nColu2 As Single, Optional nalto)

    '*********************'
    ' Dibuja un recuadro
    '*********************'

    If IsMissing(nalto) Then
       nalto = Impresora.nAltoLetra  ' De acuerdo a letra actual
    End If

    nColu1 = Int(Impresora.nAnchoHoja * nColu1 / Impresora.nEscala) ' Redefino columna desde
    nColu2 = Int(Impresora.nAnchoHoja * nColu2 / Impresora.nEscala) ' Redefino columna hasta

    nFila1 = nalto * nFila1 - nalto * 0.5 ' Redefino fila desde
    nFila2 = nalto * nFila2 ' Redefino fila hasta

    Printer.Line (nColu1, nFila1)-(nColu2, nFila2), , B

End Sub

Public Sub Say_Recuadro(ByVal nFila1 As Single, ByVal nColu1 As Single, ByVal nFila2 As Single, _
                        ByVal cTexto As String, Optional ByVal cRaya As String, Optional nalto)

    '*********************'
    ' Dibuja un texto con raya al costado
    '*********************'

    If cRaya = "" Then
       cRaya = "S"
    End If

    If IsMissing(nalto) Then
       nalto = Impresora.nAltoLetra  ' De acuerdo a letra actual
    End If
    
    nColu1 = Int(Impresora.nAnchoHoja * nColu1 / Impresora.nEscala) ' Redefino columna desde
    nFila1 = nalto * nFila1 ' Redefino fila desde
    nFila2 = nalto * nFila2 ' Redefino fila hasta
    Printer.CurrentX = nColu1 + Impresora.nAnchoLetra * 0.2
    Printer.CurrentY = nFila1
    Printer.Print cTexto

    If cRaya = "S" Then
       nFila1 = nFila1 - nalto * 0.5 ' Redefino fila desde
       Printer.Line (nColu1, nFila1)-(nColu1, nFila2)
    End If

End Sub

Public Sub Raya(ByVal nFila1 As Single, ByVal nColu1 As Single, ByVal nColu2 As Single, Optional nalto)
    '*********************'
    ' Dibuja una raya
    '*********************'
    nColu1 = Int(Impresora.nAnchoHoja * nColu1 / Impresora.nEscala) ' Redefino columna desde
    nColu2 = Int(Impresora.nAnchoHoja * nColu2 / Impresora.nEscala) ' Redefino columna hasta
    If IsMissing(nalto) Then
       nFila1 = Impresora.nAltoLetra * nFila1 + Int(Impresora.nAltoLetra / 2) ' Redefino fila desde
     Else
       nFila1 = nalto * nFila1 + Int(nalto / 2) ' Redefino fila desde
    End If

    Printer.Line (nColu1, nFila1)-(nColu2, nFila1)
End Sub

Public Sub Imagen(ByVal bmpimagen, ByVal y1, ByVal x1)
    '***********************
    ' Dibuja en la impresora la imagen bmpimagen
    '***********************
    Dim nColu1, nFila1
    
    nColu1 = Int(Impresora.nAnchoHoja * x1 / Impresora.nEscala)
    nFila1 = Impresora.nAltoLetra * y1
    Printer.PaintPicture bmpimagen, nColu1, nFila1
End Sub

