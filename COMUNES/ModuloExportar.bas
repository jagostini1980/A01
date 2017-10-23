Attribute VB_Name = "ModuloExportar"
Option Explicit

Public Sub ArmarExcelGrid(Dialogo As CommonDialog)
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
    'MousePointer = vbHourglass
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
'    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub EncabezadoExcelGrid(ex As Excel.Application, ByVal Titulo As String, ByVal FilaInicial As Integer, Columnas As Integer)
Dim col As Integer
Dim i As Long
    With ex
    
        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = Titulo
        .Range("A1:H1").Select
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
            .Size = 18
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Bold = True
        End With
        
        .Range("A" & FilaInicial, LetraColumna(Columnas) & FilaInicial).Select
        '.Rows(Trim(Str(FilaInicial)) & ":" & Trim(Str(FilaInicial))).Select
        
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
        
        '.Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & LetraColumna(Listado.ColCount - 1) & Trim(Str(FilaInicial))).Select
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

Public Sub DatosExcelGrid(ex As Excel.Application, listado As SGGrid, ByVal FilaInicial As Integer, ByRef FilasVisibles As Long, Optional Total As Boolean = False)
Dim Fila As Long
Dim col As Integer
Dim LetraCol As String
Dim IndexCol As Integer
Dim i As Integer
Dim Row As SGRow

    Fila = FilaInicial
    col = 1
    With ex
        For Each Row In listado.Rows
            LetraCol = "A"
            IndexCol = 1
            If Not Row.Hidden Then
                FilasVisibles = FilasVisibles + 1
                For col = 1 To listado.ColCount - 1
                   
                   If Not Row.Cells(col).Column.Hidden Then
                      If IsNumeric(Row.Cells(col).Value) And Row.Cells(col).Value <> "" And col <> 1 Then
                         .Range(LetraCol & Trim(Fila)).Value = Replace(FormatNumber(Row.Cells(col).Value, 2, vbUseDefault, vbUseDefault, vbFalse), ",", ".")
                         'Replace(ValN(Row.Cells(col).Value), ",", ".")
                         .Range(LetraCol & Trim(Fila)).Select
                         .Selection.NumberFormat = "0.00"
                      Else
                          If Mid$(Row.Cells(col).Value, 3, 1) = "/" And Mid$(Row.Cells(col).Value, 6, 1) = "/" Then
                             .Range(LetraCol & Trim(Fila)).Value = Month(Row.Cells(col).Value) & "/" & Day(Row.Cells(col).Value) & "/" & Year(Row.Cells(col).Value)
                          Else
                             .Range(LetraCol & Trim(Fila)).Value = Row.Cells(col).Value
                          End If
                      End If
                      IndexCol = IndexCol + 1
                      LetraCol = LetraColumna(IndexCol)
                     
                   End If
                Next
                Fila = Fila + 1
            End If
        Next
    End With
    FilasVisibles = FilasVisibles - 2
    If Total Then
        ex.Range("A" & Trim(Fila - 1), LetraCol & Trim(Fila - 1)).Font.Bold = True
    End If
End Sub

Public Sub GuardarPlanillaGrid(ex As Excel.Application, NombreArchivo As String, Filtro As Integer)
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

Public Sub FormatearExcelGrid(ex As Excel.Application, ByVal FilaInicial As Integer, Filas As Long, Columnas As Integer, Optional AlterColor As Long)
    With ex
        'ESTOS SON LOS BORDES QUE LE PONGO A LA PLANILLA
        'BORDES DEL ENCABEZADO
        
        .Range(LetraColumna(1) & "2:" & LetraColumna(Columnas) & "2").Select
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        'BORDES DE LOS DATOS
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & _
            LetraColumna(Columnas) & Filas + FilaInicial).Select
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
        'por último, hago que repita el encabezado por cada página
        With .ActiveSheet.PageSetup
            .PrintTitleRows = "$1:$" & Trim(Str(FilaInicial))
            .PrintTitleColumns = ""
            .RightHeader = "Página &P"
        End With
        
        .Range("A" & FilaInicial, LetraColumna(Columnas) & FilaInicial).Interior.Color = &HC0C0C0

        If Not IsMissing(AlterColor) Then
            Dim i As Long
            For i = FilaInicial + 1 To Filas + FilaInicial
              If i Mod 2 = 0 Then
                .Range("A" & i, LetraColumna(Columnas) & i).Interior.Color = AlterColor
              End If
            Next
        End If
    End With
End Sub



