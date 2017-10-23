Attribute VB_Name = "ModuloExportarCta"
Option Explicit

Public Sub ArmarExcelCta(Dialogo As CommonDialog)
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
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
'    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub DatosExcelCta(ex As Excel.Application, listado As TreeView, ByVal FilaInicial As Integer)
Dim Fila As Long
Dim col As Integer
Dim i As Integer
    Fila = FilaInicial + 1
    col = 1
    With ex
        For i = 1 To listado.Nodes.Count
            If listado.Nodes(i).Checked Then
            .Range(LetraColumna(1) & Trim(Fila)).Value = listado.Nodes(i).Text
            Fila = Fila + 1
            End If
        Next
    End With
End Sub

Public Sub GuardarPlanillaCta(ex As Excel.Application, NombreArchivo As String, Filtro As Integer)
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

Public Sub EncabezadoExcelCta(ex As Excel.Application, ByVal Titulo As String, ByVal FilaInicial As Integer)
Dim col As Integer
Dim i As Long
    With ex
    
        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = Titulo
        .Range("A1:F1").Select
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
        For col = 1 To 1 'Listado.ColumnHeaders.Count
            .Range(LetraColumna(col) & Trim(FilaInicial)).Select
            With .ActiveCell
                .FormulaR1C1 = "Cuentas Contables" 'Listado.ColumnHeaders(col).Text
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
        
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & LetraColumna(1) & Trim(Str(FilaInicial))).Select
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

Public Sub FormatearExcelCta(ex As Excel.Application, listado As TreeView, ByVal FilaInicial As Integer, Optional AlterColor As Long)
  Dim i As Integer
  Dim Alter As Boolean
  Dim Fila As Integer
    With ex
        .Range("A" & FilaInicial, LetraColumna(1) & FilaInicial).Interior.Color = &HC0C0C0
        Fila = 1
        Alter = True
        
        If Not IsMissing(AlterColor) Then
            For i = FilaInicial + 1 To FilaInicial + listado.Nodes.Count
              If listado.Nodes(i - FilaInicial).Checked Then
                If Alter Then .Range("A" & FilaInicial + Fila, LetraColumna(1) & FilaInicial + Fila).Interior.Color = AlterColor
                
                Alter = Not Alter
                Fila = Fila + 1
              End If
            Next
        End If
        'ESTOS SON LOS BORDES QUE LE PONGO A LA PLANILLA
        'BORDES DEL ENCABEZADO
        
        .Range(LetraColumna(1) & "2:" & LetraColumna(1) & "2").Select
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        'BORDES DE LOS DATOS
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & _
               LetraColumna(1) & Fila + FilaInicial - 1).Select
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
    End With
End Sub

Public Function LetraColumnaCta(col As Integer) As String
Dim Columnas As String
    
    Columnas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If col <= 26 Then
        LetraColumnaCta = Mid$(Columnas, col, 1)
    Else
        LetraColumnaCta = LetraColumna((col - 1) \ 26) & Mid$(Columnas, IIf(col Mod 26 = 0, 26, col Mod 26), 1)
    End If
End Function

Public Sub FormatearExcel(ex As Excel.Application, listado As ListView, ByVal FilaInicial As Integer, Optional AlterColor As Long)
  Dim i As Integer
        
    With ex
        .Range("A" & FilaInicial, LetraColumna(listado.ColumnHeaders.Count - 1) & FilaInicial).Interior.Color = &HC0C0C0

        If Not IsMissing(AlterColor) Then
            For i = FilaInicial + 1 To FilaInicial + listado.ListItems.Count
              .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count) & i).Font.Color = listado.ListItems(i - FilaInicial).ForeColor

              If i Mod 2 = 0 Then
                 .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count - 1) & i).Interior.Color = AlterColor
              End If
                If .Range("A" & i).Value = "SubTotal" Or .Range("A" & i).Value = "Totales ==>" Then
                   .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count - 1) & i).Font.Bold = True
                   '.Range("A" & i, LetraColumna(Listado.ColumnHeaders.Count) & i).Font.Color = Listado.ForeColor
                End If
            Next
        End If
        'ESTOS SON LOS BORDES QUE LE PONGO A LA PLANILLA
        'BORDES DEL ENCABEZADO
        
        .Range(LetraColumna(1) & "2:" & LetraColumna(listado.ColumnHeaders.Count - 1) & "2").Select
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        'BORDES DE LOS DATOS
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & _
               LetraColumna(listado.ColumnHeaders.Count - 1) & listado.ListItems.Count + FilaInicial - 1).Select
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
    End With
End Sub

Public Sub FormatearExcelConTotal(ex As Excel.Application, listado As ListView, ByVal FilaInicial As Integer, Optional AlterColor As Long)
  Dim i As Integer
        
    With ex
        .Range("A" & FilaInicial, LetraColumna(listado.ColumnHeaders.Count - 1) & FilaInicial).Interior.Color = &HC0C0C0

        If Not IsMissing(AlterColor) Then
            For i = FilaInicial + 1 To FilaInicial + listado.ListItems.Count
              .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count) & i).Font.Color = listado.ListItems(i - FilaInicial).ForeColor

              If i Mod 2 = 0 Then
                 .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count - 1) & i).Interior.Color = AlterColor
              End If
                If .Range("A" & i).Value = "SubTotal" Or .Range("A" & i).Value = "Totales ==>" Then
                   .Range("A" & i, LetraColumna(listado.ColumnHeaders.Count - 1) & i).Font.Bold = True
                   '.Range("A" & i, LetraColumna(Listado.ColumnHeaders.Count) & i).Font.Color = Listado.ForeColor
                End If
            Next
        End If
        'ESTOS SON LOS BORDES QUE LE PONGO A LA PLANILLA
        'BORDES DEL ENCABEZADO
        
        .Range(LetraColumna(1) & "2:" & LetraColumna(listado.ColumnHeaders.Count - 1) & "2").Select
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        'BORDES DE LOS DATOS
        .Range(LetraColumna(1) & Trim(Str(FilaInicial)) & ":" & _
               LetraColumna(listado.ColumnHeaders.Count - 1) & listado.ListItems.Count + FilaInicial).Select
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
    End With
End Sub





