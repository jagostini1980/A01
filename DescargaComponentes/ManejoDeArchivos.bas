Attribute VB_Name = "ManejoDeArchivos"
Option Explicit
Const ConstIniFile = "version.ini"

Public Servidor As String
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long



Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long

'Esta el siguiente archivo o directorio
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long

'Esta cierra el Handle de búsqueda
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


' Constantes
'------------------------------------------------------------------------------

'Constantes de atributos de archivos
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

'Otras constantes
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1


'UDT
'------------------------------------------------------------------------------

'Estructura para las fechas de los archivos
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Estructura necesaria para la información de archivos
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'-----------------------------------------------------------------------
    'Funciones
'-----------------------------------------------------------------------


'Esta función es para formatear los nombres de archivos y directorios. Elimina los CHR(0)
'------------------------------------------------------------------------
Function Eliminar_Nulos(OriginalStr As String) As String
    
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    Eliminar_Nulos = OriginalStr

End Function

'Esta función es la principal que permite buscar _
 los archivos y listarlos en el ListBox


Function FindFilesAPI(Path As String, _
                      SearchStr As String, _
                      FileCount As Long, _
                      DirCount As Integer, _
                      ListBox As ListBox)


    Dim Filename As String
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer


    If Right(Path, 1) <> "\" Then Path = Path & "\"
        ' Buscamos por mas directorios
        nDir = 0
        ReDim dirNames(nDir)
        Cont = True
        hSearch = FindFirstFile(Path & "*", WFD)
            If hSearch <> INVALID_HANDLE_VALUE Then
                Do While Cont
                    DirName = Eliminar_Nulos(WFD.cFileName)
                    ' Ignore the current and encompassing directories.
                    If (DirName <> ".") And (DirName <> "..") Then
                        ' Check for directory with bitwise comparison.
                            If GetFileAttributes(Path & DirName) _
                                And FILE_ATTRIBUTE_DIRECTORY Then
                                
                                dirNames(nDir) = DirName
                                DirCount = DirCount + 1
                                nDir = nDir + 1
                                ReDim Preserve dirNames(nDir)
                            
                            End If
                    End If
                    Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
                Loop
                
                Cont = FindClose(hSearch)
            
            End If

        hSearch = FindFirstFile(Path & SearchStr, WFD)
        Cont = True
        If hSearch <> INVALID_HANDLE_VALUE Then
            While Cont
                Filename = Eliminar_Nulos(WFD.cFileName)
                    If (Filename <> ".") And (Filename <> "..") Then
                        FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) _
                                                                  + WFD.nFileSizeLow
                        FileCount = FileCount + 1
                        ListBox.AddItem Path & Filename
                    End If
                Cont = FindNextFile(hSearch, WFD) ' Get next file
            Wend
        Cont = FindClose(hSearch)
        End If

        ' Si estos son Sub Directorios......
    'If nDir > 0 Then

        'For i = 0 To nDir - 1
        '    FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", _
                                                SearchStr, FileCount, DirCount, ListBox)
        'Next i
    'End If

End Function

Public Function Obtiene_seccion(Strseccion As String, StrEntrada As String) As String
Dim entry As Variant
Dim buffer As String * 255
Dim ret As Long
Dim Res As String
Dim Path As String * 145
Dim ReturnLength As Integer
Dim SysPath As String

    ReturnLength = GetSystemDirectory(Path, Len(Path))
    SysPath = Trim(Mid(Path, 1, ReturnLength))
    
    'MsgBox App.Path & "\" & ConstIniFile
    
    ret = GetPrivateProfileString(Strseccion, _
    StrEntrada, Default, buffer, Len(buffer) - 1, _
      App.Path & "\" & ConstIniFile)
    
    Res = Mid(buffer, 1, InStr(buffer, Chr(0)) - 1)
    If Trim(Res) = "" Then
        MsgBox "ERROR! Archivo " & ConstIniFile & "NO encontrado!!", vbCritical
        Obtiene_seccion = ""
    End If
    Obtiene_seccion = Res
End Function

Public Function Obtiene_WinDIR() As String
Dim entry As Variant
Dim buffer As String * 145
Dim ret As Long
Dim Res As String
Dim Path As String * 255
Dim ReturnLength As Integer
Dim SysPath As String

    ReturnLength = GetSystemDirectory(Path, Len(Path))
    SysPath = Trim(Mid(Path, 1, ReturnLength))
    Obtiene_WinDIR = SysPath
End Function

Public Function FileExist(ByVal sFile As String) As Boolean
    'comprobar si existe este fichero
    Dim WFD As WIN32_FIND_DATA
    Dim hFindFile As Long

    hFindFile = FindFirstFile(sFile, WFD)
    'Si no se ha encontrado
    If hFindFile = INVALID_HANDLE_VALUE Then
        FileExist = False
    Else
        FileExist = True
        'Cerrar el handle de FindFirst
        hFindFile = FindClose(hFindFile)
    End If

End Function

Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String

    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    sGetINI = Left$(sTemp, nLength)
    
End Function


