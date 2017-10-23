Attribute VB_Name = "Habilita_Menues"
Public Sub HabilitarMenus(Usuario As String, Conec As String, Formulario As Form)
  Dim C
  Dim RsGruposAModulos As New ADODB.Recordset
    RsGruposAModulos.CursorLocation = adUseClient
    RsGruposAModulos.CursorType = adOpenKeyset
    
    RsGruposAModulos.Open "SpACTraerAccesos @Usuario='" + Usuario + "'", Conec, , adLockBatchOptimistic
    For Each C In Formulario.Controls
       If TypeOf C Is Menu Then
         If C.Caption <> "-" Then
            RsGruposAModulos.Find "M_Modulo='" & C.Name + "'", , , 1
           If Not RsGruposAModulos.EOF Then
              C.Enabled = True
           Else
              C.Enabled = False
           End If
         End If
       End If
    Next
    RsGruposAModulos.Close
    Set RsGruposAModulos = Nothing
'En esta sección del código me habilita los menues que estan siempre hablilidatos
    Formulario.A011000.Enabled = True
    'Formulario.A012000.Enabled = True
    'Formulario.A013000.Enabled = True
    'Formulario.A014000.Enabled = True
    'Formulario.A015000.Enabled = True
    Formulario.A017000.Enabled = True
    Formulario.A019999.Enabled = True
    
    Formulario.A01Sistema.Enabled = True
End Sub
