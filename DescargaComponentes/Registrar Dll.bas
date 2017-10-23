Attribute VB_Name = "RegistrarDLL"
Declare Function DLLSelfRegister Lib "VB6STKIT.DLL" (ByVal DllName As String) As Integer

Public Function SelfRegisterDLL(DllName As String) As Boolean
Dim liRet As Integer
On Error Resume Next
    liRet = DLLSelfRegister(DllName)
    If liRet = 0 Then
        SelfRegisterDLL = True
    Else
        SelfRegisterDLL = False
    End If
End Function
