Attribute VB_Name = "mdl_senha"
' SENHA.BAS -------------------------------------------------------
'
'   SGP® - Rotina de Kriptografar Senha
'
'   Rotinas: Kriptografa
'            DesKriptografa
'
'   © 1996 by Tasso Teixeira
'   Cerrado Informática
'
'------------------------------------------------------------------
Option Explicit
Function Kriptografa(ByVal pSenha As String) As String
    Dim caracter As Integer
    Dim i As Integer
    Dim total As Integer
    total = Len(pSenha)
    pSenha = UCase(pSenha)
    For i = 1 To total
        caracter = Asc(Mid(pSenha, i, 1)) + 28
        Mid(pSenha, i, 1) = Chr(caracter)
    Next
    Kriptografa = pSenha
End Function
Function DesKriptografa(pSenha As String) As String
    Dim caracter As Integer
    Dim i As Integer
    Dim total As Integer
    total = Len(pSenha)
    For i = 1 To total
        caracter = Asc(Mid(pSenha, i, 1)) - 28
        Mid(pSenha, i, 1) = Chr(caracter)
    Next
    DesKriptografa = pSenha
End Function
