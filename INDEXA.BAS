Attribute VB_Name = "indexa_variaveis"
Option Explicit
Global bd_sgp As Database
Function ChamaDrive() As String
    Dim dados As String
    Dim NumeroArquivo As Integer
    NumeroArquivo = FreeFile
    ChamaDrive = "C"
    On Error GoTo FileError
    Open "C:\ARQLOCAL.TXT" For Input As NumeroArquivo
    Line Input #NumeroArquivo, dados
    If Mid(dados, 1, 6) = "Drive=" Then
        ChamaDrive = Mid(dados, 7, 1)
        ChDrive ChamaDrive
    End If
    Close #NumeroArquivo
    Exit Function
FileError:
    Exit Function
End Function

