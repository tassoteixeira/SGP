Attribute VB_Name = "FuncoesDiversas"
Option Explicit

Public gSQL As String
Public bdAccess As Boolean
Public bdMySql As Boolean
Public bdSqlServer As Boolean
Public bdSqlServerAzure As Boolean
Public bdOracle As Boolean

Public gArquivoIni As String
Public gArqTxt As New FileSystemObject
Public gArquivoTMP As TextStream
Public gDrive As String
Public gDiretorioData As String
Public gNomeBancoDados As String
Public gConn As adodb.Connection
Public gConnNuvem As adodb.Connection
Public gNomeUsuarioBD As String
Public gSenhaBD As String

'Arquivo INI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function ChamaDrive() As String
    On Error GoTo FileError
    
    gDrive = ReadINI("LOCAL", "Drive", gArquivoIni)
    gDiretorioData = ReadINI("LOCAL", "Diretorio BD", gArquivoIni)
    gNomeBancoDados = ReadINI("LOCAL", "Nome do Banco de Dados", gArquivoIni)
    
'Nome Interno BD
    ChDrive gDrive
    ChDir gDiretorioData
    Exit Function

FileError:
    MsgBox "ERRO NA ROTINA ChamaDrive"
    Exit Function

End Function

Public Sub CriaLogCadastroDll(ByVal pLinhaLog As String)
    Dim xNomeArquivo As String
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: CadastroDll_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "CadastroDll_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set gArquivoTMP = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set gArquivoTMP = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    gArquivoTMP.WriteLine (Time & " - " & pLinhaLog)
    
    'Fecha arquivo texto
    gArquivoTMP.Close
    Set gArquivoTMP = Nothing
    Exit Sub
FileError:
    MsgBox Error
    'MsgBox "Erro ao criar LOG TEF: " & xTipoLog, vbInformation, "Erro: CriaLogCadastroDll"
    Exit Sub
End Sub
Public Sub CriaLogCadastroDll2(ByVal pLinhaDeLog As String, ByVal pMensagemErro As String, ByVal pDetalhe As String)
    Dim xNomeArquivo As String

    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: CadastroDll_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "CadastroDll_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set gArquivoTMP = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set gArquivoTMP = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    gArquivoTMP.WriteLine (Time & " - " & pLinhaDeLog)
    If pMensagemErro <> "" Then
        gArquivoTMP.WriteLine (Time & " - Mensagem de Erro:" & pMensagemErro)
    End If
    If pDetalhe <> "" Then
        gArquivoTMP.WriteLine (Time & " - Detalhe:" & pDetalhe)
    End If
    
    'Fecha arquivo texto
    gArquivoTMP.Close
    Set gArquivoTMP = Nothing
    Exit Sub
FileError:
    MsgBox Error
    'MsgBox "Erro ao criar LOG TEF: " & xTipoLog, vbInformation, "Erro: CriaLogCadastroDll"
    Exit Sub
End Sub
Function preparaArredonda(ByVal pString As String, ByVal pCasasDecimais As Integer) As String
    Dim i As Integer
    Dim xString As String
    preparaArredonda = ""
    If bdAccess Then
        xString = Chr(39) & "0000000000."
        For i = 1 To pCasasDecimais
            xString = xString & "0"
        Next
        xString = xString & Chr(39)
        preparaArredonda = "FORMAT(" & pString & ", " & xString & ")"
    ElseIf bdSqlServer Then
        preparaArredonda = "ROUND(" & pString & ", " & pCasasDecimais & ")"
    End If
End Function
Function preparaData(ByVal xData As Date) As String
    preparaData = ""
    If bdAccess Then
        preparaData = Chr(35) & Format(xData, "mm/dd/yyyy") & Chr(35)
    ElseIf bdSqlServerAzure Then
        preparaData = Chr(39) & Format(xData, "yyyy/mm/dd") & Chr(39)
    ElseIf bdSqlServer Then
        preparaData = Chr(39) & Format(xData, "dd/mm/yyyy") & Chr(39)
    End If
End Function
Function preparaDataHora(ByVal xDataHora As Date) As String
    preparaDataHora = ""
    If bdAccess Then
        preparaDataHora = Chr(35) & Format(xDataHora, "mm/dd/yyyy hh:mm:ss") & Chr(35)
    ElseIf bdSqlServerAzure Then
        preparaDataHora = Chr(39) & Format(xDataHora, "yyyy/mm/dd hh:mm:ss") & Chr(39)
    ElseIf bdSqlServer Then
        preparaDataHora = Chr(39) & Format(xDataHora, "dd/mm/yyyy hh:mm:ss") & Chr(39)
    End If
End Function
Function preparaHora(ByVal xHora As Date) As String
    preparaHora = ""
    If bdAccess Then
        preparaHora = Chr(35) & Format(xHora, "hh:mm:ss") & Chr(35)
    ElseIf bdSqlServer Then
        preparaHora = Chr(39) & Format(xHora, "hh:mm:ss") & Chr(39)
    End If
End Function
Function preparaHoraConsulta(ByVal pCampo As String, ByVal pOperador As String, ByVal pHora As Date) As String
    preparaHoraConsulta = ""
    If bdAccess Then
        preparaHoraConsulta = pCampo & " " & pOperador & " " & Chr(35) & Format(pHora, "hh:mm:ss") & Chr(35)
    ElseIf bdSqlServer Then
        preparaHoraConsulta = "DATEPART(HOUR, " & pCampo & ") " & pOperador & " " & Hour(pHora)
        preparaHoraConsulta = preparaHoraConsulta & " AND DATEPART(MINUTE, " & pCampo & ") " & pOperador & " " & Minute(pHora)
        preparaHoraConsulta = preparaHoraConsulta & " AND DATEPART(SECOND, " & pCampo & ") " & pOperador & " " & Second(pHora)
    End If
End Function
Function preparaTexto(ByVal xTexto As String) As String
    preparaTexto = ""
    If bdAccess Then
        preparaTexto = Chr(39) & xTexto & Chr(39)
    ElseIf bdSqlServer Then
        preparaTexto = Chr(39) & xTexto & Chr(39)
    End If
End Function
Function preparaValor(ByVal pValor As Currency) As String
    Dim xString As String
    preparaValor = ""
    If bdAccess Then
        xString = Format(pValor, "0000000000.0000")
        Mid(xString, 11, 1) = "."
    ElseIf bdSqlServer Then
        xString = Format(pValor, "0000000000.0000")
        Mid(xString, 11, 1) = "."
    End If
    preparaValor = xString
End Function
Function preparaBooleano(ByVal pBooleano As Boolean) As String
    preparaBooleano = ""
    If bdAccess Then
        If pBooleano = True Then
            preparaBooleano = "-1"
        Else
            preparaBooleano = "0"
        End If
    ElseIf bdSqlServer Then
        If pBooleano = True Then
            preparaBooleano = "1"
        Else
            preparaBooleano = "0"
        End If
    End If
End Function
Sub sqlBoolean(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & xString1 & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & xString2
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            If xString1 = "True" Then
                gSQL = gSQL & 1 & xString2
            Else
                gSQL = gSQL & 0 & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "True" Then
                gSQL = gSQL & xString1 & "1"
            Else
                gSQL = gSQL & xString1 & "0"
            End If
        End If
    End If
End Sub
Sub sqlData(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(35) & Format(xString1, "mm/dd/yyyy") & Chr(35) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(35) & Format(xString2, "mm/dd/yyyy") & Chr(35)
            End If
        End If
    ElseIf bdSqlServerAzure Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "yyyy/mm/dd") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "yyyy/mm/dd") & Chr(39)
            End If
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "dd/mm/yyyy") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "dd/mm/yyyy") & Chr(39)
            End If
        End If
    End If
End Sub
Sub sqlDataHora(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(35) & Format(xString1, "mm/dd/yyyy") & " " & Format(xString1, "hh:mm:ss") & Chr(35) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(35) & Format(xString2, "mm/dd/yyyy") & " " & Format(xString2, "hh:mm:ss") & Chr(35)
            End If
        End If
    ElseIf bdSqlServerAzure Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "yyyy/mm/dd") & " " & Format(xString1, "hh:mm:ss") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "yyyy/mm/dd") & " " & Format(xString2, "hh:mm:ss") & Chr(39)
            End If
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "dd/mm/yyyy") & " " & Format(xString1, "hh:mm:ss") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "dd/mm/yyyy") & " " & Format(xString2, "hh:mm:ss") & Chr(39)
            End If
        End If
    End If
End Sub
Sub sqlHora(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(35) & Format(xString1, "hh:mm:ss") & Chr(35) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(35) & Format(xString2, "hh:mm:ss") & Chr(35)
            End If
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "hh:mm:ss") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "hh:mm:ss") & Chr(39)
            End If
        End If
    End If
End Sub
Sub sqlNumero(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & CLng(xString1) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & CLng(xString2)
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            gSQL = gSQL & CLng(xString1) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & CLng(xString2)
        End If
    End If
End Sub
Sub sqlTexto(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(39) & verificaCaracterEspecial(xString1) & Chr(39) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(39) & verificaCaracterEspecial(xString2) & Chr(39)
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(39) & verificaCaracterEspecial(xString1) & Chr(39) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(39) & verificaCaracterEspecial(xString2) & Chr(39)
        End If
    End If
End Sub
Sub sqlValor(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    Dim x_valor As String
    If bdAccess Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "0000000000.0000;-000000000.0000")
            Mid(xString1, 11, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "0000000000.0000;-000000000.0000")
            Mid(xString2, 11, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "0000000000.0000;-000000000.0000")
            Mid(xString1, 11, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "0000000000.0000;-000000000.0000")
            Mid(xString2, 11, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    End If
End Sub
Sub sqlValor4(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    Dim x_valor As String
    If bdAccess Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "00000000.0000")
            Mid(xString1, 9, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "00000000.0000")
            Mid(xString2, 9, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "00000000.0000")
            Mid(xString1, 9, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "00000000.0000")
            Mid(xString2, 9, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    End If
End Sub
Public Function ReadINI(Section As String, Key As String, FileName As String) As String
'Filename=nome do arquivo ini
'section=O que esta entre []
'key=nome do que se encontra antes do sinal de igual
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), FileName)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
Public Sub WriteINI(Section As String, Key As String, Text As String, FileName As String)
'Filename=nome do arquivo ini
'section=O que esta entre []
'key=nome do que se encontra antes do sinal de igual
'text= valor que vem depois do igual
    WritePrivateProfileString Section, Key, Text, FileName
End Sub
Private Function verificaCaracterEspecial(ByVal xString As String) As String
Dim i As Long
    verificaCaracterEspecial = ""
    For i = 1 To Len(xString)
        verificaCaracterEspecial = verificaCaracterEspecial & Mid(xString, i, 1)
        If Mid(xString, i, 1) = "'" Then
            verificaCaracterEspecial = verificaCaracterEspecial & "'"
        End If
    Next
End Function

Public Function RetiraString(ByVal pNumero As Integer, ByVal pString As String) As String
    Dim xIndex As Integer
    Dim xInicio As Integer
    Dim xNumero As Integer
    
    RetiraString = ""
    xInicio = 1
    xNumero = 1
    If Len(pString) > 0 Then
        Do Until xIndex > Len(pString)
            xIndex = xIndex + 1
            If Mid(pString, xIndex, 3) = "|@|" Then
                If xNumero = pNumero Then
                    RetiraString = Mid(pString, xInicio, xIndex - xInicio)
                    Exit Function
                End If
                xIndex = xIndex + 2
                xNumero = xNumero + 1
                xInicio = xIndex + 1
            End If
        Loop
    End If
End Function

