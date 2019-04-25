Attribute VB_Name = "ImpressoraTermica"
Option Explicit

'Variáveis de Impressao
Dim lNomeArquivo As String
Dim lLocalImpressao As Integer
Dim lLinha As String
Public gStringImpTermica As String
Public g_impressora_matricial As Boolean
Public g_tamanho_impressora As Integer

Public gArquivoTMP As TextStream
Public gArquivoTMV As TextStream
Public gArquivoHTML As TextStream

Public Function DefineImpressoraTermicaComoPadrao() As Boolean

DefineImpressoraTermicaComoPadrao = False
Dim Impressora As Printer
Dim Contador As Byte
Contador = 0
For Each Impressora In Printers
    If UCase(Impressora.DeviceName) Like "*TM-T20*" Or UCase(Impressora.DeviceName) Like "*TM-T8*" Or UCase(Impressora.DeviceName) Like "*MP-4*" Or UCase(Impressora.DeviceName) Like "*MP-2*" Or UCase(Impressora.DeviceName) Like "*MP-1*" Then
        Set Printer = Printers(Contador)
        DefineImpressoraTermicaComoPadrao = True
        Exit Function
    End If
    Contador = Contador + 1
Next
'MsgBox "Nenhuma Impressora térmica foi encontrada!"
End Function


Public Sub ImpTermicaAbreRelatorio()
    DefineImpressoraTermicaComoPadrao
    lNomeArquivo = BioCriaImprime
    'seleciona medidas para centímetros
    BioImprime "@@Printer.ScaleMode = 7"
    BioImprime "@@Printer.PaperSize = 1"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    'teste para imprimir letra correta
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    Printer.FontName = "Lucida Console 7cpi"
    BioImprime "@@Printer.FontName = Lucida Console 7cpi"
    BioImprime "@@Printer.CurrentY = 0"
End Sub
Public Sub ImpTermicaImprimeDados(ByVal pLinhaDados As String, ByVal pNegrito As Boolean)
    Dim xNegrito As String
    
    If pNegrito = True Then
        xNegrito = "True"
    Else
        xNegrito = "False"
    End If
    BioImprime "@Printer.Print " & pLinhaDados
    BioImprime "@@Printer.FontBold = " & xNegrito
End Sub
Public Sub ImpTermicaFechaRelatorio(ByVal pNomeRelatorio)
    BioImprime "@Printer.Print  "
    BioImprime "@Printer.Print  "
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    lLocalImpressao = 1
    gStringImpTermica = lLocalImpressao & lNomeArquivo & "|@|" & pNomeRelatorio & "|@|"
    frm_preview.Show 1
End Sub

Public Function BioCriaImprime() As String
    Dim xNomeArquivo As String
    xNomeArquivo = Format(Day(Date), "00") & Format(Time, "hhmmss") & ".TMV"
    Set gArquivoTMV = gArqTxt.CreateTextFile(xNomeArquivo, True)
    Mid(xNomeArquivo, 10, 3) = "TMP"
    Set gArquivoTMP = gArqTxt.CreateTextFile(xNomeArquivo, True)
    BioCriaImprime = xNomeArquivo
End Function
Public Sub BioImprime(ByVal pString As String)
    gArquivoTMP.WriteLine (pString)
    If pString = "@@Printer.NewPage" Or pString = "@@Printer.CurrentY = 0" Then
        gArquivoTMV.WriteLine (" ")
    ElseIf Mid(pString, 1, 2) <> "@@" Then
        gArquivoTMV.WriteLine (Mid(pString, 16, Len(pString) - 15))
    End If
End Sub
Public Sub BioFechaImprime()
    gArquivoTMV.Close
    gArquivoTMP.Close
    Set gArquivoTMV = Nothing
    Set gArquivoTMP = Nothing
End Sub
Public Function RetiraGString(ByVal pNumero As Integer) As String
    Dim xIndex As Integer
    Dim xInicio As Integer
    Dim xNumero As Integer
    
    RetiraGString = ""
    xInicio = 1
    xNumero = 1
    If Len(gStringImpTermica) > 0 Then
        Do Until xIndex > Len(gStringImpTermica)
            xIndex = xIndex + 1
            If Mid(gStringImpTermica, xIndex, 3) = "|@|" Then
                If xNumero = pNumero Then
                    RetiraGString = Mid(gStringImpTermica, xInicio, xIndex - xInicio)
                    Exit Function
                End If
                xIndex = xIndex + 2
                xNumero = xNumero + 1
                xInicio = xIndex + 1
            End If
        Loop
    End If
End Function
Public Function ImprimeTexto(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim i As Integer
    ImprimeTexto = ""
    f_coluna_i = f_coluna_i + 0.08
    f_coluna_f = f_coluna_f - 0.08
    Do
        If Trim(f_string) = "" Then
            Exit Do
        End If
        i = i + 1
        If Printer.TextWidth(Mid(f_string, 1, i)) >= (f_coluna_f - f_coluna_i) Then
            Exit Do
        End If
        If Printer.TextWidth(f_string) = 0 Then
            Exit Do
        End If
        ImprimeTexto = Mid(f_string, 1, i)
        If Len(f_string) = i Then
            Exit Do
        End If
    Loop
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeTexto, f_local
End Function
Public Sub Imprime(ByVal f_coluna As Currency, ByVal f_linha As Currency, ByVal f_string As String, ByVal f_local As Integer)
    If f_local = 0 Then
        'emissao_movimentacao_diaria.Picture1.CurrentX = f_coluna '/ 2
        'emissao_movimentacao_diaria.Picture1.CurrentY = f_linha '/ 2
        'emissao_movimentacao_diaria.Picture1.Print f_string
    Else
        Printer.CurrentX = f_coluna
        Printer.CurrentY = f_linha
        Printer.Print f_string
    End If
End Sub
Public Function ChamaDrive() As Boolean
    On Error GoTo FileError
    
    ChamaDrive = False
    gDrive = ReadINI("LOCAL", "Drive", ArqSgpIni)
    gDiretorioData = ReadINI("LOCAL", "Diretorio BD", ArqSgpIni)
    'lDiretorioAplicativo = ReadINI("LOCAL", "Diretorio Aplicativo", ArqSgpIni)
    
    'gNomeBancoDados = ReadINI("LOCAL", "Nome do Banco de Dados", ArqSgpIni)
    'gNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", ArqSgpIni)
    ChDrive gDrive
    ChDir gDiretorioData
    ChamaDrive = True
    Exit Function

FileError:
    MsgBox "Não foi possível definir a unidade: " & gDrive & vbCrLf & "Se o problema continuar não será possível acessar o banco de dados." & Chr(10) & Chr(10) & "Pode ser que a unidade " & gDrive & " não esteja mapeada na rede." & vbCrLf & ArqSgpIni, vbCritical, "Erro de definição de Unidade!"
End Function


