Attribute VB_Name = "Fiscal"
Option Explicit
'Declaração da DLL com suas Funções
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AcionaGuilhotinaMFD Lib "BEMAFI32.DLL" (ByVal iTipoCorte As Integer) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FITEF_ImprimeTEF Lib "BEMAFI32.DLL" (ByVal cIdentificacao As String, ByVal cFormaPagamento As String, ByVal cValorCompra As String) As Integer
Public Declare Function Bematech_FI_IniciaModoTEF Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FinalizaModoTEF Lib "BEMAFI32.DLL" () As Integer

Public Declare Sub DLLG2_AdicionaParam Lib "DLLG2.dll" (ByVal Handle As Long, ByVal NomeParam As String, ByVal ValorParam As String, ByVal TipoParam As Long)
Public Declare Function DLLG2_EncerraDriver Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Public Declare Function DLLG2_ExecutaComando Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Comando As String) As Long
Public Declare Function DLLG2_IniciaDriver Lib "DLLG2.dll" (ByVal Canal As String) As Long
Public Declare Function DLLG2_LeRegistrador Lib "DLLG2.dll" (ByVal Handle As Long, ByVal NomeRegistrador As String, ByVal NomeComando As String, ByVal TamNomeComando As Long) As Long
Public Declare Function DLLG2_LimpaParams Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Public Declare Function DLLG2_ListaParams Lib "DLLG2.dll" (ByVal Handle As Long, ByVal LstParams As String, ByVal TamLstParams As Long) As String
Public Declare Function DLLG2_ObtemCodErro Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Public Declare Function DLLG2_ObtemRetornos Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Retornos As String, ByVal TamRetorno As Long) As String
Public Declare Function DLLG2_ObtemNomeLog Lib "DLLG2.dll" (ByVal NomeArquivo As String, ByVal TamNomeArquivo As Long) As String
Public Declare Sub DLLG2_SetaArquivoLog Lib "DLLG2.dll" (ByVal NomeArquivo As String)
Public Declare Function DLLG2_Versao Lib "DLLG2.dll" (ByVal Versao As String, ByVal TamVersao As Long) As String

'Metodos Cupom
Public Declare Function Daruma_FI_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal TextoLivre As String) As Integer
Public Declare Function Daruma_FI_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem As String) As Integer

'Metodos para Recebimentos e Relatorios
Public Declare Function Daruma_FI_LeituraX Lib "Daruma32.dll" () As Integer

'Metodos de Status
Public Declare Function Daruma_FI_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StsCF As String) As Integer
Public Declare Function Daruma_FI_NumeroSerie Lib "Daruma32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Daruma_FI_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer

'Metodos Relatorios Fiscais e Relatorios
Public Declare Function Daruma_FI_AbreRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioGerencial Lib "Daruma32.dll" (ByVal Texto As String) As Integer
Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Daruma_FI_RetornaErroExtendido Lib "Daruma32.dll" (ByVal ErroExtendido As String) As Integer

Public Declare Function Daruma_FI_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorPago As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Daruma_FI_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal Texto As String) As Integer
Public Declare Function Daruma_FI_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer

'Metodos TEF
Public Declare Function Daruma_TEF_FechaRelatorio Lib "Daruma32.dll" () As Integer


'Arquivo INI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global gQtdViasTEF As Integer
Global gNumeroControleSolicitacao As Long
Public gTefResposta As Boolean
Public gTefString As String
Public gNumeroCupom As Long
Public gValorRecebido As String
Public gValorDesconto As String
Public gBandeira As String
Public gConsultaCheque As Boolean
Public gCpfCnpj As String
Public gNomeCliente As String
Public gObservacao1 As String
Public gObservacao2 As String
Public gTextoAntesCV As String
Public gLinhasEntreCV As Integer
Public gFechamentoIniciado As Boolean
Public gDadosProdutos As Variant
Public gDadosTCS As Variant
Public gLegislacaoPermiteIssEcf As Boolean
Public gContadorNaoFiscal As String
Public gCodigoTcsEcf As Integer
Public BemaRetorno As Integer
Public gArqTxt As New FileSystemObject
Public gArquivo As TextStream
Public gArqTxt2 As New FileSystemObject
Public gArquivo2 As TextStream
Public gSQL As String
Public gLinhasEmBloco As Integer
Public gNomeEmpresa As String
Public gValorDescontoConcedido As Currency
Public gTipoDocumentoFiscal As String
Public gDrive As String
Public gDiretorioData As String
Public gTipoDesconto As String
Public gNumeroAutorizacaoPostoAki As String
Public gCodigoColaborador As Integer
Public gNomeColaborador As String
Public gAvaliacaoColaborador As Integer
Public gTrocaOleo As Boolean
Public gPontuacao As Boolean
Public gVersao As String


Public gImpBematech As Boolean
Public gImpSchalter As Boolean
Public gImpMecaf As Boolean
Public gImpQuick As Boolean
Public gImpElgin As Boolean
Public gImpDaruma As Boolean
Public gQuickCanal As Long

Public gNomeGerenciadorPadrao As String
Public gDiretorioReq As String
Public gDiretorioResp As String

Public gString As String

Public Const ArqSgpIni As String = "c:\Cerrado.Net\sgp.ini"

'habilita/desabilida CTRL+ALT+DEL e CTRL+ESC
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
Public Function AbrirArquivo(ByVal pNomeArquivo As String, ByVal pTipoLeitura As String) As Boolean
    Dim i As Integer
    
    AbrirArquivo = False
    If gArqTxt.FileExists(pNomeArquivo) Then
        For i = 1 To 15
            If AbrirArquivo2(pNomeArquivo, pTipoLeitura) Then
                AbrirArquivo = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function AbrirArquivo2(ByVal pNomeArquivo As String, ByVal pTipoLeitura As String) As Boolean
    AbrirArquivo2 = False
    
    On Error GoTo FileError
    
    If UCase(pTipoLeitura) = "LEITURA" Then
        Set gArquivo = gArqTxt.OpenTextFile(pNomeArquivo, ForReading)
    End If
    AbrirArquivo2 = True
    
    Exit Function

FileError:
    Call AguardaTempo(2)
End Function
Public Function CopiarArquivo(ByVal pOrigem As String, ByVal pDestino As String) As Boolean
    Dim i As Integer
    
    CopiarArquivo = False
    If gArqTxt.FileExists(pOrigem) Then
        For i = 1 To 15
            If CopiarArquivo2(pOrigem, pDestino) Then
                CopiarArquivo = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function CopiarArquivo2(ByVal pOrigem As String, ByVal pDestino As String) As Boolean
    
    On Error GoTo FileError
    
    CopiarArquivo2 = False
    Call gArqTxt.CopyFile(pOrigem, pDestino, True)
    CopiarArquivo2 = True
    
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " CopiarArquivo2: Erro n.:" & Err & " - Arquivo Origem: " & pOrigem & " - Arquivo Destino: " & pDestino & " - ErroTexto: " & Error)
    Call AguardaTempo(2)
End Function
Sub CriaLogTEF(xTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    Dim xArqTxt As New FileSystemObject
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "TEF_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "c:\VB5\SGP\DATA\TEF_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If xArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = xArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = xArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (xTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    MsgBox Error
    MsgBox "Erro ao criar LOG TEF: " & xTipoLog, vbInformation, "Erro: CriaLogTEF"
    Exit Sub
End Sub
Sub CriaLogTefEspecial(xTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    Dim xArqTxt As New FileSystemObject
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "TEF_ESPECIAL_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "c:\VB5\SGP\DATA\TEF_ESPECIAL_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If xArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = xArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = xArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (Time & " - " & xTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    MsgBox Error
    MsgBox "Erro ao criar LOG TEF_ESPECIAL: " & xTipoLog, vbInformation, "Erro: CriaLogTefEspecial"
    Exit Sub
End Sub
Sub DisabelCtrlAltDel(bdisabled As Boolean)
    'Dim X As Long
    'X = SystemParametersInfo(97, bdisabled, CStr(1), 0)
End Sub
Public Sub AguardaTempo(ByVal pSegundos As Integer)
    Dim xHoraInicial As Date
    
    'Aguarda x segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= pSegundos
        DoEvents
    Loop
End Sub
Sub CentraForm(unForm As Form)
    unForm.Left = (Screen.Width - unForm.Width) / 2
    unForm.Top = (Screen.Height - unForm.Height) / 2
End Sub
Public Function CriarArquivo(ByVal pNomeArquivo As String) As Boolean
    Dim i As Integer
    
    CriarArquivo = False
    If Not gArqTxt.FileExists(pNomeArquivo) Then
        For i = 1 To 15
            If CriarArquivo2(pNomeArquivo) Then
                CriarArquivo = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function CriarArquivo2(ByVal pNomeArquivo As String) As Boolean
    CriarArquivo2 = False
    
    On Error GoTo FileError
    
    Set gArquivo = gArqTxt.CreateTextFile(pNomeArquivo)
    CriarArquivo2 = True
    
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " CriarArquivo2: Erro n.:" & Err & " - Arquivo: " & pNomeArquivo & " - ErroTexto: " & Error)
    Call AguardaTempo(2)
End Function
Public Function ExcluirArquivo(ByVal pArquivo As String) As Boolean
    Dim i As Integer
    
    ExcluirArquivo = False
    If gArqTxt.FileExists(pArquivo) Then
        For i = 1 To 15
            If ExcluirArquivo2(pArquivo) Then
                ExcluirArquivo = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function ExcluirArquivo2(ByVal pArquivo As String) As Boolean
    ExcluirArquivo2 = False
    
    On Error GoTo FileError
    
    Call gArqTxt.DeleteFile(pArquivo, True)
    ExcluirArquivo2 = True
    
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " ExcluirArquivo2: Erro n.:" & Err & " - Arquivo: " & pArquivo & " - ErroTexto: " & Error)
    Call AguardaTempo(2)
End Function
Public Function fMascaraHora(ByVal pHora As String) As String
    fMascaraHora = ""
    If IsDate(pHora) Then
        fMascaraHora = Format(pHora, "HH:mm:ss")
    End If
End Function
Function fValidaValor(ByVal valor_x As String) As Currency
    Dim x_posicao As Integer
    Dim x_tamanho As Integer
    Dim x_flag As Boolean
    x_flag = False
    If valor_x <> "" Then
        x_tamanho = Len(valor_x)
        x_posicao = Len(valor_x)
        Do Until x_posicao = 0
            If Not IsNumeric(Mid(valor_x, x_posicao, 1)) Then
                If x_flag = False Then
                    Mid(valor_x, x_posicao, 1) = "."
                    x_flag = True
                Else
                    valor_x = " " & Mid(valor_x, 1, x_posicao - 1) & Mid(valor_x, x_posicao + 1, x_tamanho - x_posicao)
                End If
            End If
            x_posicao = x_posicao - 1
        Loop
    End If
    fValidaValor = Val(Trim(valor_x))
End Function
Sub GravaRegistroECF()
    Dim xNumeroSerieEcf As String
    Dim xDataEcf As String
    Dim xHoraEcf As String
    Dim xDataMicro As String
    Dim xHoraMicro As String
    Dim xString As String
    Dim xDados As String
    Dim i As Integer
    Dim xLocalErro As Integer
    
    On Error GoTo FileError
    
    xLocalErro = 1
    xNumeroSerieEcf = Space(15)
    xDataEcf = Space(6)
    xHoraEcf = Space(6)
    xLocalErro = 10
    If gTipoDocumentoFiscal = "" Then
        gTipoDocumentoFiscal = "NFCe"
        xLocalErro = 20
        Call CriaLogTEF(Date & " " & Time & " GravaRegistroECF gTipoDocumentoFiscal estava em branco e definiu NFCe. ->" & gTipoDocumentoFiscal & "<-")
    End If
    xLocalErro = 30
    Call CriaLogTEF(Date & " " & Time & " GravaRegistroECF gTipoDocumentoFiscal= ->" & gTipoDocumentoFiscal & "<-")
    xLocalErro = 40
    If gTipoDocumentoFiscal = "NFCe" Then
        xLocalErro = 50
        xNumeroSerieEcf = 1
        xDataEcf = Format(Date, "dd/mm/yyyy")
        xHoraEcf = Format(Time, "hhmmss")
        xLocalErro = 60
    Else
        xLocalErro = 70
        xDados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", ArqSgpIni)
        xLocalErro = 80
        If xDados = "BEMATECH" Then
            BemaRetorno = Bematech_FI_NumeroSerie(xNumeroSerieEcf)
            BemaRetorno = Bematech_FI_DataHoraImpressora(xDataEcf, xHoraEcf)
            xDataEcf = CDate(Mid(xDataEcf, 1, 2) & "/" & Mid(xDataEcf, 3, 2) & "/20" & Mid(xDataEcf, 5, 2))
        ElseIf xDados = "SCHALTER" Then
        ElseIf xDados = "MECAF" Then
        ElseIf xDados = "QUICK" Then
            xDataEcf = EcfQuickBuscaData
            xHoraEcf = EcfQuickBuscaHora
        ElseIf xDados = "ELGIN" Then
            xNumeroSerieEcf = Space(20)
            BemaRetorno = Elgin_NumeroSerie(xNumeroSerieEcf)
            xDataEcf = Space(6)
            xHoraEcf = Space(6)
            BemaRetorno = Elgin_DataHoraImpressora(xDataEcf, xHoraEcf)
            xDataEcf = CDate(Mid(xDataEcf, 1, 2) & "/" & Mid(xDataEcf, 3, 2) & "/20" & Mid(xDataEcf, 5, 2))
        ElseIf xDados = "DARUMA" Then
            xNumeroSerieEcf = Space(15)
            BemaRetorno = Daruma_FI_NumeroSerie(xNumeroSerieEcf)
            xDataEcf = Space(6)
            xHoraEcf = Space(6)
            BemaRetorno = Daruma_FI_DataHoraImpressora(xDataEcf, xHoraEcf)
            xDataEcf = CDate(Mid(xDataEcf, 1, 2) & "/" & Mid(xDataEcf, 3, 2) & "/20" & Mid(xDataEcf, 5, 2))
        End If
    End If
    xLocalErro = 100
    xDataMicro = Format(Date, "dd/mm/yyyy")
    xHoraMicro = Format(Time, "hh:mm:ss")
    xLocalErro = 120
    Call CriaLogTEF(Date & " " & Time & " GravaRegistroECF NS ECF: " & xNumeroSerieEcf & " - DATA ECF: " & xDataEcf)
    
    'Ms   -> nada
    'retc -> registro, ecf, tef, cerrado
    '.dep -> nada
    Set gArquivo2 = gArqTxt2.CreateTextFile("C:\WINDOWS\SYSTEM\Msretc.dep")
    
    
'    xString = ""
'    For i = 1 To 15
'        If Mid(xNumeroSerieEcf, i, 1) >= "0" And Mid(xNumeroSerieEcf, i, 1) <= "9" Then
'            xString = xString & Mid(xNumeroSerieEcf, i, 1)
'        End If
'    Next
'    gArquivo2.WriteLine (xString)
    xLocalErro = 150
    gArquivo2.WriteLine (xNumeroSerieEcf)
    gArquivo2.WriteLine (xDataEcf & " - " & xHoraEcf)
    gArquivo2.WriteLine (xDataMicro & " - " & xHoraMicro)
    gArquivo2.Close
    Exit Sub
FileError:
    Call CriaLogTEF(Date & " " & Time & " Erro GravaRegistroECF: xLocalErro=" & xLocalErro & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    Exit Sub
End Sub
Function ImprimeEncerramentoCupomFiscal() As Boolean
    Dim xValor As String
    Dim xString As String
    Dim xHoraInicial As Date
    Dim i As Integer
    
    On Error GoTo FileError
    
    ImprimeEncerramentoCupomFiscal = False
    
    If gTipoDocumentoFiscal = "NFCe" Then
        ImprimeEncerramentoCupomFiscal = True
        Exit Function
    End If
    
    'Desconto para o Cupom Fiscal
    If gImpBematech Then
        If gFechamentoIniciado = False Then
            xValor = Mid(Format(fValidaValor(gValorDesconto), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(gValorDesconto), "000000000000.00"), 14, 2)
            BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xValor)
        End If
    ElseIf gImpQuick Then
        If gValorDesconto > 0 Then
            Call EcfQuickAcresceSubTotal(False, 0, gValorDesconto)
        End If
    ElseIf gImpElgin Then
        If gValorDesconto > 0 Then
            BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", Str(gValorDesconto))
        Else
            BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", "0")
        End If
    ElseIf gImpDaruma Then
        If gValorDesconto > 0 Then
            xString = Format(fValidaValor(gValorDesconto), "000000000000.00")
        Else
            xString = "0,00"
        End If
        BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", xString)
    End If
    
    
    'Efetua Forma de Pagamento
    xValor = Mid(Format(fValidaValor(gValorRecebido), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(gValorRecebido), "000000000000.00"), 14, 2)
    If gImpBematech Then
        If UCase(gBandeira) = "TECBAN" Then
            If gConsultaCheque Then
                BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Cheque TecBan   ", xValor)
            Else
                BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Cartao TecBan   ", xValor)
            End If
        ElseIf UCase(gBandeira) = "TCSMART" Then
            BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Ticket Car Smart", xValor)
        Else
            BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Cartao          ", xValor)
        End If
    ElseIf gImpQuick Then
        If gConsultaCheque Then
            BemaRetorno = EcfQuickPagaCupom(0, "CONSULTA CHEQUE", "", gValorRecebido)
        Else
            If EcfQuickPagaCupom(0, "TEF", "", gValorRecebido) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
                Exit Function
            End If
        End If
        'BemaRetorno = EcfQuickEncerraDocumento("", "Cerrado Informatica (62) 3277-1017")
    ElseIf gImpElgin Then
        If gConsultaCheque Then
            BemaRetorno = Elgin_EfetuaFormaPagamentoMFD("CONSULTA CHEQUE", Str(gValorRecebido), "0", "")
        Else
            BemaRetorno = Elgin_EfetuaFormaPagamentoMFD("TEF", xValor, "1", "")
            If BemaRetorno <> 1 Then
                Exit Function
            End If
        End If
    ElseIf gImpDaruma Then
        xValor = Format(fValidaValor(gValorRecebido), "000000000000.00")
        If UCase(gBandeira) = "TECBAN" Then
            If gConsultaCheque Then
                BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Cheque TecBan", xValor, "")
            Else
            'Cartao Credito
                'BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Cartao TecBan   ", xValor, "")
                BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Cartao Credito", xValor, "")
            End If
        ElseIf UCase(gBandeira) = "TCSMART" Then
            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Ticket Car Smart", xValor, "")
        Else
            'BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Cartao          ", xValor, "")
            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Cartao Credito", xValor, "")
        End If
    End If
    
    xString = ""
    'Fecha Cupom Fiscal
    If Len(gCpfCnpj) > 0 Then
        xValor = "CPF/CNPJ:                                       "
        Mid(xValor, 11, 20) = gCpfCnpj
        xString = xString & xValor
    End If
    If Len(gNomeCliente) > 0 Then
        xValor = "NOME..:                                         "
        Mid(xValor, 9, 40) = gNomeCliente
        xString = xString & xValor
    End If
    If Len(gObservacao1) > 0 Then
        xValor = "                                                "
        Mid(xValor, 1, 48) = gObservacao1
        xString = xString & xValor
    End If
    If Len(gObservacao2) > 0 Then
        'xValor = "                                                "
        'Mid(xValor, 1, 48) = gObservacao2
        'xString = xString & xValor
        xString = xString & gObservacao2
    End If
    If gImpBematech Then
        BemaRetorno = Bematech_FI_TerminaFechamentoCupom(xString)
    ElseIf gImpQuick Then
        If EcfQuickEncerraDocumento("", "Cerrado Informatica (62) 3277-1017") Then
            BemaRetorno = 1
        Else
            BemaRetorno = -1
        End If
    ElseIf gImpElgin Then
        If Len(xString) = 0 Then
            xString = "Cerrado Informatica (62) 3277-1017"
        End If
        BemaRetorno = Elgin_TerminaFechamentoCupom(xString)
    ElseIf gImpDaruma Then
        BemaRetorno = Daruma_FI_TerminaFechamentoCupom(xString)
    End If
    'xHoraInicial = Time
    'Do Until DateDiff("s", xHoraInicial, Time) >= 3
    '    DoEvents
    'Loop
    ImprimeEncerramentoCupomFiscal = True
    
    
    Exit Function
FileError:
    MsgBox "Não foi possível imprimir o fechamento do cupom fiscal.", vbCritical, "ImprimeEncerramentoCupomFiscal"
    Exit Function
End Function
Public Function RenomearArquivo(ByVal pOrigem As String, ByVal pDestino As String) As Boolean
    Dim i As Integer
    
    RenomearArquivo = False
    If gArqTxt.FileExists(pOrigem) Then
        For i = 1 To 15
            If RenomearArquivo2(pOrigem, pDestino) Then
                RenomearArquivo = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function RenomearArquivo2(ByVal pOrigem As String, ByVal pDestino As String) As Boolean
    RenomearArquivo2 = False
    
    On Error GoTo FileError
    
    Call gArqTxt.MoveFile(pOrigem, pDestino)
    RenomearArquivo2 = True
    
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " RenomearArquivo2: Erro n.:" & Err & " - Arquivo Origem: " & pOrigem & " - Arquivo Destino: " & pDestino & " - ErroTexto: " & Error)
    Call AguardaTempo(2)
End Function
Public Function ReadINI(Section As String, Key As String, Filename As String) As String
'Filename=nome do arquivo ini
'section=O que esta entre []
'key=nome do que se encontra antes do sinal de igual
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), Filename)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
Function RetiraString(ByVal pNumero As Integer, ByVal pString As String) As String
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
Sub ValidaInteiro(ByRef tecla As Integer)
    Dim char As String
    char = Chr$(tecla)
    If char < "0" Or char > "9" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Public Function VerificaNSImpressoraFiscal(ByVal xDataBase As Date) As Boolean
    Dim xString As String
    Dim xNumeroSerieEcf As String
    Dim xDataEcf As Date
    Dim xDataMicro As Date
    
    VerificaNSImpressoraFiscal = False
    If gArqTxt.FileExists("C:\WINDOWS\SYSTEM\Msretc.dep") Then
        Set gArquivo = gArqTxt.OpenTextFile("C:\WINDOWS\SYSTEM\Msretc.dep", ForReading)
        xString = gArquivo.ReadLine
        xNumeroSerieEcf = xString
        xString = gArquivo.ReadLine
        xDataEcf = CDate(Mid(xString, 1, 10))
        xString = gArquivo.ReadLine
        xDataMicro = CDate(Mid(xString, 1, 10))
        gArquivo.Close
    Else
        Call CriaLogTEF(Date & " " & Time & " VerificaNSImpressoraFiscal Falhou: - NumeroSerieEcf=??????????")
        Exit Function
    End If
    
    
    'Cliente do Marcelo   4708990506781
    If xNumeroSerieEcf = "4708990506781" Then
        VerificaNSImpressoraFiscal = True
    End If
    
    
    
    'Auto Posto Mantiqueira 4708990507288A
    If xNumeroSerieEcf = "4708990507288A" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Pedro Ludovico 4708011120060
    If xNumeroSerieEcf = "4708011120060" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Colorado (Bispo e Siqueira) 4708030559073
    If xNumeroSerieEcf = "4708030559073" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Goiá (Bispo e Batista) 4708990919192
    If xNumeroSerieEcf = "4708990919192" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Maitran (Siqueira Batista e Bispo) 4708020432562
    If xNumeroSerieEcf = "4708020432562" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Cruzeiro (Siqueira & Helrighel) 4708990922087
    If xNumeroSerieEcf = "4708990922087" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Rio Dourados 4708011015814
    If xNumeroSerieEcf = "4708011015814" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Rio Formoso 4708011016055
    If xNumeroSerieEcf = "4708011016055" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Tiradentes 4708990922000
    If xNumeroSerieEcf = "4708990922000" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Cidade Livre 4708991024757
    If xNumeroSerieEcf = "4708991024757" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Atlânta 4708011016114
    If xNumeroSerieEcf = "4708011016114" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Diamante (V & V Auto Posto) 4708030253762
    If xNumeroSerieEcf = "4708030253762" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Vera Cruz 470 *************************************************
    If xNumeroSerieEcf = "470" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto União (Almeida e Mendes Silva) 4708990711570
    If xNumeroSerieEcf = "4708990711570" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Rubi 4708030765414
    If xNumeroSerieEcf = "4708030765414" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Auto Posto Esmeralda 4708010499640
    If xNumeroSerieEcf = "4708010499640" Then
        VerificaNSImpressoraFiscal = True
    End If
    'Desenvolvimento EMULADOR
    If Trim(xNumeroSerieEcf) = "EMULADOR" Then
        VerificaNSImpressoraFiscal = True
    End If
    
    If VerificaNSImpressoraFiscal = False Then
        Call CriaLogTEF(Date & " " & Time & " VerificaNSImpressoraFiscal Falhou: " & " - NumeroSerieEcf=" & xNumeroSerieEcf)
    End If
End Function

Public Function EcfQuickAbreCreditoDebito(ByVal pNomeMeioPagamento As String, ByVal pValor As Currency) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAbreCreditoDebito = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickAdicionaParametro("NomeMeioPagamento", pNomeMeioPagamento, 7) Then
        End If
        If EcfQuickAdicionaParametro("Valor", pValor, 6) Then
        End If
        If EcfQuickExecutaComando("AbreCreditoDebito") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAbreCreditoDebito = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAbreCreditoDebito Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAbreCupomFiscal(ByVal pNome As String, ByVal pEndereco As String, ByVal pCNPJ As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAbreCupomFiscal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Len(pNome) > 30 Then
            pNome = Mid(pNome, 1, 30)
        End If
        If Len(pEndereco) > 80 Then
            pEndereco = Mid(pEndereco, 1, 80)
        End If
        If Len(pCNPJ) > 29 Then
            pCNPJ = Mid(pCNPJ, 1, 29)
        End If
        If pNome <> "" Then
            If EcfQuickAdicionaParametro("NomeConsumidor", pNome, 7) Then
            End If
        End If
        If pEndereco <> "" Then
            If EcfQuickAdicionaParametro("EnderecoConsumidor", pEndereco, 7) Then
            End If
        End If
        If pCNPJ <> "" Then
            If EcfQuickAdicionaParametro("IdConsumidor", pCNPJ, 7) Then
            End If
        End If
        If EcfQuickExecutaComando("AbreCupomFiscal") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAbreCupomFiscal = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAbreCupomFiscal Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAbreCupomNaoFiscal(ByVal pEndereco As String, ByVal pCNPJ As String, ByVal pNome As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAbreCupomNaoFiscal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Len(pNome) > 30 Then
            pNome = Mid(pNome, 1, 30)
        End If
        If Len(pEndereco) > 80 Then
            pEndereco = Mid(pEndereco, 1, 80)
        End If
        If Len(pCNPJ) > 29 Then
            pCNPJ = Mid(pCNPJ, 1, 29)
        End If
        If pNome <> "" Then
            If EcfQuickAdicionaParametro("NomeConsumidor", pNome, 7) Then
            End If
        End If
        If pEndereco <> "" Then
            If EcfQuickAdicionaParametro("EnderecoConsumidor", pEndereco, 7) Then
            End If
        End If
        If pCNPJ <> "" Then
            If EcfQuickAdicionaParametro("IdConsumidor", pCNPJ, 7) Then
            End If
        End If
        If EcfQuickExecutaComando("AbreCupomNaoFiscal") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAbreCupomNaoFiscal = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAbreCupomNaoFiscal Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAbreGaveta() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAbreGaveta = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("TempoAcionamentoGaveta", "10", 4) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("AbreGaveta") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAbreGaveta = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAbreGaveta Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAbreGerencial(ByVal pCodigoGerencial As Integer, ByVal pNomeGerencial As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAbreGerencial = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
'        If pCodigoGerencial > 0 Then
'            If EcfQuickAdicionaParametro("CodGerencial", pCodigoGerencial, 0) Then
'            End If
'        End If
        If pNomeGerencial <> "" Then
            If EcfQuickAdicionaParametro("NomeGerencial", pNomeGerencial, 7) Then
            End If
        End If
        If EcfQuickExecutaComando("AbreGerencial") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAbreGerencial = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAbreGerencial Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAcresceItemFiscal(ByVal pOrdem As Integer, ByVal pCancelaDesconto As Boolean, ByVal pValorAcrescimo As Currency, ByVal pValorDesconto As Currency) As Boolean
    Dim xCodigoErro As Long
    Dim xValor As Currency
    Dim xCancela As Integer
    
    On Error GoTo trata_erro
    
    EcfQuickAcresceItemFiscal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xCancela = 0
        If pCancelaDesconto Then
            xCancela = 1
        End If
        If Not EcfQuickAdicionaParametro("Cancelar", xCancela, 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("NumItem", pOrdem, 4) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xValor = pValorAcrescimo
        If pValorDesconto > 0 Then
            xValor = -pValorDesconto
        End If
        If Not EcfQuickAdicionaParametro("ValorAcrescimo", xValor, 6) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("AcresceItemFiscal") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAcresceItemFiscal = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAcresceItemFiscal Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAcresceSubTotal(ByVal pCancelaDesconto As Boolean, ByVal pValorAcrescimo As Currency, ByVal pValorDesconto As Currency) As Boolean
    Dim xCodigoErro As Long
    Dim xValor As Currency
    Dim xCancela As Integer
    
    On Error GoTo trata_erro
    
    EcfQuickAcresceSubTotal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xCancela = 0
        If pCancelaDesconto Then
            xCancela = 1
        End If
        If Not EcfQuickAdicionaParametro("Cancelar", xCancela, 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xValor = pValorAcrescimo
        If pValorDesconto > 0 Then
            xValor = -pValorDesconto
        End If
        If Not EcfQuickAdicionaParametro("ValorAcrescimo", xValor, 6) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("AcresceSubtotal") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickAcresceSubTotal = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAcresceSubTotal Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickCancelaCupom() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickCancelaCupom = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("CancelaCupom") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickCancelaCupom = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickCancelaCupom Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickDefineGerencial(ByVal pCodigoGerencial As Integer, ByVal pNomeGerencial As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickDefineGerencial = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
'        If pCodigoGerencial > 0 Then
'            If EcfQuickAdicionaParametro("CodGerencial", pCodigoGerencial, 0) Then
'            End If
'        End If
        If pNomeGerencial <> "" Then
            If EcfQuickAdicionaParametro("NomeGerencial", pNomeGerencial, 7) Then
            End If
            If EcfQuickAdicionaParametro("DescricaoGerencial", pNomeGerencial, 7) Then
            End If
        End If
        If EcfQuickExecutaComando("DefineGerencial") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickDefineGerencial = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickDefineGerencial Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickLeituraX() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickLeituraX = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Destino", "I", 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("EmiteLeituraX") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickLeituraX = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickLeituraX Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickPagaCupom(ByVal pCodigoMeioPagamento As Integer, ByVal pNomeMeioPagamento As String, ByVal pTextoAdicional As String, ByVal pValor As Currency) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickPagaCupom = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pNomeMeioPagamento = "" Then
            If Not EcfQuickAdicionaParametro("CodMeioPagamento", pCodigoMeioPagamento, 0) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        Else
            If Not EcfQuickAdicionaParametro("NomeMeioPagamento", pNomeMeioPagamento, 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        pTextoAdicional = Mid(pTextoAdicional, 1, 80)
        If Not EcfQuickAdicionaParametro("TextoAdicional", pTextoAdicional, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Valor", pValor, 6) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("PagaCupom") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Comunicação com ECF!"
            Else
                EcfQuickPagaCupom = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickPagaCupom Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickReducaoZ() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickReducaoZ = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Hora", fMascaraHora(Str(Time)), 3) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("EmiteReducaoZ") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro = 8092 Then
                If Not EcfQuickLimpaParametro Then
                    EcfQuickEncerraDriver
                    Exit Function
                End If
                If EcfQuickExecutaComando("EmiteReducaoZ") Then
                    xCodigoErro = EcfQuickObtemCodigoErro
                    If xCodigoErro > 0 Then
                        MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
                    Else
                        EcfQuickReducaoZ = True
                    End If
                End If
            ElseIf xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickReducaoZ = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickReducaoZ Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickReducaoZPendente() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickReducaoZPendente = False
    xCodigoErro = EcfQuickLeRegistrador("Indicadores", "Inteiro", 4)
    If xCodigoErro >= 16384 Then
        xCodigoErro = xCodigoErro - 16384
    End If
    If xCodigoErro >= 8192 Then
        xCodigoErro = xCodigoErro - 8192
    End If
    If xCodigoErro >= 4096 Then
        xCodigoErro = xCodigoErro - 4096
    End If
    If xCodigoErro >= 2048 Then
        xCodigoErro = xCodigoErro - 2048
    End If
    If xCodigoErro >= 1024 Then
        xCodigoErro = xCodigoErro - 1024
    End If
    If xCodigoErro >= 512 Then
        xCodigoErro = xCodigoErro - 512
    End If
    If xCodigoErro >= 256 Then
        xCodigoErro = xCodigoErro - 256
    End If
    If xCodigoErro >= 128 Then
        xCodigoErro = xCodigoErro - 128
        EcfQuickReducaoZPendente = True
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickReducaoZPendente Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickSemPapel() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickSemPapel = False
    xCodigoErro = EcfQuickLeRegistrador("Indicadores", "Inteiro", 4)
    If xCodigoErro >= 16384 Then
        xCodigoErro = xCodigoErro - 16384
    End If
    If xCodigoErro >= 8192 Then
        xCodigoErro = xCodigoErro - 8192
    End If
    If xCodigoErro >= 4096 Then
        xCodigoErro = xCodigoErro - 4096
    End If
    If xCodigoErro >= 2048 Then
        xCodigoErro = xCodigoErro - 2048
    End If
    If xCodigoErro >= 1024 Then
        xCodigoErro = xCodigoErro - 1024
    End If
    If xCodigoErro >= 512 Then
        xCodigoErro = xCodigoErro - 512
    End If
    If xCodigoErro >= 256 Then
        xCodigoErro = xCodigoErro - 256
        EcfQuickSemPapel = True
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickSemPapel Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickSetaArquivoLog() As Boolean
    On Error GoTo trata_erro
    
    EcfQuickSetaArquivoLog = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        Call DLLG2_SetaArquivoLog("C:\Ecf_Quick.log")
        EcfQuickSetaArquivoLog = True
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickSetaArquivoLog Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickObtemNomeLog() As String
    Dim xRetorno As String
    Dim xString As String
    
    On Error GoTo trata_erro
        
    EcfQuickObtemNomeLog = ""
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        xRetorno = DLLG2_ObtemNomeLog(xString, 0)
        EcfQuickObtemNomeLog = xRetorno
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickObtemNomeLog Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickBuscaData() As String
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickBuscaData = ""
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickAdicionaParametro("NomeData", "Data", 7) Then
            If EcfQuickExecutaComando("LeData") Then
                xCodigoErro = EcfQuickObtemCodigoErro
                If xCodigoErro > 0 Then
                    'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
                Else
                    EcfQuickBuscaData = EcfQuickObtemRetornos()
                End If
            Else
                MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
            End If
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickBuscaData Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAcertaHorarioVerao() As String
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAcertaHorarioVerao = ""
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        'If EcfQuickAdicionaParametro("EntradaHV", "Indicador", 0) Then
            If EcfQuickExecutaComando("AcertaHorarioVerao") Then
                xCodigoErro = EcfQuickObtemCodigoErro
                If xCodigoErro > 0 Then
                    MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
                Else
                    EcfQuickAcertaHorarioVerao = True
                End If
            Else
                MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
            End If
        'End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickAcertaHorarioVerao Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickBuscaHora() As String
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickBuscaHora = ""
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickAdicionaParametro("NomeHora", "Hora", 7) Then
            If EcfQuickExecutaComando("LeHora") Then
                xCodigoErro = EcfQuickObtemCodigoErro
                If xCodigoErro > 0 Then
                    'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
                Else
                    EcfQuickBuscaHora = EcfQuickObtemRetornos()
                End If
            Else
                MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
            End If
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickBuscaHora Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickVendeItem(ByVal pAliquotaICMS As Boolean, ByVal pCodigoAliquota As Integer, ByVal pCodigoDepartamento As Byte, ByVal pCodigoProduto As String, ByVal pNomeDepartamento As String, ByVal pNomeProduto As String, ByVal pPercentualAliquota As Currency, ByVal pPrecoUnitario As Currency, ByVal pQuantidade As Currency, ByVal pUnidade As String) As Boolean
    Dim xCodigoErro As Long
    Dim xAliquotaICMS As Integer
    
    On Error GoTo trata_erro
    
    EcfQuickVendeItem = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pAliquotaICMS Then
            xAliquotaICMS = 1
        Else
            xAliquotaICMS = 0
        End If
        If Not EcfQuickAdicionaParametro("AliquotaICMS", Str(xAliquotaICMS), 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodAliquota", Str(pCodigoAliquota), 4) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodDepartamento", Str(pCodigoDepartamento), 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodProduto", Str(pCodigoProduto), 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pNomeDepartamento <> "" Then
            If Not EcfQuickAdicionaParametro("NomeDepartamento", Str(pNomeDepartamento), 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        If Not EcfQuickAdicionaParametro("NomeProduto", pNomeProduto, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pPercentualAliquota > 0 Then
            If Not EcfQuickAdicionaParametro("PercentualAliquota", Str(pPercentualAliquota), 6) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        If Not EcfQuickAdicionaParametro("PrecoUnitario", pPrecoUnitario, 6) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Quantidade", pQuantidade, 6) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Unidade", pUnidade, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("VendeItem") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickVendeItem = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickVendeItem Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickLeMeioPagamento(ByVal pNome As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickLeMeioPagamento = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("NomeMeioPagamento", pNome, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("LeMeioPagamento") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro = 8014 Then
                EcfQuickLeMeioPagamento = False
            ElseIf xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickLeMeioPagamento = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickLeMeioPagamento Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickLeRegistrador(ByVal pComando As String, ByVal pValor As String, ByVal pTipo As Long) As String
    Dim xCodigoErro As Long
    Dim xRetorno As Long
    Dim xNomeComando As String
    Dim xTamNomeComando As Long
    
    On Error GoTo trata_erro
    
    EcfQuickLeRegistrador = ""
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xNomeComando = "LeRegistrador"
'        If EcfQuickAdicionaParametro(xNomeComando, pValor, pTipo) Then
'            If EcfQuickExecutaComando(xNomeComando) Then
'                xCodigoErro = EcfQuickObtemCodigoErro
'                If xCodigoErro > 0 Then
'                    MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
'                Else
                    xTamNomeComando = 20
                    xRetorno = DLLG2_LeRegistrador(gQuickCanal, pComando, xNomeComando, xTamNomeComando)
                    If xRetorno = 0 Then
                        xNomeComando = Mid(xNomeComando, 1, Len(Trim(xNomeComando)) - 1)
                        If EcfQuickExecutaComando(xNomeComando) Then
                            xCodigoErro = EcfQuickObtemCodigoErro
                            If xCodigoErro > 0 Then
                                'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
                            Else
                                EcfQuickLeRegistrador = EcfQuickObtemRetornos()
                            End If
                        End If
                    End If
'                End If
'            Else
'                MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
'            End If
'        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickLeRegistrador Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickDataReducaoZ() As Date
    Dim xDadosReducaoZ As String
    Dim xData As String
    
    On Error GoTo trata_erro
    
    EcfQuickDataReducaoZ = Date - 1
    xDadosReducaoZ = EcfQuickLeRegistrador("DadosUltimaReducaoZ", "String", 7)
    If Len(xDadosReducaoZ) >= 578 Then
        xData = Mid(xDadosReducaoZ, 573, 2) & "/"
        xData = xData & Mid(xDadosReducaoZ, 575, 2) & "/20"
        xData = xData & Mid(xDadosReducaoZ, 577, 2)
        If IsDate(xData) Then
            EcfQuickDataReducaoZ = CDate(xData)
        End If
    ElseIf Len(xDadosReducaoZ) = 470 Then
        xData = Mid(xDadosReducaoZ, 464, 2) & "/"
        xData = xData & Mid(xDadosReducaoZ, 466, 2) & "/20"
        xData = xData & Mid(xDadosReducaoZ, 468, 2)
        If IsDate(xData) Then
            EcfQuickDataReducaoZ = CDate(xData)
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickDataReducaoZ Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickDefineMeioPagamento(ByVal pNomeMeioPagamento As String, ByVal pDescricaoMeioPagamento As String, ByVal pPermiteVinculado As Boolean) As Boolean
    Dim xCodigoErro As Long
    Dim xPermite As Integer
    
    On Error GoTo trata_erro
    
    EcfQuickDefineMeioPagamento = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        pNomeMeioPagamento = Mid(pNomeMeioPagamento, 1, 16)
        If Not EcfQuickAdicionaParametro("NomeMeioPagamento", pNomeMeioPagamento, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        pDescricaoMeioPagamento = Mid(pDescricaoMeioPagamento, 1, 80)
        If Not EcfQuickAdicionaParametro("DescricaoMeioPagamento", pDescricaoMeioPagamento, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        xPermite = 0
        If pPermiteVinculado Then
            xPermite = 1
        End If
        If Not EcfQuickAdicionaParametro("PermiteVinculado", Str(xPermite), 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("DefineMeioPagamento") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickDefineMeioPagamento = True
            End If
        Else
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickDefineMeioPagamento Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickAdicionaParametro(ByVal pNome As String, ByVal pValor As String, ByVal pTipo As Long) As Boolean
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAdicionaParametro = False
    DLLG2_AdicionaParam gQuickCanal, pNome, pValor, pTipo
    EcfQuickAdicionaParametro = True
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickAdicionaParametro Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickEncerraDocumento(ByVal pOperador As String, ByVal pTextoPromocional As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickEncerraDocumento = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pOperador <> "" Then
            If Not EcfQuickAdicionaParametro("Operador", pOperador, 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        pTextoPromocional = Mid(pTextoPromocional, 1, 492)
        If Not EcfQuickAdicionaParametro("TextoPromocional", pTextoPromocional, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("EncerraDocumento") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                If gTipoDocumentoFiscal <> "NFCe" Then
                    Call Bematech_FI_FinalizaModoTEF
                End If
                'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickEncerraDocumento = True
            End If
        Else
            If gTipoDocumentoFiscal <> "NFCe" Then
                Call Bematech_FI_FinalizaModoTEF
            End If
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickEncerraDocumento Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickEncerraDriver() As Boolean
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickEncerraDriver = False
    xRetorno = DLLG2_EncerraDriver(gQuickCanal)
    If xRetorno >= 0 Then
        EcfQuickEncerraDriver = True
    Else
        MsgBox "Não foi possível fechar a comunicaçao com a ECF Quick.", vbCritical, "Erro de Comunicação!"
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickEncerraDriver Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickExecutaComando(ByVal pComando As String) As Boolean
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickExecutaComando = False
    xRetorno = DLLG2_ExecutaComando(gQuickCanal, pComando)
    If xRetorno = 1 Then
        EcfQuickExecutaComando = True
    Else
        MsgBox "Erro ao executar comando na ECF Quick." & vbCrLf & "Comando=" & pComando & vbCrLf & "Erro=" & xRetorno, vbCritical, "Erro de Comunicação!"
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickExecutaComando Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickImprimeTexto(ByVal pTextoLivre As String) As Boolean
    Dim xCodigoErro As Long
    Dim xString As String
    
    On Error GoTo trata_erro
    
    EcfQuickImprimeTexto = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        pTextoLivre = Mid(pTextoLivre, 1, 492)
        If Len(pTextoLivre) < 48 Then
            xString = Space(48)
            xString = Mid(pTextoLivre, 1, Len(pTextoLivre))
            pTextoLivre = xString
        End If
        If EcfQuickAdicionaParametro("TextoLivre", pTextoLivre, 7) Then
        End If
        If EcfQuickExecutaComando("ImprimeTexto") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                If gTipoDocumentoFiscal <> "NFCe" Then
                    Call Bematech_FI_FinalizaModoTEF
                End If
                'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickImprimeTexto = True
            End If
        Else
            If gTipoDocumentoFiscal <> "NFCe" Then
                Call Bematech_FI_FinalizaModoTEF
            End If
            MsgBox "Não foi possível executar comando na ECF Quick.", vbCritical, "Erro de Comunicação!"
        End If
        EcfQuickEncerraDriver
    End If
    Exit Function

trata_erro:
    If gQuickCanal <> -1 Then
        EcfQuickEncerraDriver
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickImprimeTexto Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickIniciaDriver() As Boolean
    Dim xPortaEcf As String
    
    On Error GoTo trata_erro
    
    EcfQuickIniciaDriver = False
    
    If Len(xPortaEcf) = 0 Then
        xPortaEcf = ReadINI("CUPOM FISCAL", "Porta ECF", ArqSgpIni)
        If Len(xPortaEcf) = 0 Then
            xPortaEcf = "COM2"
            MsgBox "Falta defenir no sgp.ini (Porta ECF=COM?)" & vbCrLf & "O sistema irá definir COM2", vbCritical, "Configuração Ausente no SPG.INI"
        End If
    End If
    gQuickCanal = DLLG2_IniciaDriver(xPortaEcf)
    If gQuickCanal = 0 Then
        EcfQuickIniciaDriver = True
    Else
        MsgBox "Erro ao comunicar com a ECF Quick." & vbCrLf & "Porta de comunicação=" & xPortaEcf, vbCritical, "Erro de Comunicação!"
    End If
    Exit Function

trata_erro:
    If Err.Number = 53 Then
        MsgBox "Erro ao Comunicar com impressora fiscal Quick." & vbCrLf & "A dll de comunicação DLLG2.dll não está instalada neste computador!", vbCritical, "Erro de Comunicação!"
    Else
        MsgBox "Erro ao Comunicar com impressora fiscal Quick.", vbCritical, "Erro de Comunicação!"
    End If
    Call CriaLogTEF(Time & " - Erro: EcfQuickIniciaDriver Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickLimpaParametro() As Boolean
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickLimpaParametro = False
    xRetorno = DLLG2_LimpaParams(gQuickCanal)
    If gQuickCanal >= 0 Then
        EcfQuickLimpaParametro = True
    Else
        MsgBox "provável erro. xRetorno=" & xRetorno
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickLimpaParametro Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickListaParametro() As Boolean
    Dim xRetorno As String
    Dim xString As String
    
    On Error GoTo trata_erro
    
    EcfQuickListaParametro = False
    'If EcfQuickIniciaDriver Then
        xRetorno = DLLG2_ListaParams(gQuickCanal, xString, 10)
        MsgBox "xString=" & xString & vbCrLf & "xRetorno" & xRetorno
        EcfQuickListaParametro = True
    'End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickListaParametro Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickObtemCodigoErro() As Long
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickObtemCodigoErro = 999999
    xRetorno = DLLG2_ObtemCodErro(gQuickCanal)
    EcfQuickObtemCodigoErro = xRetorno
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickObtemCodigoErro Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickObtemRetornos() As String
    Dim xRetorno As String
    Dim xString As String
    
    On Error GoTo trata_erro
    EcfQuickObtemRetornos = ""
    xRetorno = DLLG2_ObtemRetornos(gQuickCanal, xString, 0)
    If Mid(xRetorno, 1, 9) = "ValorData" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 12, 10)
    ElseIf Mid(xRetorno, 1, 9) = "ValorHora" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 12, 8)
    ElseIf Mid(xRetorno, 1, 12) = "ValorInteiro" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 14, Len(xRetorno) - 13)
    ElseIf Mid(xRetorno, 1, 10) = "ValorTexto" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 12, Len(xRetorno) - 11)
    ElseIf Mid(xRetorno, 1, 22) = "ValorNumericoIndicador" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 24, 1)
    ElseIf Mid(xRetorno, 1, 10) = "ValorMoeda" Then
        EcfQuickObtemRetornos = Mid(xRetorno, 12, Len(xRetorno) - 11)
    Else
        EcfQuickObtemRetornos = xRetorno
    End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickObtemRetornos Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickVersao() As String
    Dim xRetorno As String
    Dim xString As String
    
    On Error GoTo trata_erro
    
    EcfQuickVersao = "Não foi possível identificar a versão da ECF Quick"
    'If EcfQuickIniciaDriver Then
        xRetorno = DLLG2_Versao(xString, 0)
        EcfQuickVersao = xRetorno
    'End If
    Exit Function

trata_erro:
    Call CriaLogTEF(Time & " - Erro: EcfQuickVersao Erro=" & Err.Number & " - " & Err.Description)
End Function

