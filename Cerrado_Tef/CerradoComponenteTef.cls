VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CerradoComponenteTef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lArqTxt As New FileSystemObject

Public Function ImprimeTEF(ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pNumeroCupom As Long, ByVal pValorRecebido As String, ByVal pQtdViasTef As Integer, ByVal pBandeira As String) As Boolean
On Error GoTo trata_erro
    ImprimeTEF = False
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gNumeroCupom = pNumeroCupom
    gValorRecebido = pValorRecebido
    gQtdViasTEF = pQtdViasTef
    gBandeira = pBandeira
    NomeiaDiretorioDaBandeira
    
    gTefString = "ImprimeTEF"
    analizador_tef.Show 1
    If gTefResposta Then
        ImprimeTEF = True
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "ImprimeTEF"
End Function
Public Function SolicitacaoADM(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pQtdViasTef As Integer, ByVal pBandeira As String, ByVal pLinhasEntreCV As Integer, ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String) As Boolean
On Error GoTo trata_erro
    SolicitacaoADM = False
    Call DisabelCtrlAltDel(True)
    
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    Call CriaLogTEF(Date & " " & Time & " SolicitacaoADM: " & " - pQtdViasTef: " & pQtdViasTef & " - pBandeira: " & pBandeira)
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gQtdViasTEF = pQtdViasTef
    gBandeira = pBandeira
    gLinhasEntreCV = pLinhasEntreCV
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gTipoDesconto = ""
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    NomeiaDiretorioDaBandeira
    
    gTefString = "SolicitacaoADM"
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoADM = True
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoADM"
End Function
Public Function SolicitacaoAlteraPrecoTCS(ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String) As Boolean
On Error GoTo trata_erro
    SolicitacaoAlteraPrecoTCS = False
    Call DisabelCtrlAltDel(True)
    Call CriaLogTEF(Date & " " & Time & " SolicitacaoAlteraPrecoTCS")
    gBandeira = "TCSMART"
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    gTipoDesconto = ""
    NomeiaDiretorioDaBandeira
    
    gTefString = "SolicitacaoAlteraPrecoTCS"
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoAlteraPrecoTCS = True
    End If
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoAlteraPrecoTCS"
End Function
Public Function SolicitacaoCNC(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pQtdViasTef As Integer, ByVal pBandeira As String, ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String) As Boolean
On Error GoTo trata_erro
    SolicitacaoCNC = False
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gQtdViasTEF = pQtdViasTef
    gBandeira = pBandeira
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    gTipoDesconto = ""
    NomeiaDiretorioDaBandeira
    
    gTefString = "SolicitacaoCNC"
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoCNC = True
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoCNC"
End Function
Public Function SolicitacaoConsultaCH(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pNumeroCupom As Long, ByVal pValorRecebido As String, ByVal pQtdViasTef As Integer, ByVal pBandeira As String, ByVal pCpfCnpj As String, ByVal pNomeCliente As String, ByVal pObservacao1 As String, ByVal pObservacao2 As String) As Boolean
On Error GoTo trata_erro
    SolicitacaoConsultaCH = False
    Call DisabelCtrlAltDel(True)
    
    Call CriaLogTEF(Date & " " & Time & " SolicitacaoConsultaCH: " & " - pQtdViasTef: " & pQtdViasTef & " - pBandeira: " & pBandeira)
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gNumeroCupom = pNumeroCupom
    gValorRecebido = pValorRecebido
    gQtdViasTEF = pQtdViasTef
    gBandeira = pBandeira
    gCpfCnpj = pCpfCnpj
    gNomeCliente = pNomeCliente
    gObservacao1 = pObservacao1
    gObservacao2 = pObservacao2
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gTipoDesconto = ""
    NomeiaDiretorioDaBandeira
    
    gTefString = "SolicitacaoConsultaCH"
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoConsultaCH = True
    End If
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoConsultaCH"
End Function
Public Function SolicitacaoDesconto(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pNumeroCupom As Long, ByVal pValorVendaBruta As String, ByVal pQtdViasTef As Integer, ByVal pProdutos As String, ByVal pLinhasEntreCV As Integer, ByVal pTextoAntesCV As String, ByVal pTipoDesconto As String, ByVal pNumeroAutorizacaoPostoAki As String, ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String, ByVal pTrocaOleo As Boolean, ByVal pPontuacao As Boolean) As Currency
On Error GoTo trata_erro
    SolicitacaoDesconto = 0
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gNumeroCupom = pNumeroCupom
    gValorRecebido = pValorVendaBruta
    gValorDesconto = 0
    gValorDescontoConcedido = 0
    gQtdViasTEF = pQtdViasTef
    gBandeira = "TEFCERRADO"
    gConsultaCheque = False
    gCpfCnpj = ""
    gNomeCliente = ""
    gObservacao1 = ""
    gTextoAntesCV = pTextoAntesCV
    If Len(RetiraString(2, pProdutos)) > 0 Then
        gDadosProdutos = Split(pProdutos, vbCrLf)
        If Len(RetiraString(2, gDadosProdutos(0))) = 0 Then
            gObservacao2 = gDadosProdutos(0)
        End If
    Else
        gObservacao2 = pProdutos
    End If
    gLinhasEntreCV = pLinhasEntreCV
    gTefString = "SolicitacaoDSC"
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gTipoDesconto = pTipoDesconto
    gNumeroAutorizacaoPostoAki = pNumeroAutorizacaoPostoAki
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    gTrocaOleo = pTrocaOleo
    gPontuacao = pPontuacao
    NomeiaDiretorioDaBandeira
    
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoDesconto = gValorDescontoConcedido / 100
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoDesconto"
End Function
Public Function SolicitacaoTEF(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pNumeroCupom As Long, ByVal pValorRecebido As String, ByVal pValorDesconto As String, ByVal pQtdViasTef As Integer, ByVal pBandeira As String, ByVal pConsultaCheque As Boolean, ByVal pCpfCnpj As String, ByVal pNomeCliente As String, ByVal pObservacao1 As String, ByVal pObservacao2 As String, ByVal pLinhasEntreCV As Integer, ByVal pTextoAntesCV As String, ByVal pFechamentoIniciado As Boolean, ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String) As Boolean
On Error GoTo trata_erro
    SolicitacaoTEF = False
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gNumeroCupom = pNumeroCupom
    gValorRecebido = pValorRecebido
    gValorDesconto = pValorDesconto
    gQtdViasTEF = pQtdViasTef
    gBandeira = pBandeira
    gConsultaCheque = pConsultaCheque
    gCpfCnpj = pCpfCnpj
    gNomeCliente = pNomeCliente
    gObservacao1 = pObservacao1
    gTextoAntesCV = pTextoAntesCV
    If Len(RetiraString(2, pObservacao2)) > 0 Then
        gDadosProdutos = Split(pObservacao2, vbCrLf)
        If Len(RetiraString(2, gDadosProdutos(0))) = 0 Then
            gObservacao2 = gDadosProdutos(0)
        End If
    Else
        gObservacao2 = pObservacao2
    End If
    gLinhasEntreCV = pLinhasEntreCV
    gFechamentoIniciado = pFechamentoIniciado
    gTefString = "SolicitacaoTEF"
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gTipoDesconto = ""
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    NomeiaDiretorioDaBandeira
    
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoTEF = True
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoTEF"
End Function
Public Function SolicitacaoTefTCS(ByVal pTipoDocumentoFiscal As String, ByVal pNumeroControleSolicitacaoTEF As Long, ByVal pNumeroCupom As Long, ByVal pValorRecebido As String, ByVal pValorDesconto As String, ByVal pQtdViasTef As Integer, ByVal pDadosTCS As String, ByVal pLegislacaoPermiteIssEcf As Boolean, ByVal pCodigoTcsEcf As Integer, ByVal pContadorNaoFiscal As String, ByVal pCpfCnpj As String, ByVal pNomeCliente As String, ByVal pObservacao1 As String, ByVal pObservacao2 As String, ByVal pLinhasEntreCV As Integer, ByVal pTextoAntesCV As String, ByVal pCodigoColaborador As Integer, ByVal pNomeColaborador As String) As Boolean
Dim i As Integer

On Error GoTo trata_erro
    
    SolicitacaoTefTCS = False
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gNumeroControleSolicitacao = pNumeroControleSolicitacaoTEF
    gNumeroCupom = pNumeroCupom
    gValorRecebido = pValorRecebido
    gValorDesconto = pValorDesconto
    gQtdViasTEF = pQtdViasTef
    gBandeira = "TCSMART"
    gLegislacaoPermiteIssEcf = pLegislacaoPermiteIssEcf
    gCodigoTcsEcf = pCodigoTcsEcf
    gContadorNaoFiscal = pContadorNaoFiscal
    gDadosTCS = Split(pDadosTCS, "|@@|")
    For i = 0 To UBound(gDadosTCS) - 1
        gDadosTCS(i) = gDadosTCS(i) & "|@|"
    Next
    gConsultaCheque = False
    gCpfCnpj = pCpfCnpj
    gNomeCliente = pNomeCliente
    gObservacao1 = pObservacao1
    gObservacao2 = pObservacao2
    gTextoAntesCV = pTextoAntesCV
    gLinhasEntreCV = pLinhasEntreCV
    gTefString = "SolicitacaoTEF"
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    gTipoDesconto = ""
    gCodigoColaborador = pCodigoColaborador
    gNomeColaborador = pNomeColaborador
    gAvaliacaoColaborador = -1
    NomeiaDiretorioDaBandeira
    
    analizador_tef.Show 1
    If gTefResposta Then
        SolicitacaoTefTCS = True
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Function

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SolicitacaoTEF"
End Function
Public Sub VerificaPendencia(ByVal pTipoDocumentoFiscal As String)
    Dim xNomeBandeira As String
    Dim xString As String
    
On Error GoTo trata_erro
    Call DisabelCtrlAltDel(True)
    'BemaRetorno = Bematech_FI_AbrePortaSerial() 'TerraTef
    gTipoDocumentoFiscal = pTipoDocumentoFiscal
    Call CriaLogTEF(Date & " " & Time & " VerificaPendencia gTipoDocumentoFiscal= ->" & gTipoDocumentoFiscal & "<-")
    If gTipoDocumentoFiscal = "" Then
        gTipoDocumentoFiscal = "NFCe"
        Call CriaLogTEF(Date & " " & Time & " VerificaPendencia gTipoDocumentoFiscal estava em branco e definiu NFCe. ->" & gTipoDocumentoFiscal & "<-")
    End If
    Call GravaRegistroECF
    
    xNomeBandeira = ""
    If lArqTxt.FileExists("C:\Backup_SGP\IntPos.001") Then
        Set gArquivo = gArqTxt.OpenTextFile("C:\Backup_SGP\IntPos.001", ForReading)
        Do Until gArquivo.AtEndOfStream
            xString = gArquivo.ReadLine
            If Mid(xString, 1, 7) = "010-000" Then
                xNomeBandeira = Mid(xString, 11, Len(xString) - 10)
                Exit Do
            End If
        Loop
        gArquivo.Close
    End If
    
    If xNomeBandeira = "HIPERTEF" Then
    ElseIf xNomeBandeira = "VISANET" Or xNomeBandeira = "REDECARD" Or xNomeBandeira = "AMEX" Then
        Call gArqTxt.CopyFile("C:\Backup_SGP\IntPos.001", "C:\TEF_DIAL\RESP\IntPos.001", True)
    ElseIf xNomeBandeira = "TEFCERRADO" Then
        Call gArqTxt.CopyFile("C:\Backup_SGP\IntPos.001", "C:\TefCerrado\RESP\IntPos.001", True)
    End If
    
    If lArqTxt.FileExists("C:\HiperTEF\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "HIPERTEF"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "Outras"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\TEF_DISC\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "TecBan"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\TCS\RX\IntTcs.001") Then
        gBandeira = "TCSMART"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\SMARTEF\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "SMARTEF"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\SUPERTEF\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "SUPERTEF"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\HiperTEF\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "HIPERTEF"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\CardTech_NEUS\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "PAGCARD"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\TefCerrado\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "TEFCERRADO"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\Tef_Neus\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "TEFNEUS"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    ElseIf lArqTxt.FileExists("C:\GetNet\RESP\IntPos.001") Then
        'MsgBox "Cancela NCN"
        gBandeira = "GODCARD"
        gTefString = "SolicitacaoNCN"
        NomeiaDiretorioDaBandeira
        analizador_tef.Show 1
    End If
    'BemaRetorno = Bematech_FI_FechaPortaSerial() 'TerraTef
    Call DisabelCtrlAltDel(False)
    Exit Sub

trata_erro:
    Call DisabelCtrlAltDel(False)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "VerificaPendencia"
End Sub
Private Sub NomeiaDiretorioDaBandeira()
    If UCase(gBandeira) = "TECBAN" Then
        gNomeGerenciadorPadrao = "C:\TEF_DISC\tef_disc.exe"
        gDiretorioReq = "C:\TEF_DISC\REQ\"
        gDiretorioResp = "C:\TEF_DISC\RESP\"
    ElseIf UCase(gBandeira) = "SMARTEF" Then
        gNomeGerenciadorPadrao = ""
        gDiretorioReq = "C:\SMARTEF\REQ\"
        gDiretorioResp = "C:\SMARTEF\RESP\"
    ElseIf UCase(gBandeira) = "SUPERTEF" Then
        gNomeGerenciadorPadrao = ""
        gDiretorioReq = "C:\SUPERTEF\REQ\"
        gDiretorioResp = "C:\SUPERTEF\RESP\"
    ElseIf UCase(gBandeira) = "HIPERTEF" Or UCase(gBandeira) = "HIPERCARD" Then
        gNomeGerenciadorPadrao = ""
        gDiretorioReq = "C:\HiperTEF\REQ\"
        gDiretorioResp = "C:\HiperTEF\RESP\"
    ElseIf UCase(gBandeira) = "PAGCARD" Then
        gNomeGerenciadorPadrao = "C:\CardTech_NEUS\SAC.exe"
        gDiretorioReq = "C:\CardTech_NEUS\REQ\"
        gDiretorioResp = "C:\CardTech_NEUS\RESP\"
    ElseIf UCase(gBandeira) = "TEFCERRADO" Then
        gNomeGerenciadorPadrao = "C:\TefCerrado\TefCerrado.exe"
        gDiretorioReq = "C:\TefCerrado\REQ\"
        gDiretorioResp = "C:\TefCerrado\RESP\"
    ElseIf UCase(gBandeira) = "TCSMART" Then
        gNomeGerenciadorPadrao = ""
        gDiretorioReq = "C:\TCS\TX\"
        gDiretorioResp = "C:\TCS\RX\"
    ElseIf UCase(gBandeira) = "TEFNEUS" Then
        gNomeGerenciadorPadrao = "C:\Tef_Neus\TefNeus.exe"
        gDiretorioReq = "C:\Tef_Neus\REQ\"
        gDiretorioResp = "C:\Tef_Neus\RESP\"
    ElseIf UCase(gBandeira) = "GODCARD" Then
        gNomeGerenciadorPadrao = "C:\GetNet\GetNet.exe"
        gDiretorioReq = "C:\GetNet\REQ\"
        gDiretorioResp = "C:\GetNet\RESP\"
    Else
        gNomeGerenciadorPadrao = "C:\tef_dial\tef_dial.exe"
        gDiretorioReq = "C:\TEF_DIAL\REQ\"
        gDiretorioResp = "C:\TEF_DIAL\RESP\"
    End If
End Sub
