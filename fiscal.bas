Attribute VB_Name = "Fiscal"
'Declaracoes para DarumaFramawork
Public Declare Function eDefinirProduto_Daruma Lib "DarumaFrameWork.dll" (ByVal sProduto As String) As Integer
Public Declare Function regAlterarValor_Daruma Lib "DarumaFrameWork.dll" (ByVal pszChave As String, ByVal pszValor As String) As Integer
Public Declare Function eBuscarPortaVelocidade_ECF_Daruma Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function rVerificarReducaoZ_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal zPendente As String) As Integer
'Abertura de cupom fiscal
Public Declare Function iCFAbrir_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal CPF As String, ByVal Nome As String, ByVal Endereco As String) As Integer


'Declaração da DLL com suas Funções
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal Indice As String) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_DataMovimentoUltimaReducaoMFD Lib "BEMAFI32.DLL" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencialMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BEMAFI32.DLL" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_GeraRegistrosCAT52MFD Lib "BEMAFI32.DLL" (ByVal cArquivo As String, ByVal cData As String) As Integer
Public Declare Function Bematech_FI_GeraRegistrosCAT52MFDEx Lib "BEMAFI32.DLL" (ByVal cArquivo As String, ByVal cData As String, ByVal cArqDestino As String) As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_MarcaModeloTipoImpressoraMFD Lib "BEMAFI32.DLL" (ByVal Marca As String, ByVal Modelo As String, ByVal Tipo As String) As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BEMAFI32.DLL" (ByVal pNumeroNaoFiscal As String) As Integer
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_NumeroSerieMFD Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BEMAFI32.DLL" (ByVal Descontos As String) As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BEMAFI32.DLL" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaIdAplicativoMFD Lib "BEMAFI32.DLL" (ByVal cAplicativo As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_Sangria Lib "BEMAFI32.DLL" (ByVal cValor As String) As Integer
Public Declare Function Bematech_FI_StatusEstendidoMFD Lib "BEMAFI32.DLL" (ByRef iStatus As Integer) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_UsaRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BEMAFI32.DLL" (ByVal ValorPago As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_VerificaZPendente Lib "BEMAFI32.DLL" (ByVal StatusRZ As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BEMAFI32.DLL" () As Integer

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

Public Declare Function Gera_AtoCotepe1704 Lib "DLLG2_Gerador.dll" (ByVal ComPortOrFileName As String, ByVal Modelo As String, ByVal RegFileName As String, ByVal DataReducao As String) As Long

'Metodos Cupom
Public Declare Function Daruma_FI_AbreCupom Lib "Daruma32.dll" (ByVal CPF_ou_CNPJ As String) As Integer
Public Declare Function Daruma_FI_VendeItem Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Daruma_FI_FechaCupomResumido Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal TextoLivre As String) As Integer
Public Declare Function Daruma_FI_IdentificaConsumidor Lib "Daruma32.dll" (ByVal NomeConsumidor As String, ByVal Endereco As String, ByVal CPF_ou_CNPJ As String) As Integer
Public Declare Function Daruma_FI_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_FechaCupom Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_CancelaItemAnterior Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_CancelaItemGenerico Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FI_CancelaCupom Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AumentaDescricaoItem Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
Public Declare Function Daruma_FI_UsaUnidadeMedida Lib "Daruma32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Daruma_FI_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_EstornoFormasPagamento Lib "Daruma32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer

'Metodos para Recebimentos e Relatorios
Public Declare Function Daruma_FI_LeituraX Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ReducaoZAjustaDataHora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer

'Metodos de Status
Public Declare Function Daruma_FI_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaTruncamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_VerificaZPendente Lib "Daruma32.dll" (ByVal zPendente As String) As Integer
Public Declare Function Daruma_FI_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StsCF As String) As Integer

'Metodos de Informacao do ECF e Contadores
Public Declare Function Daruma_FI_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Daruma_FI_NumeroCupom Lib "Daruma32.dll" (ByVal NumeroCupom As String) As Integer
Public Declare Function Daruma_FI_DataHoraReducao Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer

'Metodos Relatorios Fiscais e Relatorios
Public Declare Function Daruma_FI_AbreRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioGerencial Lib "Daruma32.dll" (ByVal Texto As String) As Integer
Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Daruma_FI_RetornaErroExtendido Lib "Daruma32.dll" (ByVal ErroExtendido As String) As Integer

'Metodos Prog e Config
Public Declare Function Daruma_FI_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal DescricaoFormasPgto As String) As Integer

'Metodos Totalizadores Gerais
Public Declare Function Daruma_FI_SubTotal Lib "Daruma32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Daruma_FI_SaldoAPagar Lib "Daruma32.dll" (ByVal Saldo As String) As Integer

'Metodos TEF
Public Declare Function Daruma_TEF_FechaRelatorio Lib "Daruma32.dll" () As Integer

Dim BemaRetorno As Integer

Dim lNomeArquivoTMP As String
Dim lNomeArquivoReq As String
Dim lNomeArquivoResp As String
Dim lRetorno As Integer
Global gParametroECF As String
Global gQuickCanal As Long
Global gPortaECF As String

Public Declare Function IniPortaStr Lib "MP20FI32.DLL" (ByVal Porta As String) As Integer
Public Declare Function FormataTX Lib "MP20FI32.DLL" (ByVal Buffer As String) As Integer
Public Declare Function FechaPorta Lib "MP20FI32.DLL" () As Integer
'Variável para Controlar retorno da DLL
Global RetornoCF As Integer
'Variável que irá conter o Comando a ser enviado para a Impressora
Global ComandoCF As String
Global gQtdViasTEF As Integer
Global gNumeroControleSolicitacao As Long
'*************************************************************
'Objetivo  -> Abrir a Porta de comunicação e Verificar o Retorno da DLL
'**************************************************************
Public Function AnalizaRetornoBematech(ByVal xBemaRetorno As Integer) As String
    If xBemaRetorno = 0 Then
        AnalizaRetornoBematech = "Erro de Comunicação!"
    ElseIf xBemaRetorno = -4 Then
        AnalizaRetornoBematech = "Arquivo ini não encontrado ou parâmetro inválido para o nome da porta!"
    ElseIf xBemaRetorno = -5 Then
        AnalizaRetornoBematech = "Erro ao abrir a porta de comunicação!"
    ElseIf xBemaRetorno = -6 Then
        AnalizaRetornoBematech = "A Impressora se encontra DESLIGADA!"
    ElseIf xBemaRetorno = -8 Then
        AnalizaRetornoBematech = "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt!"
    Else
        AnalizaRetornoBematech = "Erro não identificado!"
    End If
End Function
Public Function Abre_ProtocoloCF(x_porta As Integer)
    If x_porta = 1 Then
        RetornoCF = IniPortaStr("COM1")
    Else
        RetornoCF = IniPortaStr("COM2")
    End If
    If RetornoCF <> 1 Then
        MsgBox "Problemas ao Abrir a Porta de Comunicação"
        Fecha_ProtocoloCF
    End If
 End Function
'*************************************************************
'Objetivo  -> Enviar Comando para a Dll
'**************************************************************
Public Function Envia_ComandoCF()
    RetornoCF = FormataTX(ComandoCF)
    If RetornoCF <> 0 Then
        MsgBox "Problemas ao Enviar Comando"
        Fecha_ProtocoloCF
    End If
End Function
'*************************************************************
'Objetivo  -> Enviar Comando para a Dll
'**************************************************************
Public Function Fecha_ProtocoloCF()
    RetornoCF = FechaPorta()
    If RetornoCF <> 1 Then
        MsgBox "Problemas ao Fechar a Porta de Comunicação"
    End If
End Function
'*************************************************************
'Objetivo  -> Enviar Comando para a Dll
'**************************************************************
Public Function Testa_ImpressoraCF() As Boolean
    Testa_ImpressoraCF = False
    'BemaRetorno = Bematech_FI_AbrePortaSerial()
    BemaRetorno = Bematech_FI_VerificaImpressoraLigada()
    'BemaRetorno = Bematech_FI_FechaPortaSerial()
    If BemaRetorno = 1 Then
        Testa_ImpressoraCF = True
    Else
        Call AnalizaRetornoBematech(BemaRetorno)
    End If
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|19|" + Chr(27)
    'RetornoCF = FormataTX(ComandoCF)
    'If RetornoCF <> 0 Then
    '    Testa_ImpressoraCF = False
    '    MsgBox "Problemas ao Enviar Comando"
    'End If
    'Fecha_ProtocoloCF
End Function
Public Function PedidoCompartilhamentoECF(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As String
    
    On Error GoTo FileError
    
    PedidoCompartilhamentoECF = ""
    If pComando = "Abre Porta Serial" Then
        PedidoCompartilhamentoECF = EcfAbrePorta(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando)
    ElseIf pComando = "Verifica Impressora Ligada" Then
        PedidoCompartilhamentoECF = EcfVerificaImpressoraLigada(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando)
    ElseIf pComando = "Flags Fiscais" Then
        PedidoCompartilhamentoECF = EcfFlagsFiscais(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, Val(pParametro))
    ElseIf pComando = "Dados Ultima Reducao" Then
        PedidoCompartilhamentoECF = EcfDadosUltimaReducao(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando)
    ElseIf pComando = "Nomeia Departamento" Then
        PedidoCompartilhamentoECF = EcfNomeiaDepartamento(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Retorno Impressora" Then
        PedidoCompartilhamentoECF = EcfRetornoImpressora(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Inicia Fechamento Cupom" Then
        PedidoCompartilhamentoECF = EcfIniciaFechamentoCupom(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Numero do Cupom" Then
        PedidoCompartilhamentoECF = EcfNumeroCupom(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Data e Hora" Then
        PedidoCompartilhamentoECF = EcfDataHora(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Abre Cupom" Then
        PedidoCompartilhamentoECF = EcfAbreCupom(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Vende Item Departamento" Then
        PedidoCompartilhamentoECF = EcfVendeItemDepartamento(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
     ElseIf pComando = "Efetua Forma Pagamento Descricao Forma" Then
        PedidoCompartilhamentoECF = EcfEfetuaFormaPagamentoDescricaoForma(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Termina Fechamento Cupom" Then
        PedidoCompartilhamentoECF = EcfTerminaFechamentoCupom(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Abre Comprovante Nao Fiscal Vinculado" Then
        PedidoCompartilhamentoECF = EcfAbreComprovanteNaoFiscalVinculado(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Usa Comprovante Nao Fiscal Vinculado" Then
        PedidoCompartilhamentoECF = EcfUsaComprovanteNaoFiscalVinculado(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    ElseIf pComando = "Fecha Comprovante Nao Fiscal Vinculado" Then
        PedidoCompartilhamentoECF = EcfFechaComprovanteNaoFiscalVinculado(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro)
    Else
        MsgBox "O comando=" & pComando & ", não foi processado!", vbInformation, "Comando não interpretado!"
    End If
    Call CriaLogECF(Time & " - pComando:" & pComando & " - PedidoCompartilhamentoECF:" & PedidoCompartilhamentoECF & " - Rotina: PedidoCompartilhamentoECF")
    Call CriaLogECF(" ")
    Call DeletaArquivo(lNomeArquivoReq)
    'Call DeletaArquivo(lNomeArquivoResp)
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro:" & Error & " - Rotina: PedidoCompartilhamentoECF")
End Function
Public Function EcfAbrePorta(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String) As Integer
    
    On Error GoTo FileError
    
    EcfAbrePorta = 0
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, "") Then
    End If
    If LeArquivoRetornoEcf Then
        EcfAbrePorta = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfAbrePorta")
End Function
Public Function EcfVerificaImpressoraLigada(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String) As Integer
    
    On Error GoTo FileError
    
    EcfVerificaImpressoraLigada = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, "") Then
    End If
    If LeArquivoRetornoEcf Then
        EcfVerificaImpressoraLigada = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfVerificaImpressoraLigada")
End Function

Public Function EcfFlagsFiscais(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As Integer) As Integer
    
    On Error GoTo FileError
    
    EcfFlagsFiscais = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, str(pParametro)) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfFlagsFiscais = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfFlagsFiscais")
End Function

Public Function EcfDadosUltimaReducao(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String) As Integer
    
    On Error GoTo FileError
    
    EcfDadosUltimaReducao = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, "") Then
    End If
    If LeArquivoRetornoEcf Then
        EcfDadosUltimaReducao = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfDadosUltimaReducao")
End Function

Public Function EcfNomeiaDepartamento(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfNomeiaDepartamento = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfNomeiaDepartamento = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfNomeiaDepartamento")
End Function

Public Function EcfRetornoImpressora(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfRetornoImpressora = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfRetornoImpressora = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfRetornoImpressora")
End Function

Public Function EcfIniciaFechamentoCupom(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfIniciaFechamentoCupom = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfIniciaFechamentoCupom = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfIniciaFechamentoCupom")
End Function

Public Function EcfNumeroCupom(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfNumeroCupom = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfNumeroCupom = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfNumeroCupom")
End Function
Public Function EcfBematechReducaoZPendente() As Boolean
    Dim xStatusRZ As String
    'Dim xCodigoErro As Long
    'Dim xStrCodigoErro As String
    'Dim xFaseErro As Integer
    
    On Error GoTo trata_erro
    
    EcfBematechReducaoZPendente = False
    xStatusRZ = Space(1)
    Call CriaLogCupom("Bematech_FI_VerificaZPendente(xStatusRZ)")
    BemaRetorno = Bematech_FI_VerificaZPendente(xStatusRZ)
    Call CriaLogCupom("Bematech_FI_VerificaZPendente(xStatusRZ) - xStatusRZ=" & xStatusRZ & " - BemaRetorno=" & BemaRetorno)
    If BemaRetorno = 1 Then
        If xStatusRZ = "1" Then
            EcfBematechReducaoZPendente = True
        End If
    'Else
    '    Call AnalizaRetornoBematech(BemaRetorno)
    End If
    
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro: EcfBematechReducaoZPendente Erro=" & Err.Number & Err.Description)
End Function

Public Function EcfDataHora(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfDataHora = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfDataHora = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfDataHora")
End Function

Public Function EcfAbreCupom(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfAbreCupom = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfAbreCupom = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfAbreCupom")
End Function

Public Function EcfVendeItemDepartamento(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfVendeItemDepartamento = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfVendeItemDepartamento = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfVendeItemDepartamento")
End Function

Public Function EcfEfetuaFormaPagamentoDescricaoForma(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfEfetuaFormaPagamentoDescricaoForma = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfEfetuaFormaPagamentoDescricaoForma = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfEfetuaFormaPagamentoDescricaoForma")
End Function

Public Function EcfTerminaFechamentoCupom(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfTerminaFechamentoCupom = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfTerminaFechamentoCupom = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfTerminaFechamentoCupom")
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAbreCreditoDebito Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAbreCupomFiscal Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAbreGaveta Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAbreGerencial Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAcresceItemFiscal Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAcresceSubTotal Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickCancelaCupom Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickCancelaItemFiscal(ByVal pItemCupom As Integer) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickCancelaItemFiscal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pItemCupom > 0 Then
            If EcfQuickAdicionaParametro("NumItem", CStr(pItemCupom), 4) Then
            End If
        End If
        If EcfQuickExecutaComando("CancelaItemFiscal") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickCancelaItemFiscal = True
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickCancelaItemFiscal Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickConverteCodigoAliquota(ByVal pAliquotaCodigoFiscal As String) As Integer
    EcfQuickConverteCodigoAliquota = -2
    If pAliquotaCodigoFiscal = "FF" Then
        EcfQuickConverteCodigoAliquota = -2
    ElseIf pAliquotaCodigoFiscal = "II" Then
        EcfQuickConverteCodigoAliquota = -3
    ElseIf pAliquotaCodigoFiscal = "NN" Then
        EcfQuickConverteCodigoAliquota = -4
    ElseIf Len(pAliquotaCodigoFiscal) = 2 Then
        If IsNumeric(Mid(pAliquotaCodigoFiscal, 1, 1)) And IsNumeric(Mid(pAliquotaCodigoFiscal, 2, 1)) Then
            EcfQuickConverteCodigoAliquota = (Val(pAliquotaCodigoFiscal) - 1)
        End If
    End If
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickLeituraX Erro=" & Err.Number & " - " & Err.Description)
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
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickPagaCupom Erro=" & Err.Number & " - " & Err.Description)
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
        If Not EcfQuickAdicionaParametro("Hora", fMascaraHora(str(Time)), 3) Then
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
                    'Aguarda 20 segundos
                    Sleep (20000)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickReducaoZ Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickReducaoZPendente() As Boolean
    Dim xCodigoErro As Long
    Dim xStrCodigoErro As String
    Dim xFaseErro As Integer
    
    On Error GoTo trata_erro
    
    EcfQuickReducaoZPendente = False
    xFaseErro = 1
    xStrCodigoErro = EcfQuickLeRegistrador("Indicadores", "Inteiro", 4)
    If xStrCodigoErro = "" Then
        xStrCodigoErro = "0"
    Else
        xFaseErro = 2
        xCodigoErro = CLng(xStrCodigoErro)
    End If
    xFaseErro = 3
    If xCodigoErro >= 16384 Then
        xFaseErro = 4
        xCodigoErro = xCodigoErro - 16384
    End If
    xFaseErro = 5
    If xCodigoErro >= 8192 Then
        xFaseErro = 6
        xCodigoErro = xCodigoErro - 8192
    End If
    xFaseErro = 7
    If xCodigoErro >= 4096 Then
        xFaseErro = 8
        xCodigoErro = xCodigoErro - 4096
    End If
    xFaseErro = 9
    If xCodigoErro >= 2048 Then
        xFaseErro = 10
        xCodigoErro = xCodigoErro - 2048
    End If
    xFaseErro = 11
    If xCodigoErro >= 1024 Then
        xFaseErro = 12
        xCodigoErro = xCodigoErro - 1024
    End If
    xFaseErro = 13
    If xCodigoErro >= 512 Then
        xFaseErro = 14
        xCodigoErro = xCodigoErro - 512
    End If
    xFaseErro = 15
    If xCodigoErro >= 256 Then
        xFaseErro = 16
        xCodigoErro = xCodigoErro - 256
    End If
    xFaseErro = 17
    If xCodigoErro >= 128 Then
        xFaseErro = 18
        xCodigoErro = xCodigoErro - 128
        EcfQuickReducaoZPendente = True
    End If
    xFaseErro = 19
    If xCodigoErro >= 64 Then
        xFaseErro = 20
        xCodigoErro = xCodigoErro - 64
    End If
    xFaseErro = 21
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro: EcfQuickReducaoZPendente Erro=" & Err.Number & " Fase:" & xFaseErro & " - " & Err.Description)
End Function
Public Function EcfQuickSemPapel() As Boolean
    Dim xCodigoErro As Long
    Dim xStrCodigoErro As String
    Dim xFaseErro As Integer
    
    On Error GoTo trata_erro
    
    xFaseErro = 1
    EcfQuickSemPapel = False
    xStrCodigoErro = EcfQuickLeRegistrador("Indicadores", "Inteiro", 4)
    If xStrCodigoErro = "" Then
        xStrCodigoErro = "0"
    Else
        xFaseErro = 2
        xCodigoErro = CLng(xStrCodigoErro)
    End If
    xFaseErro = 3
    If xCodigoErro >= 16384 Then
        xFaseErro = 4
        xCodigoErro = xCodigoErro - 16384
    End If
    xFaseErro = 5
    If xCodigoErro >= 8192 Then
        xFaseErro = 6
        xCodigoErro = xCodigoErro - 8192
    End If
    xFaseErro = 7
    If xCodigoErro >= 4096 Then
        xFaseErro = 8
        xCodigoErro = xCodigoErro - 4096
    End If
    xFaseErro = 9
    If xCodigoErro >= 2048 Then
        xFaseErro = 10
        xCodigoErro = xCodigoErro - 2048
    End If
    xFaseErro = 11
    If xCodigoErro >= 1024 Then
        xFaseErro = 12
        xCodigoErro = xCodigoErro - 1024
    End If
    xFaseErro = 13
    If xCodigoErro >= 512 Then
        xFaseErro = 14
        xCodigoErro = xCodigoErro - 512
    End If
    xFaseErro = 15
    If xCodigoErro >= 256 Then
        xFaseErro = 16
        xCodigoErro = xCodigoErro - 256
        xFaseErro = 17
        EcfQuickSemPapel = True
    End If
    xFaseErro = 18
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro: EcfQuickSemPapel Erro=" & Err.Number & " Fase:" & xFaseErro & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickSetaArquivoLog Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickObtemNomeLog Erro=" & Err.Number & " - " & Err.Description)
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
                    MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickBuscaData Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickAcertaHorarioVerao() As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAcertaHorarioVerao = False
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
                    'MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickAcertaHorarioVerao Erro=" & Err.Number & " - " & Err.Description)
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
                    MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickBuscaHora Erro=" & Err.Number & " - " & Err.Description)
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
        If Not EcfQuickAdicionaParametro("AliquotaICMS", str(xAliquotaICMS), 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodAliquota", str(pCodigoAliquota), 4) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodDepartamento", str(pCodigoDepartamento), 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("CodProduto", str(pCodigoProduto), 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pNomeDepartamento <> "" Then
            If Not EcfQuickAdicionaParametro("NomeDepartamento", str(pNomeDepartamento), 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        If Not EcfQuickAdicionaParametro("NomeProduto", pNomeProduto, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pPercentualAliquota > 0 Then
            If Not EcfQuickAdicionaParametro("PercentualAliquota", str(pPercentualAliquota), 6) Then
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickVendeItem Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickLeMeioPagamento Erro=" & Err.Number & " - " & Err.Description)
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
                                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickLeRegistrador Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickDataReducaoZ Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickDefineGerencial Erro=" & Err.Number & " - " & Err.Description)
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
        If Not EcfQuickAdicionaParametro("PermiteVinculado", str(xPermite), 0) Then
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickDefineMeioPagamento Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickAdicionaParametro(ByVal pNome As String, ByVal pValor As String, ByVal pTipo As Long) As Boolean
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickAdicionaParametro = False
    DLLG2_AdicionaParam gQuickCanal, pNome, pValor, pTipo
    EcfQuickAdicionaParametro = True
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro: EcfQuickAdicionaParametro Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickEmiteLeituraMF(ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoDocumento As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickEmiteLeituraMF = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("DataInicial", pDataInicial, 2) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("DataFinal", pDataFinal, 2) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("S", pDestino, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If pTipoDocumento <> "" Then
            If Not EcfQuickAdicionaParametro("TipoDocumento", pTipoDocumento, 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        If EcfQuickExecutaComando("EmiteLeituraFitaDetalhe") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 And xCodigoErro <> 11010 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickEmiteLeituraMF = True
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickEmiteLeituraMF Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickEmiteMemoriaFiscal(ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pDestino As String, ByVal pLeituraSimplificada As Boolean) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickEmiteMemoriaFiscal = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("DataInicial", pDataInicial, 2) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("DataFinal", pDataFinal, 2) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("Destino", pDestino, 7) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If Not EcfQuickAdicionaParametro("LeituraSimplificada", pLeituraSimplificada, 0) Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        If EcfQuickExecutaComando("EmiteLeituraMF") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 And xCodigoErro <> 11010 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickEmiteMemoriaFiscal = True
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickEmiteMemoriaFiscal Erro=" & Err.Number & " - " & Err.Description)
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
        If pTextoPromocional <> "" Then
            pTextoPromocional = Mid(pTextoPromocional, 1, 492)
            If Not EcfQuickAdicionaParametro("TextoPromocional", pTextoPromocional, 7) Then
                EcfQuickEncerraDriver
                Exit Function
            End If
        End If
        If EcfQuickExecutaComando("EncerraDocumento") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickEncerraDocumento = True
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickEncerraDocumento Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickEncerraDriver Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickExecutaComando Erro=" & Err.Number & " - " & Err.Description)
End Function
Public Function EcfQuickImprimeTexto(ByVal pTextoLivre As String) As Boolean
    Dim xCodigoErro As Long
    
    On Error GoTo trata_erro
    
    EcfQuickImprimeTexto = False
    gQuickCanal = -1
    If EcfQuickIniciaDriver Then
        If Not EcfQuickLimpaParametro Then
            EcfQuickEncerraDriver
            Exit Function
        End If
        pTextoLivre = Mid(pTextoLivre, 1, 492)
        If EcfQuickAdicionaParametro("TextoLivre", pTextoLivre, 7) Then
        End If
        If EcfQuickExecutaComando("ImprimeTexto") Then
            xCodigoErro = EcfQuickObtemCodigoErro
            If xCodigoErro > 0 Then
                MsgBox "Erro de retorno ao executar comando na ECF Quick." & vbCrLf & "Erro n.:" & xCodigoErro, vbCritical, "Erro de Retorno!"
            Else
                EcfQuickImprimeTexto = True
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickImprimeTexto Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickIniciaDriver() As Boolean
    On Error GoTo trata_erro
    
    EcfQuickIniciaDriver = False
    If Len(gPortaECF) = 0 Then
        gPortaECF = ReadINI("CUPOM FISCAL", "Porta ECF", gArquivoIni)
        If Len(gPortaECF) = 0 Then
            gPortaECF = "COM2"
            MsgBox "Falta defenir no sgp.ini (Porta ECF=COM?)" & vbCrLf & "O sistema irá definir COM2", vbCritical, "Configuração Ausente no SPG.INI"
        End If
    End If
    gQuickCanal = DLLG2_IniciaDriver(gPortaECF)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickIniciaDriver Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickLimpaParametro Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickListaParametro Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Function EcfQuickObtemCodigoErro() As Long
    Dim xRetorno As Long
    
    On Error GoTo trata_erro
    
    EcfQuickObtemCodigoErro = 999999
    xRetorno = DLLG2_ObtemCodErro(gQuickCanal)
    EcfQuickObtemCodigoErro = xRetorno
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro: EcfQuickObtemCodigoErro Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickObtemRetornos Erro=" & Err.Number & " - " & Err.Description)
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
    Call CriaLogCupom(Time & " - Erro: EcfQuickVersao Erro=" & Err.Number & " - " & Err.Description)
End Function


Public Function EcfAbreComprovanteNaoFiscalVinculado(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfAbreComprovanteNaoFiscalVinculado = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfAbreComprovanteNaoFiscalVinculado = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfAbreComprovanteNaoFiscalVinculado")
End Function

Public Function EcfUsaComprovanteNaoFiscalVinculado(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfUsaComprovanteNaoFiscalVinculado = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfUsaComprovanteNaoFiscalVinculado = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfUsaComprovanteNaoFiscalVinculado")
End Function

Public Function EcfFechaComprovanteNaoFiscalVinculado(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Integer
    
    On Error GoTo FileError
    
    EcfFechaComprovanteNaoFiscalVinculado = 0
    
    If CriaArquivoPedidoEcf(pUnidadeEcfInstalada, pOrigem, pNomeECF, pComando, pParametro) Then
    End If
    If LeArquivoRetornoEcf Then
        EcfFechaComprovanteNaoFiscalVinculado = lRetorno
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro de integridade!" & " - Rotina: EcfFechaComprovanteNaoFiscalVinculado")
End Function

Public Function CriaArquivoPedidoEcf(ByVal pUnidadeEcfInstalada As String, ByVal pOrigem As String, ByVal pNomeECF As String, ByVal pComando As String, ByVal pParametro As String) As Boolean
    
    On Error GoTo FileError
    
    CriaArquivoPedidoEcf = False
    lNomeArquivoTMP = pUnidadeEcfInstalada & ":\Pedido" & Format(Time, "hhmmss") & ".TMP"
    lNomeArquivoReq = Mid(lNomeArquivoTMP, 1, 16) & "ECF"
    lNomeArquivoResp = "C:\Retorno" & Mid(lNomeArquivoReq, 10, 6) & ".ECF"
    Set gArquivoTMP = gArqTxt.CreateTextFile(lNomeArquivoTMP)
    gArquivoTMP.WriteLine ("[PEDIDO ECF]")
    gArquivoTMP.WriteLine ("Nome ECF=" & pNomeECF)
    gArquivoTMP.WriteLine ("Origem=" & pOrigem)
    gArquivoTMP.WriteLine ("Comando=" & pComando)
    If pParametro <> "" Then
        gArquivoTMP.WriteLine ("Parametro=" & pParametro)
    End If
    gArquivoTMP.Close
    If gArqTxt.FileExists(lNomeArquivoTMP) Then
        gArqTxt.MoveFile (lNomeArquivoTMP), (lNomeArquivoReq)
        CriaArquivoPedidoEcf = True
    End If
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro ao gerar arquivo " & lNomeArquivoTMP & " - Rotina: CriaArquivoPedidoEcf")
End Function

Sub CriaLogECF(ByVal xTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "ECF_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "ECF_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (xTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    MsgBox Error
    MsgBox "Erro ao criar LOG ECF: " & xTipoLog, vbInformation, "Erro: CriaLogCupom"
    Exit Sub
End Sub
Public Function LeArquivoRetornoEcf() As Boolean
    Dim xHoraInicial As Date
    Dim xLeRetorno As Boolean
    Dim xRetorno As String
    Dim xComando As String
    
    On Error GoTo FileError
    
    lRetorno = 0
    LeArquivoRetornoEcf = False
    xLeRetorno = False
    
    'Aguarda 10 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 10
        If gArqTxt.FileExists(lNomeArquivoResp) Then
            xLeRetorno = True
            Exit Do
        End If
    Loop
    
    'gArquivo.WriteLine ("[RETORNO ECF]")
    'gArquivo.WriteLine ("Nome ECF=" & lNomeECF)
    'gArquivo.WriteLine ("Origem=" & lComputadorOrigem)
    'gArquivo.WriteLine ("Comando=" & lComandoECF)
    'gArquivo.WriteLine ("Retorno=" & pRetorno)
    'gArquivo.WriteLine ("Parametro=" & gParametroECF)
    If xLeRetorno = True Then
        xRetorno = ReadINI("RETORNO ECF", "Retorno", lNomeArquivoResp)
        xComando = ReadINI("RETORNO ECF", "Comando", lNomeArquivoResp)
        gParametroECF = ReadINI("RETORNO ECF", "Parametro", lNomeArquivoResp)
        Call CriaLogECF(Time & " - Arquivo " & lNomeArquivoResp & " - Comando:" & xComando & ", Retorno:" & xRetorno & " - Rotina: LeArquivoRetornoEcf")
        If xComando = "Flags Fiscais" Then
            LeArquivoRetornoEcf = True
            lRetorno = Val(xRetorno)
        Else
            If xRetorno = "1" Then
                LeArquivoRetornoEcf = True
                lRetorno = Val(xRetorno)
            End If
        End If
        DeletaArquivo (lNomeArquivoResp)
    Else
        Call CriaLogECF(Time & " - Não foi possível encontrar o arquivo " & lNomeArquivoResp & " - Rotina: LeArquivoRetornoEcf")
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro ao analizar conteúdo do arquivo " & lNomeArquivoResp & " - Rotina: LeArquivoRetornoEcf")
End Function
Private Function DeletaArquivo(ByVal pNomeArquivo As String) As Boolean
    
    On Error GoTo FileError
    
    DeletaArquivo = False
    
    If gArqTxt.FileExists(pNomeArquivo) Then
        Call gArqTxt.DeleteFile(pNomeArquivo, True)
        DeletaArquivo = True
        Call CriaLogECF(Time & " - O arquivo " & pNomeArquivo & " foi deletado com sucesso!" & " - Rotina: DeletaArquivo")
    Else
        Call CriaLogECF(Time & " - O arquivo " & pNomeArquivo & " não existe!" & " - Rotina: DeletaArquivo")
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Time & " - Erro ao excluir o arquivo " & pNomeArquivo & " - Error:" & Error & " - Rotina: DeletaArquivo")
End Function

