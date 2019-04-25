VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovDocEletronicoCabecalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdEstabelecimento As Integer
Private mDataEmissao As Date
Private mEntrada As Boolean
Private mSaida As Boolean
Private mModelo As String
Private mSerie As String
Private mNumero As String
Private mHoraSaida As Date
Private mDataEntradaSaida As Date
Private mEmissaoPropria As Boolean
Private mIdClienteFornecedor As Long
Private mCodigoSituacao As Integer
Private mChaveAcesso As String
Private mFormaPagamento As String
Private mValorTotal As Currency
Private mValorDesconto As Currency
Private mValorAbatimentoNaoTributado As Currency
Private mValorProdutos As Currency
Private mTipoFrete As String
Private mValorFrete As Currency
Private mValorSeguro As Currency
Private mOutrasDespesas As Currency
Private mValorBCICMS As Currency
Private mAliquotaICMS As Currency
Private mValorICMS As Currency
Private mValorBCICMSST As Currency
Private mAliquotaIcmsSt As Currency
Private mValorICMSST As Currency
Private mValorIPI As Currency
Private mValorPis As Currency
Private mValorCofins As Currency
Private mValorPisSt As Currency
Private mValorCofinsSt As Currency
Private mCombustivel As Boolean
Private mCancelado As Boolean
Private mAguaEnergiaGasTelefone As String
Private mInutilizada As Boolean
Private mIncidePisConfis As Boolean
Private mDataDigitacao As Date
Private mDataAlteracao As Date
Private mIdUsuario As Integer
Private mNumeroLote As Long
Private mNumeroRecepcao As Long
Private mNumeroProtocolo As Long
Private mEtapaConcluida As Integer
Private mNomeEtapaConcluida As String
Private mCodigoRetorno As String
Private mMensagemRetorno As String
Private mCodigoUltimoEvento As Integer
Private mObservacaoEvento As String
Private mProgramaOrigem As String

Private mPeriodo As String
Private mTipoMovimento As Integer
Private mTipoSubEstoque As Integer


Private rsMovDocumentoEletronicoCabecalho As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get IdEstabelecimento() As Integer
    IdEstabelecimento = mIdEstabelecimento
End Property
Public Property Let IdEstabelecimento(ByVal Valor As Integer)
    mIdEstabelecimento = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get Entrada() As Boolean
    Entrada = mEntrada
End Property
Public Property Let Entrada(ByVal Valor As Boolean)
    mEntrada = Valor
End Property
Public Property Get Saida() As Boolean
    Saida = mSaida
End Property
Public Property Let Saida(ByVal Valor As Boolean)
    mSaida = Valor
End Property

Public Property Get Modelo() As String
    Modelo = mModelo
End Property
Public Property Let Modelo(ByVal Valor As String)
    mModelo = Valor
End Property

Public Property Get Serie() As String
    Serie = mSerie
End Property
Public Property Let Serie(ByVal Valor As String)
    mSerie = Valor
End Property

Public Property Get Numero() As String
    mNumero = mNumero
End Property
Public Property Let Numero(ByVal Valor As String)
    mNumero = Valor
End Property
Public Property Get HoraSaida() As Date
    HoraSaida = mHoraSaida
End Property
Public Property Let HoraSaida(ByVal Valor As Date)
    mHoraSaida = Valor
End Property
Public Property Get DataEntradaSaida() As Date
    DataEntradaSaida = mDataEntradaSaida
End Property
Public Property Let DataEntradaSaida(ByVal Valor As Date)
    mDataEntradaSaida = Valor
End Property

Public Property Get EmissaoPropria() As Boolean
    EmissaoPropria = mEmissaoPropria
End Property
Public Property Let EmissaoPropria(ByVal Valor As Boolean)
    mEmissaoPropria = Valor
End Property

Public Property Get IdClienteFornecedor() As Long
    IdClienteFornecedor = mIdClienteFornecedor
End Property
Public Property Let IdClienteFornecedor(ByVal Valor As Long)
    mIdClienteFornecedor = Valor
End Property
Public Property Get CodigoSituacao() As Integer
    CodigoSituacao = mCodigoSituacao
End Property
Public Property Let CodigoSituacao(ByVal Valor As Integer)
    mCodigoSituacao = Valor
End Property
Public Property Get ChaveAcesso() As String
    ChaveAcesso = mChaveAcesso
End Property
Public Property Let ChaveAcesso(ByVal Valor As String)
    mChaveAcesso = Valor
End Property
Public Property Get FormaPagamento() As String
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As String)
    mFormaPagamento = Valor
End Property

Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property
Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property

Public Property Get ValorAbatimentoNaoTributado() As Currency
    ValorAbatimentoNaoTributado = mValorAbatimentoNaoTributado
End Property
Public Property Let ValorAbatimentoNaoTributado(ByVal Valor As Currency)
    mValorAbatimentoNaoTributado = Valor
End Property
Public Property Get ValorProdutos() As Currency
    ValorProdutos = mValorProdutos
End Property
Public Property Let ValorProdutos(ByVal Valor As Currency)
    mValorProdutos = Valor
End Property
Public Property Get TipoFrete() As String
    TipoFrete = mTipoFrete
End Property
Public Property Let TipoFrete(ByVal Valor As String)
    mTipoFrete = Valor
End Property
Public Property Get ValorFrete() As Currency
    ValorFrete = mValorFrete
End Property
Public Property Let ValorFrete(ByVal Valor As Currency)
    mValorFrete = Valor
End Property
Public Property Get ValorSeguro() As Currency
    ValorSeguro = mValorSeguro
End Property
Public Property Let ValorSeguro(ByVal Valor As Currency)
    mValorSeguro = Valor
End Property
Public Property Get OutrasDespesas() As Currency
    OutrasDespesas = mOutrasDespesas
End Property
Public Property Let OutrasDespesas(ByVal Valor As Currency)
    mOutrasDespesas = Valor
End Property

Public Property Get ValorBCICMS() As Currency
    ValorBCICMS = mValorBCICMS
End Property
Public Property Let ValorBCICMS(ByVal Valor As Currency)
    mValorBCICMS = Valor
End Property
Public Property Get AliquotaICMS() As Currency
    AliquotaICMS = mAliquotaICMS
End Property
Public Property Let AliquotaICMS(ByVal Valor As Currency)
    mAliquotaICMS = Valor
End Property
Public Property Get ValorICMS() As Currency
    ValorICMS = mValorICMS
End Property
Public Property Let ValorICMS(ByVal Valor As Currency)
    mValorICMS = Valor
End Property
Public Property Get ValorBCICMSST() As Currency
    ValorBCICMSST = mValorBCICMSST
End Property
Public Property Let ValorBCICMSST(ByVal Valor As Currency)
    mValorBCICMSST = Valor
End Property
Public Property Get AliquotaIcmsSt() As Currency
    AliquotaIcmsSt = mAliquotaIcmsSt
End Property
Public Property Let AliquotaIcmsSt(ByVal Valor As Currency)
    mAliquotaIcmsSt = Valor
End Property
Public Property Get ValorICMSST() As Currency
    ValorICMSST = mValorICMSST
End Property
Public Property Let ValorICMSST(ByVal Valor As Currency)
    mValorICMSST = Valor
End Property
Public Property Get ValorIPI() As Currency
    ValorIPI = mValorIPI
End Property
Public Property Let ValorIPI(ByVal Valor As Currency)
    mValorIPI = Valor
End Property
Public Property Get ValorPis() As Currency
    ValorPis = mValorPis
End Property
Public Property Let ValorPis(ByVal Valor As Currency)
    mValorPis = Valor
End Property
Public Property Get ValorCofins() As Currency
    ValorCofins = mValorCofins
End Property
Public Property Let ValorCofins(ByVal Valor As Currency)
    mValorCofins = Valor
End Property
Public Property Get ValorPisSt() As Currency
    ValorPisSt = mValorPisSt
End Property
Public Property Let ValorPisSt(ByVal Valor As Currency)
    mValorPisSt = Valor
End Property
Public Property Get ValorCofinsSt() As Currency
    ValorCofinsSt = mValorCofinsSt
End Property
Public Property Let ValorCofinsSt(ByVal Valor As Currency)
    mValorCofinsSt = Valor
End Property


Public Property Get Combustivel() As Boolean
    Combustivel = mCombustivel
End Property
Public Property Let Combustivel(ByVal Valor As Boolean)
    mCombustivel = Valor
End Property
Public Property Get Cancelado() As Boolean
    Cancelado = mCancelado
End Property
Public Property Let Cancelado(ByVal Valor As Boolean)
    mCancelado = Valor
End Property
Public Property Get AguaEnergiaGasTelefone() As String
    AguaEnergiaGasTelefone = mAguaEnergiaGasTelefone
End Property
Public Property Let AguaEnergiaGasTelefone(ByVal Valor As String)
    mAguaEnergiaGasTelefone = Valor
End Property

Public Property Get Inutilizada() As Boolean
    Inutilizada = mInutilizada
End Property
Public Property Let Inutilizada(ByVal Valor As Boolean)
    mInutilizada = Valor
End Property
Public Property Get IncidePisConfis() As Boolean
    IncidePisConfis = mIncidePisConfis
End Property
Public Property Let IncidePisConfis(ByVal Valor As Boolean)
    mIncidePisConfis = Valor
End Property
Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property
Public Property Get DataAlteracao() As Date
    DataAlteracao = mDataAlteracao
End Property
Public Property Let DataAlteracao(ByVal Valor As Date)
    mDataAlteracao = Valor
End Property

Public Property Get IdUsuario() As Integer
    IdUsuario = mIdUsuario
End Property
Public Property Let IdUsuario(ByVal Valor As Integer)
    mIdUsuario = Valor
End Property
Public Property Get NumeroLote() As Long
    NumeroLote = mNumeroLote
End Property
Public Property Let NumeroLote(ByVal Valor As Long)
    mNumeroLote = Valor
End Property
Public Property Get NumeroRecepcao() As Long
    NumeroRecepcao = mNumeroRecepcao
End Property
Public Property Let NumeroRecepcao(ByVal Valor As Long)
    mNumeroRecepcao = Valor
End Property
Public Property Get NumeroProtocolo() As Long
    NumeroProtocolo = mNumeroProtocolo
End Property
Public Property Let NumeroProtocolo(ByVal Valor As Long)
    mNumeroProtocolo = Valor
End Property
Public Property Get EtapaConcluida() As Integer
    EtapaConcluida = mEtapaConcluida
End Property
Public Property Let EtapaConcluida(ByVal Valor As Integer)
    mEtapaConcluida = Valor
End Property
Public Property Get NomeEtapaConcluida() As String
    NomeEtapaConcluida = mNomeEtapaConcluida
End Property
Public Property Let NomeEtapaConcluida(ByVal Valor As String)
    mNomeEtapaConcluida = Valor
End Property
Public Property Get CodigoRetorno() As String
    CodigoRetorno = mCodigoRetorno
End Property
Public Property Let CodigoRetorno(ByVal Valor As String)
    mCodigoRetorno = Valor
End Property
Public Property Get MensagemRetorno() As String
    MensagemRetorno = mMensagemRetorno
End Property
Public Property Let MensagemRetorno(ByVal Valor As String)
    mMensagemRetorno = Valor
End Property
Public Property Get CodigoUltimoEvento() As Integer
    CodigoUltimoEvento = mCodigoUltimoEvento
End Property
Public Property Let CodigoUltimoEvento(ByVal Valor As Integer)
    mCodigoUltimoEvento = Valor
End Property
Public Property Get ObservacaoEvento() As String
    ObservacaoEvento = mObservacaoEvento
End Property
Public Property Let ObservacaoEvento(ByVal Valor As String)
    mObservacaoEvento = Valor
End Property
Public Property Get ProgramaOrigem() As String
    ProgramaOrigem = mProgramaOrigem
End Property
Public Property Let ProgramaOrigem(ByVal Valor As String)
    mProgramaOrigem = Valor
End Property


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property
Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property
Public Property Get TipoSubEstoque() As Integer
    TipoSubEstoque = mTipoSubEstoque
End Property
Public Property Let TipoSubEstoque(ByVal Valor As Integer)
    mTipoSubEstoque = Valor
End Property



Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False


    gSQL = "INSERT INTO MovimentoDocumentoEletronicoCabecalho (IdEstabelecimento_MovDECabecalho,DataEmissao_MovDECabecalho,Entrada_MovDECabecalho"
    gSQL = gSQL & ",Saida_MovDECabecalho,Modelo_MovDECabecalho,Serie_MovDECabecalho,Numero_MovDECabecalho"
    gSQL = gSQL & ",HoraSaida_MovDECabecalho,DataEntradaSaida_MovDECabecalho,EmissaoPropria_MovDECabecalho"
    gSQL = gSQL & ",IdClienteFornecedor_MovDECabecalho,CodigoSituacao_MovDECabecalho,ChaveAcesso_MovDECabecalho"
    gSQL = gSQL & ",FormaPagamento_MovDECabecalho,ValorTotal_MovDECabecalho,ValorDesconto_MovDECabecalho"
    gSQL = gSQL & ",ValorAbatimentoNaoTributado_MovDECabecalho,ValorProdutos_MovDECabecalho,TipoFrete_MovDECabecalho"
    gSQL = gSQL & ",ValorFrete_MovDECabecalho,ValorSeguro_MovDECabecalho,OutrasDespesas_MovDECabecalho,ValorBcIcms_MovDECabecalho"
    gSQL = gSQL & ",AliquotaIcms_MovDECabecalho,ValorIcms_MovDECabecalho,ValorBcIcmsSt_MovDECabecalho,AliquotaIcmsSt_MovDECabecalho"
    gSQL = gSQL & ",ValorIcmsSt_MovDECabecalho,ValorIpi_MovDECabecalho,ValorPis_MovDECabecalho,ValorCofins_MovDECabecalho"
    gSQL = gSQL & ",ValorPisSt_MovDECabecalho,ValorCofinsSt_MovDECabecalho,Combustivel_MovDECabecalho,Cancelado_MovDECabecalho"
    gSQL = gSQL & ",AguaEnergiaGasTelefone_MovDECabecalho,Inutilizada_MovDECabecalho,IncidePisCofins_MovDECabecalho,DataDigitacao_MovDECabecalho"
    gSQL = gSQL & ",DataAlteracao_MovDECabecalho,IdUsuario_MovDECabecalho,NumeroLote_MovDECabecalho,NumeroRecepcao_MovDECabecalho"
    gSQL = gSQL & ",NumeroProtocolo_MovDECabecalho,EtapaConcluida_MovDECabecalho"
    gSQL = gSQL & ",NomeEtapaConcluida_MovDECabecalho, CodigoRetorno_MovDECabecalho, MensagemRetorno_MovDECabecalho, CodigoUltimoEvento_MovDECabecalho, ObservacaoEvento_MovDECabecalho, ProgramaOrigem_MovDECabecalho,Periodo_MovDECabecalho, TipoMovimento_MovDECabecalho, TipoSubEstoque_MovDECabecalho) VALUES ("
    
    Call sqlNumero(1, mIdEstabelecimento, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlBoolean(1, mEntrada, ", ")
    Call sqlBoolean(1, mSaida, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlTexto(1, mNumero, ", ")
    Call sqlHora(1, mHoraSaida, ", ")
    Call sqlData(1, mDataEntradaSaida, ", ")
    Call sqlBoolean(1, mEmissaoPropria, ", ")
    Call sqlNumero(1, mIdClienteFornecedor, ", ")
    Call sqlNumero(1, mCodigoSituacao, ", ")
    Call sqlTexto(1, mChaveAcesso, ", ")
    Call sqlTexto(1, mFormaPagamento, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlValor(1, mValorAbatimentoNaoTributado, ", ")
    Call sqlValor(1, mValorProdutos, ", ")
    Call sqlTexto(1, mTipoFrete, ", ")
    Call sqlValor(1, mValorFrete, ", ")
    Call sqlValor(1, mValorSeguro, ", ")
    Call sqlValor(1, mOutrasDespesas, ", ")
    Call sqlValor(1, mValorBCICMS, ", ")
    Call sqlValor(1, mAliquotaICMS, ", ")
    Call sqlValor(1, mValorICMS, ", ")
    Call sqlValor(1, mValorBCICMSST, ", ")
    Call sqlValor(1, mAliquotaIcmsSt, ", ")
    Call sqlValor(1, mValorICMSST, ", ")
    Call sqlValor(1, mValorIPI, ", ")
    Call sqlValor(1, mValorPis, ", ")
    Call sqlValor(1, mValorCofins, ", ")
    Call sqlValor(1, mValorPisSt, ", ")
    Call sqlValor(1, mValorCofinsSt, ", ")
    Call sqlBoolean(1, mCombustivel, ", ")
    Call sqlBoolean(1, mCancelado, ", ")
    Call sqlTexto(1, mAguaEnergiaGasTelefone, ", ")
    Call sqlBoolean(1, mInutilizada, ", ")
    Call sqlBoolean(1, mIncidePisConfis, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlNumero(1, mIdUsuario, ", ")
    Call sqlNumero(1, mNumeroLote, ", ")
    Call sqlNumero(1, mNumeroRecepcao, ", ")
    Call sqlNumero(1, mNumeroProtocolo, ", ")
    Call sqlNumero(1, mEtapaConcluida, ", ")
    Call sqlTexto(1, mNomeEtapaConcluida, ", ")
    Call sqlTexto(1, mCodigoRetorno, ", ")
    Call sqlTexto(1, mMensagemRetorno, " ,")
    Call sqlNumero(1, mCodigoUltimoEvento, ", ")
    Call sqlTexto(1, mObservacaoEvento, ", ")
    Call sqlTexto(1, mProgramaOrigem, " ,")
    Call sqlTexto(1, mPeriodo, " ,")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mTipoSubEstoque, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
       Incluir = True
    '   Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mCodigoECF, mNumeroCupom, mData, mOrdem)
    End If
    Exit Function
    
Err_Incluir:
End Function
Public Function Alterar(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
    Dim lRecordsAffected As Long
    
    On Error GoTo Err_Alterar
    
    Alterar = False
    
    gSQL = "UPDATE MovimentoDocumentoEletronicoCabecalho SET "
    Call sqlNumero(2, "IdEstabelecimento_MovDECabecalho = ", mIdEstabelecimento)
    Call sqlData(2, ", DataEmissao_MovDECabecalho = ", mDataEmissao)
    Call sqlBoolean(2, ", Entrada_MovDECabecalho = ", mEntrada)
    Call sqlBoolean(2, ", Saida_MovDECabecalho = ", mSaida)
    Call sqlTexto(2, ", Modelo_MovDECabecalho = ", mModelo)
    Call sqlTexto(2, ", Serie_MovDECabecalho = ", mSerie)
    Call sqlTexto(2, ", Numero_MovDECabecalho = ", mNumero)
    Call sqlHora(2, ", HoraSaida_MovDECabecalho = ", mHoraSaida)
    Call sqlData(2, ", DataEntradaSaida_MovDECabecalho = ", mDataEntradaSaida)
    Call sqlBoolean(2, ", EmissaoPropria_MovDECabecalho = ", mEmissaoPropria)
    Call sqlNumero(2, ", IdClienteFornecedor_MovDECabecalho = ", mIdClienteFornecedor)
    Call sqlNumero(2, ", CodigoSituacao_MovDECabecalho = ", mCodigoSituacao)
    Call sqlTexto(2, ", ChaveAcesso_MovDECabecalho = ", mChaveAcesso)
    Call sqlTexto(2, ", FormaPagamento_MovDECabecalho = ", mFormaPagamento)
    Call sqlValor(2, ", ValorTotal_MovDECabecalho = ", mValorTotal)
    Call sqlValor(2, ", ValorDesconto_MovDECabecalho = ", mValorDesconto)
    Call sqlValor(2, ", ValorAbatimentoNaoTributado_MovDECabecalho = ", mValorAbatimentoNaoTributado)
    Call sqlValor(2, ", ValorProdutos_MovDECabecalho = ", mValorProdutos)
    Call sqlTexto(2, ", TipoFrete_MovDECabecalho = ", mTipoFrete)
    Call sqlValor(2, ", ValorFrete_MovDECabecalho = ", mValorFrete)
    Call sqlValor(2, ", ValorSeguro_MovDECabecalho = ", mValorSeguro)
    Call sqlValor(2, ", OutrasDespesas_MovDECabecalho = ", mOutrasDespesas)
    Call sqlValor(2, ", ValorBcIcms_MovDECabecalho = ", mValorBCICMS)
    Call sqlValor(2, ", AliquotaIcms_MovDECabecalho = ", mAliquotaICMS)
    Call sqlValor(2, ", ValorIcms_MovDECabecalho = ", mValorICMS)
    Call sqlValor(2, ", ValorBcIcmsSt_MovDECabecalho = ", mValorBCICMSST)
    Call sqlValor(2, ", AliquotaIcmsSt_MovDECabecalho = ", mAliquotaIcmsSt)
    Call sqlValor(2, ", ValorIcmsSt_MovDECabecalho = ", mValorICMSST)
    Call sqlValor(2, ", ValorIpi_MovDECabecalho = ", mValorIPI)
    Call sqlValor(2, ", ValorPis_MovDECabecalho = ", mValorPis)
    Call sqlValor(2, ", ValorCofins_MovDECabecalho = ", mValorCofins)
    Call sqlValor(2, ", ValorPisSt_MovDECabecalho = ", mValorPisSt)
    Call sqlValor(2, ", ValorCofinsSt_MovDECabecalho = ", mValorCofinsSt)
    Call sqlBoolean(2, ", Combustivel_MovDECabecalho = ", mCombustivel)
    Call sqlBoolean(2, ", Cancelado_MovDECabecalho = ", mCancelado)
    Call sqlTexto(2, ", AguaEnergiaGasTelefone_MovDECabecalho = ", mAguaEnergiaGasTelefone)
    Call sqlBoolean(2, ", Inutilizada_MovDECabecalho = ", mInutilizada)
    Call sqlBoolean(2, ", IncidePisCofins_MovDECabecalho = ", mIncidePisConfis)
    Call sqlData(2, ", DataDigitacao_MovDECabecalho = ", mDataDigitacao)
    Call sqlData(2, ", DataAlteracao_MovDECabecalho = ", mDataAlteracao)
    Call sqlNumero(2, ", IdUsuario_MovDECabecalho = ", mIdUsuario)
    Call sqlNumero(2, ", NumeroLote_MovDECabecalho = ", mNumeroLote)
    Call sqlNumero(2, ", NumeroRecepcao_MovDECabecalho = ", mNumeroRecepcao)
    Call sqlNumero(2, ", NumeroProtocolo_MovDECabecalho = ", mNumeroProtocolo)
    Call sqlNumero(2, ", EtapaConcluida_MovDECabecalho = ", mEtapaConcluida)
    Call sqlTexto(2, ", NomeEtapaConcluida_MovDECabecalho = ", mNomeEtapaConcluida)
    Call sqlTexto(2, ", CodigoRetorno_MovDECabecalho = ", mCodigoRetorno)
    Call sqlTexto(2, ", MensagemRetorno_MovDECabecalho = ", mMensagemRetorno)
    Call sqlNumero(2, ", CodigoUltimoEvento_MovDECabecalho = ", mCodigoUltimoEvento)
    Call sqlTexto(2, ", ObservacaoEvento_MovDECabecalho = ", mObservacaoEvento)
    Call sqlTexto(2, ", ProgramaOrigem_MovDECabecalho = ", mProgramaOrigem)
    Call sqlTexto(2, ", Periodo_MovDECabecalho = ", mPeriodo)
    Call sqlNumero(2, ", TipoMovimento_MovDECabecalho = ", mTipoMovimento)
    Call sqlNumero(2, ", TipoSubEstoque_MovDECabecalho = ", mTipoSubEstoque)
    
    
    
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    gSQL = gSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    gSQL = gSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    gSQL = gSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    gSQL = gSQL & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    gSQL = gSQL & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    '        Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function


Err_Alterar:

End Function

Public Function AlterarInformacoesEtapaPagamento(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarInformacoesEtapaPagamento = False
    gSQL = "UPDATE MovimentoDocumentoEletronicoCabecalho SET "
    Call sqlTexto(2, " FormaPagamento_MovDECabecalho = ", mFormaPagamento)
    Call sqlValor(2, ", ValorTotal_MovDECabecalho = ", mValorTotal)
    Call sqlValor(2, ", ValorDesconto_MovDECabecalho = ", mValorDesconto)
    Call sqlValor(2, ", ValorProdutos_MovDECabecalho = ", mValorProdutos)
    Call sqlNumero(2, ", EtapaConcluida_MovDECabecalho = ", mEtapaConcluida)
       
    
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    gSQL = gSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    gSQL = gSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    gSQL = gSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    gSQL = gSQL & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    gSQL = gSQL & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        AlterarInformacoesEtapaPagamento = True
        'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, 0)
    End If
    Exit Function

Err_Alterar:
End Function
Public Function DefinirUtlimoEventoDocumento(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

        DefinirUtlimoEventoDocumento = False
            gSQL = "UPDATE MovimentoDocumentoEletronicoCabecalho SET "
            Call sqlNumero(2, " CodigoUltimoEvento_MovDECabecalho = ", mCodigoUltimoEvento)
            Call sqlTexto(2, ", ObservacaoEvento_MovDECabecalho = ", mObservacaoEvento)


            gSQL = gSQL & " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
            gSQL = gSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
            gSQL = gSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
            gSQL = gSQL & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
            gSQL = gSQL & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)


    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        DefinirUtlimoEventoDocumento = True
    End If


    Exit Function

Err_Alterar:
End Function
Public Function AlterarValoresTributacaoCabecalho(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Alterar

    AlterarValoresTributacaoCabecalho = False

    gSQL = "SELECT  SUM(ValorBcIcms_MovDEItem) AS TotalBcIcms, SUM(ValorIcms_MovDEItem) AS TotalIcms, SUM(ValorPis_MovDEItem) AS TotalPis, SUM(ValorCofins_MovDEItem) AS TotalCofins"
    gSQL = gSQL & "  FROM MovimentoDocumentoEletronicoItem"
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
    gSQL = gSQL & "   AND Cancelado_MovDEItem = " & preparaBooleano(False)
    
    If Localizar(1) Then
        If rsMovDocumentoEletronicoCabecalho.RecordCount > 0 Then
        
            If IsNull(rsMovDocumentoEletronicoCabecalho("TotalBcIcms").Value) Or IsNull(rsMovDocumentoEletronicoCabecalho("TotalIcms").Value) Or IsNull(rsMovDocumentoEletronicoCabecalho("TotalPis").Value) Or IsNull(rsMovDocumentoEletronicoCabecalho("TotalCofins").Value) Then
                Call CriaLogCadastroDll("[AlterarValoresTributacaoCabecalho] Registro não foi encontrado" & gSQL)
                Exit Function
            End If
            gSQL = "UPDATE MovimentoDocumentoEletronicoCabecalho SET "
            Call sqlTexto(2, " ValorBcIcms_MovDECabecalho = ", rsMovDocumentoEletronicoCabecalho("TotalBcIcms").Value)
            Call sqlValor(2, ", ValorIcms_MovDECabecalho = ", rsMovDocumentoEletronicoCabecalho("TotalIcms").Value)
            Call sqlValor(2, ", ValorPis_MovDECabecalho = ", rsMovDocumentoEletronicoCabecalho("TotalPis").Value)
            Call sqlValor(2, ", ValorCofins_MovDECabecalho = ", rsMovDocumentoEletronicoCabecalho("TotalCofins").Value)
            
            gSQL = gSQL & " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
            gSQL = gSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
            gSQL = gSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
            gSQL = gSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
            gSQL = gSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
            gSQL = gSQL & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
            gSQL = gSQL & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
            
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
            
            
            If lRecordsAffected > 0 Then
                AlterarValoresTributacaoCabecalho = True
            End If
        End If
    End If

    Exit Function

Err_Alterar:
End Function
Public Function Excluir(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
    
    Dim lRecordsAffected As Long
    Dim xCondicao As String

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoDocumentoEletronicoCabecalho"
    xCondicao = " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    
    gSQL = gSQL & xCondicao
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:

End Function

Public Function AlterarIdCliente(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pIdCliente As Integer) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarIdCliente = False
    gSQL = "UPDATE MovimentoDocumentoEletronicoCabecalho SET "
    Call sqlNumero(2, "IdClienteFornecedor_MovDECabecalho = ", pIdCliente)
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    gSQL = gSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    gSQL = gSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    gSQL = gSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    gSQL = gSQL & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    gSQL = gSQL & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarIdCliente = True
    End If
    Exit Function

Err_Alterar:
End Function

Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovDocumentoEletronicoCabecalho = New adodb.Recordset
    rsMovDocumentoEletronicoCabecalho.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovDocumentoEletronicoCabecalho.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovDocumentoEletronicoCabecalho.EOF Then
        AtualizaRecordset = True
    End If
End Function

Private Sub AtribuiValor()


    IdEstabelecimento = rsMovDocumentoEletronicoCabecalho("IdEstabelecimento_MovDECabecalho").Value
    DataEmissao = rsMovDocumentoEletronicoCabecalho("DataEmissao_MovDECabecalho").Value
    Entrada = rsMovDocumentoEletronicoCabecalho("Entrada_MovDECabecalho").Value
    Saida = rsMovDocumentoEletronicoCabecalho("Saida_MovDECabecalho").Value
    Modelo = rsMovDocumentoEletronicoCabecalho("Modelo_MovDECabecalho").Value
    Serie = rsMovDocumentoEletronicoCabecalho("Serie_MovDECabecalho").Value
    Numero = rsMovDocumentoEletronicoCabecalho("Numero_MovDECabecalho").Value
    HoraSaida = rsMovDocumentoEletronicoCabecalho("HoraSaida_MovDECabecalho").Value
    DataEntradaSaida = rsMovDocumentoEletronicoCabecalho("DataEntradaSaida_MovDECabecalho").Value
    EmissaoPropria = rsMovDocumentoEletronicoCabecalho("EmissaoPropria_MovDECabecalho").Value
    IdClienteFornecedor = rsMovDocumentoEletronicoCabecalho("IdClienteFornecedor_MovDECabecalho").Value
    CodigoSituacao = rsMovDocumentoEletronicoCabecalho("CodigoSituacao_MovDECabecalho").Value
    ChaveAcesso = rsMovDocumentoEletronicoCabecalho("ChaveAcesso_MovDECabecalho").Value
    FormaPagamento = rsMovDocumentoEletronicoCabecalho("FormaPagamento_MovDECabecalho").Value
    ValorTotal = rsMovDocumentoEletronicoCabecalho("ValorTotal_MovDECabecalho").Value
    ValorDesconto = rsMovDocumentoEletronicoCabecalho("ValorDesconto_MovDECabecalho").Value
    ValorAbatimentoNaoTributado = rsMovDocumentoEletronicoCabecalho("ValorAbatimentoNaoTributado_MovDECabecalho").Value
    ValorProdutos = rsMovDocumentoEletronicoCabecalho("ValorProdutos_MovDECabecalho").Value
    TipoFrete = rsMovDocumentoEletronicoCabecalho("TipoFrete_MovDECabecalho").Value
    ValorFrete = rsMovDocumentoEletronicoCabecalho("ValorFrete_MovDECabecalho").Value
    ValorSeguro = rsMovDocumentoEletronicoCabecalho("ValorSeguro_MovDECabecalho").Value
    OutrasDespesas = rsMovDocumentoEletronicoCabecalho("OutrasDespesas_MovDECabecalho").Value
    ValorBCICMS = rsMovDocumentoEletronicoCabecalho("ValorBcIcms_MovDECabecalho").Value
    AliquotaICMS = rsMovDocumentoEletronicoCabecalho("AliquotaIcms_MovDECabecalho").Value
    ValorICMS = rsMovDocumentoEletronicoCabecalho("ValorIcms_MovDECabecalho").Value
    ValorBCICMSST = rsMovDocumentoEletronicoCabecalho("ValorBcIcmsSt_MovDECabecalho").Value
    AliquotaIcmsSt = rsMovDocumentoEletronicoCabecalho("AliquotaIcmsSt_MovDECabecalho").Value
    ValorICMSST = rsMovDocumentoEletronicoCabecalho("ValorIcmsSt_MovDECabecalho").Value
    ValorIPI = rsMovDocumentoEletronicoCabecalho("ValorIpi_MovDECabecalho").Value
    ValorPis = rsMovDocumentoEletronicoCabecalho("ValorPis_MovDECabecalho").Value
    ValorCofins = rsMovDocumentoEletronicoCabecalho("ValorCofins_MovDECabecalho").Value
    ValorPisSt = rsMovDocumentoEletronicoCabecalho("ValorPisSt_MovDECabecalho").Value
    ValorCofinsSt = rsMovDocumentoEletronicoCabecalho("ValorCofinsSt_MovDECabecalho").Value
    Combustivel = rsMovDocumentoEletronicoCabecalho("Combustivel_MovDECabecalho").Value
    Cancelado = rsMovDocumentoEletronicoCabecalho("Cancelado_MovDECabecalho").Value
    AguaEnergiaGasTelefone = rsMovDocumentoEletronicoCabecalho("AguaEnergiaGasTelefone_MovDECabecalho").Value
    Inutilizada = rsMovDocumentoEletronicoCabecalho("Inutilizada_MovDECabecalho").Value
    IncidePisConfis = rsMovDocumentoEletronicoCabecalho("IncidePisCofins_MovDECabecalho").Value
    DataDigitacao = rsMovDocumentoEletronicoCabecalho("DataDigitacao_MovDECabecalho").Value
    If IsNull(rsMovDocumentoEletronicoCabecalho("DataAlteracao_MovDECabecalho").Value) Then
        DataAlteracao = "00:00:00"
    Else
        DataAlteracao = rsMovDocumentoEletronicoCabecalho("DataAlteracao_MovDECabecalho").Value
    End If
    IdUsuario = rsMovDocumentoEletronicoCabecalho("IdUsuario_MovDECabecalho").Value
    NumeroLote = rsMovDocumentoEletronicoCabecalho("NumeroLote_MovDECabecalho").Value
    NumeroRecepcao = rsMovDocumentoEletronicoCabecalho("NumeroRecepcao_MovDECabecalho").Value
    NumeroProtocolo = rsMovDocumentoEletronicoCabecalho("NumeroProtocolo_MovDECabecalho").Value
    EtapaConcluida = rsMovDocumentoEletronicoCabecalho("EtapaConcluida_MovDECabecalho").Value
    NomeEtapaConcluida = rsMovDocumentoEletronicoCabecalho("NomeEtapaConcluida_MovDECabecalho").Value
    CodigoRetorno = rsMovDocumentoEletronicoCabecalho("CodigoRetorno_MovDECabecalho").Value
    MensagemRetorno = rsMovDocumentoEletronicoCabecalho("MensagemRetorno_MovDECabecalho").Value
    CodigoUltimoEvento = rsMovDocumentoEletronicoCabecalho("CodigoUltimoEvento_MovDECabecalho").Value
    ObservacaoEvento = rsMovDocumentoEletronicoCabecalho("ObservacaoEvento_MovDECabecalho").Value
    ProgramaOrigem = rsMovDocumentoEletronicoCabecalho("ProgramaOrigem_MovDECabecalho").Value
    
    Periodo = rsMovDocumentoEletronicoCabecalho("Periodo_MovDECabecalho").Value
    TipoMovimento = rsMovDocumentoEletronicoCabecalho("TipoMovimento_MovDECabecalho").Value
    TipoSubEstoque = rsMovDocumentoEletronicoCabecalho("TipoSubEstoque_MovDECabecalho").Value
    
   
End Sub

Public Function ExisteDocumentoEletronico(ByVal pIdEstabelecimento As Integer, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteDocumentoEletronico = False
    xCondicao = " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    
    gSQL = PreparaSQL(xCondicao, "")
    ExisteDocumentoEletronico = Localizar(1)
    
    rsMovDocumentoEletronicoCabecalho.Close
    Set rsMovDocumentoEletronicoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function LocalizarCodigo(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE IdEstabelecimento_MovDECabecalho = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Entrada_MovDECabecalho = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDECabecalho = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDECabecalho = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDECabecalho = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDECabecalho = " & preparaTexto(pNumero)
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovDocumentoEletronicoCabecalho.Close
    Set rsMovDocumentoEletronicoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xSQL As String
    xSQL = "SELECT IdEstabelecimento_MovDECabecalho,DataEmissao_MovDECabecalho,Entrada_MovDECabecalho,Saida_MovDECabecalho,Modelo_MovDECabecalho, "
    xSQL = xSQL & "Serie_MovDECabecalho,Numero_MovDECabecalho,HoraSaida_MovDECabecalho,DataEntradaSaida_MovDECabecalho,EmissaoPropria_MovDECabecalho, "
    xSQL = xSQL & "IdClienteFornecedor_MovDECabecalho,CodigoSituacao_MovDECabecalho,ChaveAcesso_MovDECabecalho,FormaPagamento_MovDECabecalho,ValorTotal_MovDECabecalho, "
    xSQL = xSQL & "ValorDesconto_MovDECabecalho,ValorAbatimentoNaoTributado_MovDECabecalho,ValorProdutos_MovDECabecalho,TipoFrete_MovDECabecalho,ValorFrete_MovDECabecalho, "
    xSQL = xSQL & "ValorSeguro_MovDECabecalho,OutrasDespesas_MovDECabecalho,ValorBcIcms_MovDECabecalho,AliquotaIcms_MovDECabecalho,ValorIcms_MovDECabecalho, "
    xSQL = xSQL & "ValorBcIcmsSt_MovDECabecalho,AliquotaIcmsSt_MovDECabecalho,ValorIcmsSt_MovDECabecalho,ValorIpi_MovDECabecalho,ValorPis_MovDECabecalho, "
    xSQL = xSQL & "ValorCofins_MovDECabecalho,ValorPisSt_MovDECabecalho,ValorCofinsSt_MovDECabecalho,Combustivel_MovDECabecalho,Cancelado_MovDECabecalho, "
    xSQL = xSQL & "AguaEnergiaGasTelefone_MovDECabecalho,Inutilizada_MovDECabecalho,IncidePisCofins_MovDECabecalho,DataDigitacao_MovDECabecalho,DataAlteracao_MovDECabecalho, "
    xSQL = xSQL & "IdUsuario_MovDECabecalho, NumeroLote_MovDECabecalho, NumeroRecepcao_MovDECabecalho, NumeroProtocolo_MovDECabecalho, EtapaConcluida_MovDECabecalho, "
    xSQL = xSQL & "NomeEtapaConcluida_MovDECabecalho, CodigoRetorno_MovDECabecalho, MensagemRetorno_MovDECabecalho, CodigoUltimoEvento_MovDECabecalho, ObservacaoEvento_MovDECabecalho, ProgramaOrigem_MovDECabecalho, "
    xSQL = xSQL & "Periodo_MovDECabecalho, TipoMovimento_MovDECabecalho, TipoSubEstoque_MovDECabecalho "
    xSQL = xSQL & " FROM MovimentoDocumentoEletronicoCabecalho"
    
    PreparaSQL = Trim(xSQL & " " & xCondicao & " " & xTipoOrdem)
End Function
