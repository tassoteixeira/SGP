VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConciliacaoCartaoPendente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoConciliacao As String
Private mDataEmissao As Date
Private mPeriodo As String
Private mTipoMovimento As String
Private mNumeroLancamento As Integer
Private mCodigoCartao As Integer
Private mDataVencimento As Date
Private mValor As Currency
Private mNumeroMovimentoCaixa As Long
Private mTaxaAdministrativa As Currency
Private mNumeroIlha As Integer
Private mAutorizacao As String
Private mNSU As String
Private mCodigoFuncionario As Integer
Private mTipoOperacao As String
Private mHoraTransacao As Date
Private mObservacao As String
Private mOrigemERP As Boolean
Private mVersaoSistema As String
Private mDataProcessamento As Date
Private mTextoPendencia As String
Private mDataConciliacaoTransacao As Date

'Private rsConciliacaoCartaoPendente As New adodb.Recordset



'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoConciliacao() As String
    TipoConciliacao = mTipoConciliacao
End Property
Public Property Let TipoConciliacao(ByVal Valor As String)
    mTipoConciliacao = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property


Public Property Get TipoMovimento() As String
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As String)
    mTipoMovimento = Valor
End Property


Public Property Get NumeroLancamento() As Integer
    NumeroLancamento = mNumeroLancamento
End Property
Public Property Let NumeroLancamento(ByVal Valor As Integer)
    mNumeroLancamento = Valor
End Property


Public Property Get CodigoCartao() As Integer
    CodigoCartao = mCodigoCartao
End Property
Public Property Let CodigoCartao(ByVal Valor As Integer)
    mCodigoCartao = Valor
End Property


Public Property Get DataVencimento() As Date
    DataVencimento = mDataVencimento
End Property
Public Property Let DataVencimento(ByVal Valor As Date)
    mDataVencimento = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get TaxaAdministrativa() As Currency
    TaxaAdministrativa = mTaxaAdministrativa
End Property
Public Property Let TaxaAdministrativa(ByVal Valor As Currency)
    mTaxaAdministrativa = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get Autorizacao() As String
    Autorizacao = mAutorizacao
End Property
Public Property Let Autorizacao(ByVal Valor As String)
    mAutorizacao = Valor
End Property


Public Property Get NSU() As String
    NSU = mNSU
End Property
Public Property Let NSU(ByVal Valor As String)
    mNSU = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get TipoOperacao() As String
    TipoOperacao = mTipoOperacao
End Property
Public Property Let TipoOperacao(ByVal Valor As String)
    mTipoOperacao = Valor
End Property


Public Property Get HoraTransacao() As Date
    HoraTransacao = mHoraTransacao
End Property
Public Property Let HoraTransacao(ByVal Valor As Date)
    mHoraTransacao = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get OrigemERP() As Boolean
    OrigemERP = mOrigemERP
End Property
Public Property Let OrigemERP(ByVal Valor As Boolean)
    mOrigemERP = Valor
End Property


Public Property Get VersaoSistema() As String
    VersaoSistema = mVersaoSistema
End Property
Public Property Let VersaoSistema(ByVal Valor As String)
    mVersaoSistema = Valor
End Property


Public Property Get DataProcessamento() As Date
    DataProcessamento = mDataProcessamento
End Property
Public Property Let DataProcessamento(ByVal Valor As Date)
    mDataProcessamento = Valor
End Property


Public Property Get TextoPendencia() As String
    TextoPendencia = mTextoPendencia
End Property
Public Property Let TextoPendencia(ByVal Valor As String)
    mTextoPendencia = Valor
End Property


Public Property Get DataConciliacaoTransacao() As Date
    DataConciliacaoTransacao = mDataConciliacaoTransacao
End Property
Public Property Let DataConciliacaoTransacao(ByVal Valor As Date)
    mDataConciliacaoTransacao = Valor
End Property


'Fim das Propriedades da Classe

