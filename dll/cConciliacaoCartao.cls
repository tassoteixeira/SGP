VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConciliacaoCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoConciliacao As String
Private mData As Date
Private mCodigoCartao As Integer
Private mTipoOperacao As String
Private mTotalBruto As Currency
Private mTotalLiquido As Currency
Private mQuantidade As Integer
Private mTaxaAdministrativa As Currency
Private mDataProcessamento As Date
Private mTotalPendencia As Currency
Private mQuantidadePendencia As Integer
Private mTextoPendencia As String
Private mDataPendenciaResolvida As Date

Private rsConciliacaoCartao As New adodb.Recordset



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


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get CodigoCartao() As Integer
    CodigoCartao = mCodigoCartao
End Property
Public Property Let CodigoCartao(ByVal Valor As Integer)
    mCodigoCartao = Valor
End Property


Public Property Get TipoOperacao() As String
    TipoOperacao = mTipoOperacao
End Property
Public Property Let TipoOperacao(ByVal Valor As String)
    mTipoOperacao = Valor
End Property


Public Property Get TotalBruto() As Currency
    TotalBruto = mTotalBruto
End Property
Public Property Let TotalBruto(ByVal Valor As Currency)
    mTotalBruto = Valor
End Property


Public Property Get TotalLiquido() As Currency
    TotalLiquido = mTotalLiquido
End Property
Public Property Let TotalLiquido(ByVal Valor As Currency)
    mTotalLiquido = Valor
End Property


Public Property Get Quantidade() As Integer
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Integer)
    mQuantidade = Valor
End Property


Public Property Get TaxaAdministrativa() As Currency
    TaxaAdministrativa = mTaxaAdministrativa
End Property
Public Property Let TaxaAdministrativa(ByVal Valor As Currency)
    mTaxaAdministrativa = Valor
End Property


Public Property Get DataProcessamento() As Date
    DataProcessamento = mDataProcessamento
End Property
Public Property Let DataProcessamento(ByVal Valor As Date)
    mDataProcessamento = Valor
End Property


Public Property Get TotalPendencia() As Currency
    TotalPendencia = mTotalPendencia
End Property
Public Property Let TotalPendencia(ByVal Valor As Currency)
    mTotalPendencia = Valor
End Property


Public Property Get QuantidadePendencia() As Integer
    QuantidadePendencia = mQuantidadePendencia
End Property
Public Property Let QuantidadePendencia(ByVal Valor As Integer)
    mQuantidadePendencia = Valor
End Property


Public Property Get TextoPendencia() As String
    TextoPendencia = mTextoPendencia
End Property
Public Property Let TextoPendencia(ByVal Valor As String)
    mTextoPendencia = Valor
End Property


Public Property Get DataPendenciaResolvida() As Date
    DataPendenciaResolvida = mDataPendenciaResolvida
End Property
Public Property Let DataPendenciaResolvida(ByVal Valor As Date)
    mDataPendenciaResolvida = Valor
End Property
'Fim das Propriedades da Classe




'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pTipoConciliacao As String, ByVal pData As Date, ByVal pCodigoCartao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Conciliacao] = " & preparaTexto(pTipoConciliacao)
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Codigo do Cartao] = " & pCodigoCartao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsConciliacaoCartao.Close
    Set rsConciliacaoCartao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


'Início Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Tipo de Conciliacao], Data, [Codigo do Cartao], [Tipo da Operacao], [Total Bruto], [Total Liquido], Quantidade, [Taxa Administrativa], [Data do Processamento], [Total de Pendencia], [Quantidade de Pendencia], [Texto da Pendencia], [Data da Pendencia Resolvida] FROM ConciliacaoCartao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConciliacaoCartao = New adodb.Recordset
    rsConciliacaoCartao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConciliacaoCartao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConciliacaoCartao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsConciliacaoCartao("Empresa").Value
    TipoConciliacao = rsConciliacaoCartao("Tipo de Conciliacao").Value
    Data = rsConciliacaoCartao("Data").Value
    CodigoCartao = rsConciliacaoCartao("Codigo do Cartao").Value
    TipoOperacao = rsConciliacaoCartao("Tipo da Operacao").Value
    TotalBruto = rsConciliacaoCartao("Total Bruto").Value
    TotalLiquido = rsConciliacaoCartao("Total Liquido").Value
    Quantidade = rsConciliacaoCartao("Quantidade").Value
    TaxaAdministrativa = rsConciliacaoCartao("Taxa Administrativa").Value
    DataProcessamento = rsConciliacaoCartao("Data do Processamento").Value
    TotalPendencia = rsConciliacaoCartao("Total de Pendencia").Value
    QuantidadePendencia = rsConciliacaoCartao("Quantidade de Pendencia").Value
    TextoPendencia = rsConciliacaoCartao("Texto da Pendencia").Value
    If IsNull(rsConciliacaoCartao("Data da Pendencia Resolvida").Value) Then
        DataPendenciaResolvida = "00:00:00"
    Else
        DataPendenciaResolvida = rsConciliacaoCartao("Data da Pendencia Resolvida").Value
    End If
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

'Fim Funções / Procedures internas

