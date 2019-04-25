VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovNotaFiscalSaidaCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mSerie As String
Private mNumero As Long
Private mDataEmissao As Date
Private mHoraSaida As Date
Private mCodigoCliente As Long
Private mCFOP As String
Private mCancelada As Boolean
Private mSubEstoque As Integer
Private mFormaPagamento As String
Private mTotalProduto As Currency
Private mDesconto As Currency
Private mAcrescimo As Currency
Private mTotalNota As Currency
Private mBaseCalculoICMS As Currency
Private mValorICMS As Currency
Private mBaseCalculoSubstituicao As Currency
Private mValorSubstituicao As Currency
Private mValorFrete As Currency
Private mValorSeguro As Currency
Private mValorIPI As Currency
Private mSubstituicao As Currency
Private mOutros As Currency
Private mInformacaoComplementar1 As String
Private mInformacaoComplementar2 As String
Private mInformacaoComplementar3 As String
Private mInformacaoComplementar4 As String
Private mInformacaoComplementar5 As String
Private mInformacaoComplementar6 As String
Private mDataDigitacao As Date
Private mDataAlteracao As Date
Private mCodigoUsuario As Integer
Private mObservacao As String
Private mCodigoClienteTemporario As Long

Private rsMovNotaFiscalSaidaCabecalho As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Serie() As String
    Serie = mSerie
End Property
Public Property Let Serie(ByVal Valor As String)
    mSerie = Valor
End Property


Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get HoraSaida() As Date
    HoraSaida = mHoraSaida
End Property
Public Property Let HoraSaida(ByVal Valor As Date)
    mHoraSaida = Valor
End Property


Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get CFOP() As String
    CFOP = mCFOP
End Property
Public Property Let CFOP(ByVal Valor As String)
    mCFOP = Valor
End Property


Public Property Get Cancelada() As Boolean
    Cancelada = mCancelada
End Property
Public Property Let Cancelada(ByVal Valor As Boolean)
    mCancelada = Valor
End Property


Public Property Get SubEstoque() As Integer
    SubEstoque = mSubEstoque
End Property
Public Property Let SubEstoque(ByVal Valor As Integer)
    mSubEstoque = Valor
End Property


Public Property Get FormaPagamento() As String
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As String)
    mFormaPagamento = Valor
End Property


Public Property Get TotalProduto() As Currency
    TotalProduto = mTotalProduto
End Property
Public Property Let TotalProduto(ByVal Valor As Currency)
    mTotalProduto = Valor
End Property


Public Property Get Desconto() As Currency
    Desconto = mDesconto
End Property
Public Property Let Desconto(ByVal Valor As Currency)
    mDesconto = Valor
End Property


Public Property Get Acrescimo() As Currency
    Acrescimo = mAcrescimo
End Property
Public Property Let Acrescimo(ByVal Valor As Currency)
    mAcrescimo = Valor
End Property


Public Property Get TotalNota() As Currency
    TotalNota = mTotalNota
End Property
Public Property Let TotalNota(ByVal Valor As Currency)
    mTotalNota = Valor
End Property


Public Property Get BaseCalculoICMS() As Currency
    BaseCalculoICMS = mBaseCalculoICMS
End Property
Public Property Let BaseCalculoICMS(ByVal Valor As Currency)
    mBaseCalculoICMS = Valor
End Property


Public Property Get ValorICMS() As Currency
    ValorICMS = mValorICMS
End Property
Public Property Let ValorICMS(ByVal Valor As Currency)
    mValorICMS = Valor
End Property


Public Property Get BaseCalculoSubstituicao() As Currency
    BaseCalculoSubstituicao = mBaseCalculoSubstituicao
End Property
Public Property Let BaseCalculoSubstituicao(ByVal Valor As Currency)
    mBaseCalculoSubstituicao = Valor
End Property


Public Property Get ValorSubstituicao() As Currency
    ValorSubstituicao = mValorSubstituicao
End Property
Public Property Let ValorSubstituicao(ByVal Valor As Currency)
    mValorSubstituicao = Valor
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


Public Property Get ValorIPI() As Currency
    ValorIPI = mValorIPI
End Property
Public Property Let ValorIPI(ByVal Valor As Currency)
    mValorIPI = Valor
End Property


Public Property Get Substituicao() As Currency
    Substituicao = mSubstituicao
End Property
Public Property Let Substituicao(ByVal Valor As Currency)
    mSubstituicao = Valor
End Property


Public Property Get Outros() As Currency
    Outros = mOutros
End Property
Public Property Let Outros(ByVal Valor As Currency)
    mOutros = Valor
End Property


Public Property Get InformacaoComplementar1() As String
    InformacaoComplementar1 = mInformacaoComplementar1
End Property
Public Property Let InformacaoComplementar1(ByVal Valor As String)
    mInformacaoComplementar1 = Valor
End Property


Public Property Get InformacaoComplementar2() As String
    InformacaoComplementar2 = mInformacaoComplementar2
End Property
Public Property Let InformacaoComplementar2(ByVal Valor As String)
    mInformacaoComplementar2 = Valor
End Property


Public Property Get InformacaoComplementar3() As String
    InformacaoComplementar3 = mInformacaoComplementar3
End Property
Public Property Let InformacaoComplementar3(ByVal Valor As String)
    mInformacaoComplementar3 = Valor
End Property


Public Property Get InformacaoComplementar4() As String
    InformacaoComplementar4 = mInformacaoComplementar4
End Property
Public Property Let InformacaoComplementar4(ByVal Valor As String)
    mInformacaoComplementar4 = Valor
End Property


Public Property Get InformacaoComplementar5() As String
    InformacaoComplementar5 = mInformacaoComplementar5
End Property
Public Property Let InformacaoComplementar5(ByVal Valor As String)
    mInformacaoComplementar5 = Valor
End Property


Public Property Get InformacaoComplementar6() As String
    InformacaoComplementar6 = mInformacaoComplementar6
End Property
Public Property Let InformacaoComplementar6(ByVal Valor As String)
    mInformacaoComplementar6 = Valor
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


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get CodigoClienteTemporario() As Long
    CodigoClienteTemporario = mCodigoClienteTemporario
End Property
Public Property Let CodigoClienteTemporario(ByVal Valor As Long)
    mCodigoClienteTemporario = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & "   AND Data = " & preparaData(pData)
    xCondicao = xCondicao & "   AND Serie = " & preparaTexto(pSerie)
    xCondicao = xCondicao & "   AND Numero = " & pNumero
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovNotaFiscalSaidaCabecalho.Close
    Set rsMovNotaFiscalSaidaCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoNotaFiscalSaidaCabecalho ( Empresa, Data, Serie, Numero, [Data de Emissao], "
    gSQL = gSQL & "[Hora de Saida], [Codigo de Cliente], CFOP, Cancelada, SubEstoque, "
    gSQL = gSQL & "[Forma de Pagamento], [Total de Produto], Desconto, Acrescimo, [Total da Nota], "
    gSQL = gSQL & "[Base de Calculo de ICMS], [Valor de ICMS], [Base de Calculo de Substituicao], [Valor de Substituicao], [Valor de Frete], "
    gSQL = gSQL & "[Valor de Seguro], [Valor de IPI], Substituicao, Outros, [Informacao Complementar 1], "
    gSQL = gSQL & "[Informacao Complementar 2], [Informacao Complementar 3], [Informacao Complementar 4], [Informacao Complementar 5], [Informacao Complementar 6], "
    gSQL = gSQL & "[Data de Digitacao], [Data de Alteracao], [Codigo de Usuario], Observacao, [Codigo de Cliente Temporario] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlHora(1, mHoraSaida, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlTexto(1, mCFOP, ", ")
    Call sqlBoolean(1, mCancelada, ", ")
    Call sqlNumero(1, mSubEstoque, ", ")
    Call sqlTexto(1, mFormaPagamento, ", ")
    Call sqlValor(1, mTotalProduto, ", ")
    Call sqlValor(1, mDesconto, ", ")
    Call sqlValor(1, mAcrescimo, ", ")
    Call sqlValor(1, mTotalNota, ", ")
    Call sqlValor(1, mBaseCalculoICMS, ", ")
    Call sqlValor(1, mValorICMS, ", ")
    Call sqlValor(1, mBaseCalculoSubstituicao, ", ")
    Call sqlValor(1, mValorSubstituicao, ", ")
    Call sqlValor(1, mValorFrete, ", ")
    Call sqlValor(1, mValorSeguro, ", ")
    Call sqlValor(1, mValorIPI, ", ")
    Call sqlValor(1, mSubstituicao, ", ")
    Call sqlValor(1, mOutros, ", ")
    Call sqlTexto(1, mInformacaoComplementar1, ", ")
    Call sqlTexto(1, mInformacaoComplementar2, ", ")
    Call sqlTexto(1, mInformacaoComplementar3, ", ")
    Call sqlTexto(1, mInformacaoComplementar4, ", ")
    Call sqlTexto(1, mInformacaoComplementar5, ", ")
    Call sqlTexto(1, mInformacaoComplementar6, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlTexto(1, mObservacao, ", ")
    Call sqlNumero(1, mCodigoClienteTemporario, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoNotaFiscalSaidaCabecalho SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlNumero(2, ", Numero = ", mNumero)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlHora(2, ", [Hora de Saida] = ", mHoraSaida)
    Call sqlNumero(2, ", [Codigo de Cliente] = ", mCodigoCliente)
    Call sqlTexto(2, ", CFOP = ", mCFOP)
    Call sqlBoolean(2, ", Cancelada = ", mCancelada)
    Call sqlNumero(2, ", SubEstoque = ", mSubEstoque)
    Call sqlTexto(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Total de Produto] = ", mTotalProduto)
    Call sqlValor(2, ", Desconto = ", mDesconto)
    Call sqlValor(2, ", Acrescimo = ", mAcrescimo)
    Call sqlValor(2, ", [Total da Nota] = ", mTotalNota)
    Call sqlValor(2, ", [Base de Calculo de ICMS] = ", mBaseCalculoICMS)
    Call sqlValor(2, ", [Valor de ICMS] = ", mValorICMS)
    Call sqlValor(2, ", [Base de Calculo de Substituicao] = ", mBaseCalculoSubstituicao)
    Call sqlValor(2, ", [Valor de Substituicao] = ", mValorSubstituicao)
    Call sqlValor(2, ", [Valor de Frete] = ", mValorFrete)
    Call sqlValor(2, ", [Valor de Seguro] = ", mValorSeguro)
    Call sqlValor(2, ", [Valor de IPI] = ", mValorIPI)
    Call sqlValor(2, ", Substituicao = ", mSubstituicao)
    Call sqlValor(2, ", Outros = ", mOutros)
    Call sqlTexto(2, ", [Informacao Complementar 1] = ", mInformacaoComplementar1)
    Call sqlTexto(2, ", [Informacao Complementar 2] = ", mInformacaoComplementar2)
    Call sqlTexto(2, ", [Informacao Complementar 3] = ", mInformacaoComplementar3)
    Call sqlTexto(2, ", [Informacao Complementar 4] = ", mInformacaoComplementar4)
    Call sqlTexto(2, ", [Informacao Complementar 5] = ", mInformacaoComplementar5)
    Call sqlTexto(2, ", [Informacao Complementar 6] = ", mInformacaoComplementar6)
    Call sqlData(2, ", [Data de Digitacao] = ", mDataDigitacao)
    Call sqlData(2, ", [Data de Alteracao] = ", mDataAlteracao)
    Call sqlNumero(2, ", [Codigo de Usuario] = ", mCodigoUsuario)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    Call sqlNumero(2, ", [Codigo de Cliente Temporario] = ", mCodigoClienteTemporario)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Serie = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoNotaFiscalSaidaCabecalho"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Serie = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Serie, Numero, [Data de Emissao], [Hora de Saida], [Codigo de Cliente], CFOP, Cancelada, SubEstoque, [Forma de Pagamento], [Total de Produto], Desconto, Acrescimo, [Total da Nota], [Base de Calculo de ICMS], [Valor de ICMS], [Base de Calculo de Substituicao], [Valor de Substituicao], [Valor de Frete], [Valor de Seguro], [Valor de IPI], Substituicao, Outros, [Informacao Complementar 1], [Informacao Complementar 2], [Informacao Complementar 3], [Informacao Complementar 4], [Informacao Complementar 5], [Informacao Complementar 6], [Data de Digitacao], [Data de Alteracao], [Codigo de Usuario], Observacao, [Codigo de Cliente Temporario] FROM MovimentoNotaFiscalSaidaCabecalho" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovNotaFiscalSaidaCabecalho = New adodb.Recordset
    rsMovNotaFiscalSaidaCabecalho.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovNotaFiscalSaidaCabecalho.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovNotaFiscalSaidaCabecalho.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovNotaFiscalSaidaCabecalho("Empresa").Value
    Data = rsMovNotaFiscalSaidaCabecalho("Data").Value
    Serie = rsMovNotaFiscalSaidaCabecalho("Serie").Value
    Numero = rsMovNotaFiscalSaidaCabecalho("Numero").Value
    DataEmissao = rsMovNotaFiscalSaidaCabecalho("Data de Emissao").Value
    HoraSaida = rsMovNotaFiscalSaidaCabecalho("Hora de Saida").Value
    CodigoCliente = rsMovNotaFiscalSaidaCabecalho("Codigo de Cliente").Value
    CFOP = rsMovNotaFiscalSaidaCabecalho("CFOP").Value
    Cancelada = rsMovNotaFiscalSaidaCabecalho("Cancelada").Value
    SubEstoque = rsMovNotaFiscalSaidaCabecalho("SubEstoque").Value
    FormaPagamento = rsMovNotaFiscalSaidaCabecalho("Forma de Pagamento").Value
    TotalProduto = rsMovNotaFiscalSaidaCabecalho("Total de Produto").Value
    Desconto = rsMovNotaFiscalSaidaCabecalho("Desconto").Value
    Acrescimo = rsMovNotaFiscalSaidaCabecalho("Acrescimo").Value
    TotalNota = rsMovNotaFiscalSaidaCabecalho("Total da Nota").Value
    BaseCalculoICMS = rsMovNotaFiscalSaidaCabecalho("Base de Calculo de ICMS").Value
    ValorICMS = rsMovNotaFiscalSaidaCabecalho("Valor de ICMS").Value
    BaseCalculoSubstituicao = rsMovNotaFiscalSaidaCabecalho("Base de Calculo de Substituicao").Value
    ValorSubstituicao = rsMovNotaFiscalSaidaCabecalho("Valor de Substituicao").Value
    ValorFrete = rsMovNotaFiscalSaidaCabecalho("Valor de Frete").Value
    ValorSeguro = rsMovNotaFiscalSaidaCabecalho("Valor de Seguro").Value
    ValorIPI = rsMovNotaFiscalSaidaCabecalho("Valor de IPI").Value
    Substituicao = rsMovNotaFiscalSaidaCabecalho("Substituicao").Value
    Outros = rsMovNotaFiscalSaidaCabecalho("Outros").Value
    InformacaoComplementar1 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 1").Value
    InformacaoComplementar2 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 2").Value
    InformacaoComplementar3 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 3").Value
    InformacaoComplementar4 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 4").Value
    InformacaoComplementar5 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 5").Value
    InformacaoComplementar6 = rsMovNotaFiscalSaidaCabecalho("Informacao Complementar 6").Value
    DataDigitacao = rsMovNotaFiscalSaidaCabecalho("Data de Digitacao").Value
    DataAlteracao = rsMovNotaFiscalSaidaCabecalho("Data de Alteracao").Value
    CodigoUsuario = rsMovNotaFiscalSaidaCabecalho("Codigo de Usuario").Value
    Observacao = rsMovNotaFiscalSaidaCabecalho("Observacao").Value
    CodigoClienteTemporario = rsMovNotaFiscalSaidaCabecalho("Codigo de Cliente Temporario").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

