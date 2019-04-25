VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoVendaConveniencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroCupom As Long
Private mData As Date
Private mDataCupom As Date
Private mOrdem As Integer
Private mHora As Date
Private mPeriodo As String
Private mTipoMovimento As String
Private mCodigoProduto As Long
Private mValorUnitario As Currency
Private mQuantidade As Currency
Private mValorTotal As Currency
Private mFormaPagamento As Integer
Private mValorRecebido As Currency
Private mCupomCancelado As Boolean
Private mItemCancelado As Boolean
Private mOperador As Integer
Private mCodigoAliquota As Integer
Private mValorDesconto As Currency
Private mNumeroJustificativa As Long
Private mCodigoCliente As Long
Private mCodigoGrupo As Integer
Private mOrigemVenda As String
Private mIlha As Integer
Private mPrecoCusto As Currency
Private mDataEmissaoNFCe As Date
Private mNumeroNFCe As Long
Private mSerieNFCe As String

Private rsMovVendaConveniencia As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get DataCupom() As Date
    DataCupom = mDataCupom
End Property
Public Property Let DataCupom(ByVal Valor As Date)
    mDataCupom = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get Hora() As Date
    Hora = mHora
End Property
Public Property Let Hora(ByVal Valor As Date)
    mHora = Valor
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


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get ValorUnitario() As Currency
    ValorUnitario = mValorUnitario
End Property
Public Property Let ValorUnitario(ByVal Valor As Currency)
    mValorUnitario = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get FormaPagamento() As Integer
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As Integer)
    mFormaPagamento = Valor
End Property


Public Property Get ValorRecebido() As Currency
    ValorRecebido = mValorRecebido
End Property
Public Property Let ValorRecebido(ByVal Valor As Currency)
    mValorRecebido = Valor
End Property


Public Property Get CupomCancelado() As Boolean
    CupomCancelado = mCupomCancelado
End Property
Public Property Let CupomCancelado(ByVal Valor As Boolean)
    mCupomCancelado = Valor
End Property


Public Property Get ItemCancelado() As Boolean
    ItemCancelado = mItemCancelado
End Property
Public Property Let ItemCancelado(ByVal Valor As Boolean)
    mItemCancelado = Valor
End Property


Public Property Get Operador() As Integer
    Operador = mOperador
End Property
Public Property Let Operador(ByVal Valor As Integer)
    mOperador = Valor
End Property


Public Property Get CodigoAliquota() As Integer
    CodigoAliquota = mCodigoAliquota
End Property
Public Property Let CodigoAliquota(ByVal Valor As Integer)
    mCodigoAliquota = Valor
End Property


Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property


Public Property Get NumeroJustificativa() As Long
    NumeroJustificativa = mNumeroJustificativa
End Property
Public Property Let NumeroJustificativa(ByVal Valor As Long)
    mNumeroJustificativa = Valor
End Property


Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property


Public Property Get OrigemVenda() As String
    OrigemVenda = mOrigemVenda
End Property
Public Property Let OrigemVenda(ByVal Valor As String)
    mOrigemVenda = Valor
End Property


Public Property Get Ilha() As Integer
    Ilha = mIlha
End Property
Public Property Let Ilha(ByVal Valor As Integer)
    mIlha = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property

Public Property Get DataEmissaoNFCe() As Date
    DataEmissaoNFCe = mDataEmissaoNFCe
End Property
Public Property Let DataEmissaoNFCe(ByVal Valor As Date)
    mDataEmissaoNFCe = Valor
End Property


Public Property Get NumeroNFCe() As Long
    NumeroNFCe = mNumeroNFCe
End Property
Public Property Let NumeroNFCe(ByVal Valor As Long)
    mNumeroNFCe = Valor
End Property


Public Property Get SerieNFCe() As String
    SerieNFCe = mSerieNFCe
End Property
Public Property Let SerieNFCe(ByVal Valor As String)
    mSerieNFCe = Valor
End Property

'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Ilha = " & pIlha
    xCondicao = xCondicao & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] < " & mNumeroCupom
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem < " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Numero do Cupom] DESC, Data DESC, Ordem DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroData(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Ilha = " & pIlha
    xCondicao = xCondicao & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNumeroData = Localizar(1)
    If LocalizarNumeroData Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroProximaOrdem(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroProximaOrdem = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Ilha = " & pIlha
    xCondicao = xCondicao & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    xCondicao = xCondicao & " AND Ordem > " & pOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ordem ASC")
    If Localizar(1) Then
        LocalizarNumeroProximaOrdem = True
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] > " & mNumeroCupom
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem > " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Numero do Cupom] ASC, Data ASC, Ordem ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Cupom] > " & mNumeroCupom
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem > " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Numero do Cupom] ASC, Data ASC, Ordem ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pIlha As Integer, ByVal pOrigemVenda As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Ilha = " & pIlha
    xCondicao = xCondicao & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ilha DESC, [Origem da Venda] DESC, [Numero do Cupom] DESC, Data DESC, Ordem DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Venda_Conveniencia ( Empresa, [Numero do Cupom], Data, [Data do Cupom], Ordem, "
    gSQL = gSQL & "Hora, Periodo, [Tipo do Movimento], [Codigo do Produto], [Valor Unitario], "
    gSQL = gSQL & "Quantidade, [Valor Total], [Forma de Pagamento], [Valor Recebido], [Cupom Cancelado], "
    gSQL = gSQL & "[Item Cancelado], Operador, [Codigo da Aliquota], [Valor do Desconto], "
    gSQL = gSQL & "[Numero da Justificativa], [Codigo do Cliente], [Codigo do Grupo], "
    gSQL = gSQL & "[Origem da Venda], Ilha, [Preco de Custo], [Data Emissao NFCe], [Numero NFCe],[Serie NFCe] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlData(1, mDataCupom, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlNumero(1, mFormaPagamento, ", ")
    Call sqlValor(1, mValorRecebido, ", ")
    Call sqlBoolean(1, mCupomCancelado, ", ")
    Call sqlBoolean(1, mItemCancelado, ", ")
    Call sqlNumero(1, mOperador, ", ")
    Call sqlNumero(1, mCodigoAliquota, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlNumero(1, mNumeroJustificativa, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlNumero(1, mCodigoGrupo, ", ")
    Call sqlTexto(1, mOrigemVenda, ", ")
    Call sqlNumero(1, mIlha, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlData(1, mDataEmissaoNFCe, ", ")
    Call sqlNumero(1, mNumeroNFCe, ", ")
    Call sqlTexto(1, mSerieNFCe, " )")

    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlData(2, ", Data = ", mData)
    Call sqlData(2, ", [Data do Cupom] = ", mDataCupom)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlNumero(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor Recebido] = ", mValorRecebido)
    Call sqlBoolean(2, ", [Cupom Cancelado] = ", mCupomCancelado)
    Call sqlBoolean(2, ", [Item Cancelado] = ", mItemCancelado)
    Call sqlNumero(2, ", Operador = ", mOperador)
    Call sqlNumero(2, ", [Codigo da Aliquota] = ", mCodigoAliquota)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlNumero(2, ", [Numero da Justificativa] = ", mNumeroJustificativa)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    Call sqlNumero(2, ", [Codigo do Grupo] = ", mCodigoGrupo)
    Call sqlTexto(2, ", [Origem da Venda] = ", mOrigemVenda)
    Call sqlNumero(2, ", Ilha = ", mIlha)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlData(2, ", [Data Emissao NFCe] = ", mDataEmissaoNFCe)
    Call sqlNumero(2, ", [Numero NFCe] = ", mNumeroNFCe)
    Call sqlTexto(2, ", [Serie NFCe] = ", mSerieNFCe)
   
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarFormaPagamento(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarFormaPagamento = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    
    Call sqlNumero(2, "[Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor Recebido] = ", mValorRecebido)
    Call sqlNumero(2, ", Operador = ", mOperador)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarFormaPagamento = True
    End If
    Exit Function

Err_Alterar:
End Function
Public Function AlterarDadosNFCe(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarDadosNFCe = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    
    Call sqlData(2, " [Data Emissao NFCe] = ", mDataEmissaoNFCe)
    Call sqlNumero(2, ", [Numero NFCe] = ", mNumeroNFCe)
    Call sqlTexto(2, ", [Serie NFCe] = ", mSerieNFCe)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarDadosNFCe = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function GravaDesconto(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pValorDesconto As Currency) As Boolean
Dim xRecordsAffected As Long
Dim xQtdItem As Integer
Dim xTotalBruto As Currency
Dim xPercentual As Currency
Dim xDescontoAcumulado As Currency
Dim xDesconto As Currency
Dim xOrdem As Integer

On Error GoTo Err_Alterar

    xQtdItem = 0
    xTotalBruto = 0
    'Verifica quantos ítens tem no cupom a gravar desconto
    gSQL = "SELECT COUNT(1) AS QtdItens, SUM([Valor Total]) AS TotalBruto FROM Movimento_Venda_Conveniencia"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND [Item Cancelado] = " & preparaBooleano(False)
    If Localizar(1) Then
        If Not IsNull(rsMovVendaConveniencia("QtdItens").Value) Then
            xQtdItem = rsMovVendaConveniencia("QtdItens").Value
            xTotalBruto = rsMovVendaConveniencia("TotalBruto").Value
        End If
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    
    
    'Se quantidade de iten for igual a 0 sai da funcao
    If xQtdItem = 0 Then
        Exit Function
    'Se quantidade de iten for igual a 1 altera o cupom lancando o desconto
    ElseIf xQtdItem = 1 Then
        gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
        Call sqlValor(2, "[Valor do Desconto] = ", pValorDesconto)
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
        gSQL = gSQL & " AND Data = " & preparaData(pData)
        gSQL = gSQL & " AND Ilha = " & pIlha
        gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
        gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
        If xRecordsAffected > 0 Then
            GravaDesconto = True
        End If
        Exit Function
    End If
    
    xPercentual = pValorDesconto * 100 / xTotalBruto
    xDescontoAcumulado = 0
    xOrdem = 0
    'Loop nos ítens do cupom.
    gSQL = "SELECT Ordem, [Valor Total] FROM Movimento_Venda_Conveniencia"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND [Item Cancelado] = " & preparaBooleano(False)
    gSQL = gSQL & "ORDER BY Ordem ASC"
    If Localizar(0) Then
        Do Until rsMovVendaConveniencia.EOF
            xOrdem = xOrdem + 1
            If xOrdem = xQtdItem Then
                xDesconto = pValorDesconto - xDescontoAcumulado
            Else
                xDesconto = Format(rsMovVendaConveniencia("Valor Total").Value * xPercentual / 100, "0000000000.00")
                xDescontoAcumulado = xDescontoAcumulado + xDesconto
            End If
            gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
            Call sqlValor(2, "[Valor do Desconto] = ", xDesconto)
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
            gSQL = gSQL & " AND Data = " & preparaData(pData)
            gSQL = gSQL & " AND Ilha = " & pIlha
            gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
            gSQL = gSQL & " AND Ordem = " & rsMovVendaConveniencia("Ordem").Value
            gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
            If xRecordsAffected > 0 Then
                GravaDesconto = True
            Else
                GravaDesconto = False
                Exit Do
            End If
            rsMovVendaConveniencia.MoveNext
        Loop
    End If
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    
    
    Exit Function

Err_Alterar:
End Function


Public Function CancelaCupom(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    CancelaCupom = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    Call sqlBoolean(2, "[Cupom Cancelado] = ", "True")
    Call sqlBoolean(2, ", [Item Cancelado] = ", "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaCupom = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function CancelaItemCupom(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    CancelaItemCupom = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    Call sqlBoolean(2, "[Item Cancelado] = ", "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaItemCupom = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function CancelaItemCupomJustif(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer, ByVal pNumeroJustificativa As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    CancelaItemCupomJustif = False
    gSQL = "UPDATE Movimento_Venda_Conveniencia SET "
    Call sqlBoolean(2, "[Item Cancelado] = ", "True")
    Call sqlNumero(2, ", [Numero da Justificativa] = ", pNumeroJustificativa)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaItemCupomJustif = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pIlha As Integer, ByVal pOrigemVenda As String, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Venda_Conveniencia"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ilha = " & pIlha
    gSQL = gSQL & " AND [Origem da Venda] = " & preparaTexto(pOrigemVenda)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Cupom], Data, [Data do Cupom], Ordem, Hora, Periodo, [Tipo do Movimento], [Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Forma de Pagamento], [Valor Recebido], [Cupom Cancelado], [Item Cancelado], Operador, [Codigo da Aliquota], [Valor do Desconto], [Numero da Justificativa], [Codigo do Cliente], [Codigo do Grupo], [Origem da Venda], Ilha, [Preco de Custo], [Data Emissao NFCe], [Numero NFCe], [Serie NFCe] FROM Movimento_Venda_Conveniencia" & " " & pCondicao & " " & pTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovVendaConveniencia = New adodb.Recordset
    rsMovVendaConveniencia.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovVendaConveniencia.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovVendaConveniencia.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovVendaConveniencia("Empresa").Value
    NumeroCupom = rsMovVendaConveniencia("Numero do Cupom").Value
    Data = rsMovVendaConveniencia("Data").Value
    DataCupom = rsMovVendaConveniencia("Data do Cupom").Value
    Ordem = rsMovVendaConveniencia("Ordem").Value
    Hora = rsMovVendaConveniencia("Hora").Value
    Periodo = rsMovVendaConveniencia("Periodo").Value
    TipoMovimento = rsMovVendaConveniencia("Tipo do Movimento").Value
    CodigoProduto = rsMovVendaConveniencia("Codigo do Produto").Value
    ValorUnitario = rsMovVendaConveniencia("Valor Unitario").Value
    Quantidade = rsMovVendaConveniencia("Quantidade").Value
    ValorTotal = rsMovVendaConveniencia("Valor Total").Value
    FormaPagamento = rsMovVendaConveniencia("Forma de Pagamento").Value
    ValorRecebido = rsMovVendaConveniencia("Valor Recebido").Value
    CupomCancelado = rsMovVendaConveniencia("Cupom Cancelado").Value
    ItemCancelado = rsMovVendaConveniencia("Item Cancelado").Value
    Operador = rsMovVendaConveniencia("Operador").Value
    CodigoAliquota = rsMovVendaConveniencia("Codigo da Aliquota").Value
    ValorDesconto = rsMovVendaConveniencia("Valor do Desconto").Value
    NumeroJustificativa = rsMovVendaConveniencia("Numero da Justificativa").Value
    CodigoCliente = rsMovVendaConveniencia("Codigo do Cliente").Value
    CodigoGrupo = rsMovVendaConveniencia("Codigo do Grupo").Value
    OrigemVenda = rsMovVendaConveniencia("Origem da Venda").Value
    Ilha = rsMovVendaConveniencia("Ilha").Value
    PrecoCusto = rsMovVendaConveniencia("Preco de Custo").Value
    If IsNull(rsMovVendaConveniencia("Data Emissao NFCe").Value) Then
        DataEmissaoNFCe = CDate("00:00:00")
    Else
        DataEmissaoNFCe = rsMovVendaConveniencia("Data Emissao NFCe").Value
    End If
    NumeroNFCe = rsMovVendaConveniencia("Numero NFCe").Value
    SerieNFCe = rsMovVendaConveniencia("Serie NFCe").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

