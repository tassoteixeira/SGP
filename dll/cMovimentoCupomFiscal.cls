VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCupomFiscal"
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
Private mCodigoCliente As Long
Private mCodigoConveniado As Long
Private mCodigoProduto As Long
Private mValorUnitario As Currency
Private mQuantidade As Currency
Private mValorTotal As Currency
Private mFormaPagamento As Integer
Private mValorRecebido As Currency
Private mNumeroCheque As String
Private mTelefone As String
Private mCupomCancelado As Boolean
Private mItemCancelado As Boolean
Private mOperador As Integer
Private mCodigoAliquota As Integer
Private mValorDesconto As Currency
Private mNome As String
Private mCPFCNPJ As String
Private mTipoCombustivel As String
Private mCodigoECF As Integer
Private mCodigoGrupo As Integer
Private mTipoSubEstoque As Integer
Private mValorDescontoEmbutido As Currency

Private rsMovCupomFiscal As New adodb.Recordset
Private IntegracaoNuvem As New cIntegracaoNuvem


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


Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get CodigoConveniado() As Long
    CodigoConveniado = mCodigoConveniado
End Property
Public Property Let CodigoConveniado(ByVal Valor As Long)
    mCodigoConveniado = Valor
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


Public Property Get NumeroCheque() As String
    NumeroCheque = mNumeroCheque
End Property
Public Property Let NumeroCheque(ByVal Valor As String)
    mNumeroCheque = Valor
End Property


Public Property Get Telefone() As String
    Telefone = mTelefone
End Property
Public Property Let Telefone(ByVal Valor As String)
    mTelefone = Valor
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


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get CPFCNPJ() As String
    CPFCNPJ = mCPFCNPJ
End Property
Public Property Let CPFCNPJ(ByVal Valor As String)
    mCPFCNPJ = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get CodigoECF() As Integer
    CodigoECF = mCodigoECF
End Property
Public Property Let CodigoECF(ByVal Valor As Integer)
    mCodigoECF = Valor
End Property


Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property


Public Property Get TipoSubEstoque() As Integer
    TipoSubEstoque = mTipoSubEstoque
End Property
Public Property Let TipoSubEstoque(ByVal Valor As Integer)
    mTipoSubEstoque = Valor
End Property


Public Property Get ValorDescontoEmbutido() As Currency
    ValorDescontoEmbutido = mValorDescontoEmbutido
End Property
Public Property Let ValorDescontoEmbutido(ByVal Valor As Currency)
    mValorDescontoEmbutido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    If pCodigoECF > 0 Then
        xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    End If
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroProximaOrdem(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroProximaOrdem = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Ordem > " & pOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ordem ASC")
    If Localizar(1) Then
        LocalizarNumeroProximaOrdem = True
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProdutoAnteriorData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoProduto As Integer, ByVal pCupomComplementar As Boolean) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarProdutoAnteriorData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    xCondicao = xCondicao & " AND Data < " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Valor Unitario] DESC")
    If Localizar(1) Then
        LocalizarProdutoAnteriorData = True
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "cMovimentoBomba:LocalizarCombAnteriorData"
    MsgBox gSQL 'ApenasTeste
End Function


Public Function LocalizarProximaOrdemDeste(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Integer
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximaOrdemDeste = 1
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ordem DESC")
    If Localizar(1) Then
        LocalizarProximaOrdemDeste = rsMovCupomFiscal("Ordem").Value + 1
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroData(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNumeroData = Localizar(1)
    If LocalizarNumeroData Then
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
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
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Numero do Cupom] ASC, Ordem ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroData(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    If pCodigoECF > 0 Then
        xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    End If
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cupom] ASC, Ordem ASC")
    If Localizar(1) Then
        LocalizarPrimeiroData = True
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
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
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Numero do Cupom] DESC, Ordem DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoData(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    If pCodigoECF > 0 Then
        xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    End If
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cupom] DESC, Ordem DESC")
    If Localizar(1) Then
        LocalizarUltimoData = True
        AtribuiValor
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Cupom_Fiscal ( Empresa, [Numero do Cupom], Data, [Data do Cupom], Ordem, "
    gSQL = gSQL & "Hora, Periodo, [Tipo do Movimento], [Codigo do Cliente], [Codigo do Conveniado], "
    gSQL = gSQL & "[Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Forma de Pagamento], "
    gSQL = gSQL & "[Valor Recebido], [Numero do Cheque], Telefone, [Cupom Cancelado], [Item Cancelado], "
    gSQL = gSQL & "Operador, [Codigo da Aliquota], [Valor do Desconto], Nome, [CPF CNPJ], "
    gSQL = gSQL & "[Tipo de Combustivel], [Codigo da Ecf], [Codigo do Grupo], [Tipo do SubEstoque], "
    gSQL = gSQL & "[Valor Desconto Embutido] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlData(1, mDataCupom, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlNumero(1, mCodigoConveniado, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlNumero(1, mFormaPagamento, ", ")
    Call sqlValor(1, mValorRecebido, ", ")
    Call sqlTexto(1, mNumeroCheque, ", ")
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlBoolean(1, mCupomCancelado, ", ")
    Call sqlBoolean(1, mItemCancelado, ", ")
    Call sqlNumero(1, mOperador, ", ")
    Call sqlNumero(1, mCodigoAliquota, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mCPFCNPJ, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlNumero(1, mCodigoECF, ", ")
    Call sqlNumero(1, mCodigoGrupo, ", ")
    Call sqlNumero(1, mTipoSubEstoque, ", ")
    Call sqlValor(1, mValorDescontoEmbutido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
        'Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mCodigoECF, mNumeroCupom, mData, mOrdem)
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlData(2, ", Data = ", mData)
    Call sqlData(2, ", [Data do Cupom] = ", mDataCupom)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    Call sqlNumero(2, ", [Codigo do Conveniado] = ", mCodigoConveniado)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlNumero(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor Recebido] = ", mValorRecebido)
    Call sqlTexto(2, ", [Numero do Cheque] = ", mNumeroCheque)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlBoolean(2, ", [Cupom Cancelado] = ", mCupomCancelado)
    Call sqlBoolean(2, ", [Item Cancelado] = ", mItemCancelado)
    Call sqlNumero(2, ", Operador = ", mOperador)
    Call sqlNumero(2, ", [Codigo da Aliquota] = ", mCodigoAliquota)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", [CPF CNPJ] = ", mCPFCNPJ)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlNumero(2, ", [Codigo da Ecf] = ", mCodigoECF)
    Call sqlNumero(2, ", [Codigo do Grupo] = ", mCodigoGrupo)
    Call sqlNumero(2, ", [Tipo do SubEstoque] = ", mTipoSubEstoque)
    Call sqlValor(2, ", [Valor Desconto Embutido] = ", mValorDescontoEmbutido)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
        'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarDesconto(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pTotalCupom As Currency, ByVal pValorDesconto As Currency) As Boolean
Dim lRecordsAffected As Long
Dim xSomaDesconto As Currency
Dim xPercentual As Currency
Dim xValor As Currency
Dim xOrdem As Integer

On Error GoTo Err_Alterar

    AlterarDesconto = False
    xSomaDesconto = 0
    xPercentual = Format(pValorDesconto * 100 / pTotalCupom, "0000000000.0000")
    gSQL = "SELECT [Valor do Desconto], [Valor Total], Ordem"
    gSQL = gSQL & "  FROM Movimento_Cupom_Fiscal"
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & "   AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    gSQL = gSQL & " ORDER BY Ordem"
    If Localizar(55) Then
        If rsMovCupomFiscal.RecordCount > 1 Then
            Do Until rsMovCupomFiscal.EOF
                xOrdem = rsMovCupomFiscal("Ordem").Value
                xValor = Format(rsMovCupomFiscal("Valor Total").Value * xPercentual / 100, "0000000000.00")
                xSomaDesconto = xSomaDesconto + xValor
                gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
                Call sqlValor(2, "[Valor do Desconto] = ", xValor)
                gSQL = gSQL & " WHERE Empresa = " & pEmpresa
                gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
                gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
                gSQL = gSQL & " AND Data = " & preparaData(pData)
                gSQL = gSQL & " AND Ordem = " & rsMovCupomFiscal("Ordem").Value
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                    'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, rsMovCupomFiscal("Ordem").Value)
                End If
                rsMovCupomFiscal.MoveNext
            Loop
            If pValorDesconto <> xSomaDesconto Then
                If pValorDesconto < xSomaDesconto Then
                    xValor = xValor - (xSomaDesconto - pValorDesconto)
                Else
                    xValor = xValor + (pValorDesconto - xSomaDesconto)
                End If
                gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
                Call sqlValor(2, "[Valor do Desconto] = ", xValor)
                gSQL = gSQL & " WHERE Empresa = " & pEmpresa
                gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
                gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
                gSQL = gSQL & " AND Data = " & preparaData(pData)
                gSQL = gSQL & " AND Ordem = " & xOrdem
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                End If
            End If
        Else
            gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
            Call sqlValor(2, "[Valor do Desconto] = ", pValorDesconto)
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
            gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
            gSQL = gSQL & " AND Data = " & preparaData(pData)
            gSQL = gSQL & " AND Ordem = " & rsMovCupomFiscal("Ordem").Value
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected > 0 Then
                AlterarDesconto = True
            End If
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

Err_Alterar:
End Function


Public Function AlterarFormaPagamento(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarFormaPagamento = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
    
    Call sqlNumero(2, "[Codigo do Cliente] = ", mCodigoCliente)
    Call sqlNumero(2, ", [Codigo do Conveniado] = ", mCodigoConveniado)
    Call sqlNumero(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor Recebido] = ", mValorRecebido)
    Call sqlTexto(2, ", [Numero do Cheque] = ", mNumeroCheque)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlNumero(2, ", Operador = ", mOperador)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", [CPF CNPJ] = ", mCPFCNPJ)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarFormaPagamento = True
        'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, 0)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function CancelaCupom(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    CancelaCupom = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
    Call sqlBoolean(2, "[Cupom Cancelado] = ", "True")
    Call sqlBoolean(2, ", [Item Cancelado] = ", "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaCupom = True
        'Call IncluiIntegracaoNuvem("CANCELAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, 0)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function CancelaItemCupom(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    CancelaItemCupom = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal SET "
    Call sqlBoolean(2, "[Item Cancelado] = ", "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaItemCupom = True
        'Call IncluiIntegracaoNuvem("CANCELAR ITEM", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function NumeroECFnaData(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pPrimeiro As Boolean) As Long
Dim xSQL As String

On Error GoTo trata_erro

    NumeroECFnaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 1 [Numero do Cupom]"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    xSQL = xSQL & "    AND Data = " & preparaData(pData)
    xSQL = xSQL & "  ORDER BY [Numero do Cupom]"
    If pPrimeiro = True Then
        xSQL = xSQL & " ASC"
    Else
        xSQL = xSQL & " DESC"
    End If
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("Numero do Cupom").Value) Then
            NumeroECFnaData = rsMovCupomFiscal("Numero do Cupom").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function QuantidadeProdutoVendaData(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pFormaPagamento As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    QuantidadeProdutoVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Codigo do Produto] = " & pCodigoProduto
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    If pFormaPagamento > 0 Then
        xSQL = xSQL & "    AND [Forma de Pagamento] = " & pFormaPagamento
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("QuantidadeVenda").Value) Then
            QuantidadeProdutoVendaData = rsMovCupomFiscal("QuantidadeVenda").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function QuantidadeCombustivelVendaData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    QuantidadeCombustivelVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("QuantidadeVenda").Value) Then
            QuantidadeCombustivelVendaData = rsMovCupomFiscal("QuantidadeVenda").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorAliquotaVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoAliquota As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorAliquotaVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    If pCodigoAliquota > 0 Then
        xSQL = xSQL & "    AND [Codigo da Aliquota]  = " & pCodigoAliquota
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorAliquotaVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            ValorAliquotaVendaData = ValorAliquotaVendaData - rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorProdutoVendaData(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pFormaPagamento As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorProdutoVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    If pCodigoProduto > 0 Then
        xSQL = xSQL & "    AND [Codigo do Produto] = " & pCodigoProduto
    Else
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto("  ")
    End If
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    If pFormaPagamento > 0 Then
        xSQL = xSQL & "    AND [Forma de Pagamento] = " & pFormaPagamento
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorProdutoVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorProdutoGasVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoECF As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorProdutoGasVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal, Produto"
    xSQL = xSQL & "  WHERE Movimento_Cupom_Fiscal.Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Tipo de Combustivel] = " & preparaTexto("  ")
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"
    xSQL = xSQL & "    AND Produto.[Codigo do Grupo] = 3"
    If bdAccess Then
        xSQL = xSQL & "    AND Produto.Nome LIKE " & preparaTexto("*GAS*")
        xSQL = xSQL & "    AND Produto.Nome LIKE " & preparaTexto("*13*")
    ElseIf bdSqlServer Then
        xSQL = xSQL & "    AND Produto.Nome LIKE " & preparaTexto("%GAS%")
        xSQL = xSQL & "    AND Produto.Nome LIKE " & preparaTexto("%13%")
    End If
    If pCodigoECF > 0 Then
        xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Codigo da ECF] = " & pCodigoECF
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorProdutoGasVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            ValorProdutoGasVendaData = ValorProdutoGasVendaData - rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorProdutosSubstVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorProdutosSubstVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal, Aliquota"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto("  ")
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal.[Codigo da Aliquota] = Aliquota.Codigo"
    xSQL = xSQL & "    AND Aliquota.[Codigo Fiscal] = " & preparaTexto("FF")
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorProdutosSubstVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            ValorProdutosSubstVendaData = ValorProdutosSubstVendaData - rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorCombustivelVendaData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorCombustivelVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorCombustivelVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorCombustiveisVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoECF As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorCombustiveisVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] > " & preparaTexto("  ")
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    If pCodigoECF > 0 Then
        xSQL = xSQL & "    AND [Codigo da ECF] = " & pCodigoECF
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalVenda").Value) Then
            ValorCombustiveisVendaData = rsMovCupomFiscal("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            ValorCombustiveisVendaData = ValorCombustiveisVendaData - rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function DescontoProdutoVendaData(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pFormaPagamento As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    DescontoProdutoVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    If pCodigoProduto > 0 Then
        xSQL = xSQL & "    AND [Codigo do Produto] = " & pCodigoProduto
    End If
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    'xSQL = xSQL & "    AND Ordem = " & 1
    If pFormaPagamento > 0 Then
        xSQL = xSQL & "    AND [Forma de Pagamento] = " & pFormaPagamento
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            DescontoProdutoVendaData = rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function CancelamentoProdutoVendaData(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pFormaPagamento As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    CancelamentoProdutoVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS Total"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    If pCodigoProduto > 0 Then
        xSQL = xSQL & "    AND [Codigo do Produto] = " & pCodigoProduto
    End If
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(True)
    If pFormaPagamento > 0 Then
        xSQL = xSQL & "    AND [Forma de Pagamento] = " & pFormaPagamento
    End If
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("Total").Value) Then
            CancelamentoProdutoVendaData = rsMovCupomFiscal("Total").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function DescontoCombustivelVendaData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    DescontoCombustivelVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & preparaTexto(pPeriodoInicial)
    xSQL = xSQL & "    AND Periodo <= " & preparaTexto(pPeriodoFinal)
    xSQL = xSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    'xSQL = xSQL & "    AND Ordem = " & 1
    
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    rsMovCupomFiscal.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        If Not IsNull(rsMovCupomFiscal("TotalDesconto").Value) Then
            DescontoCombustivelVendaData = rsMovCupomFiscal("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscal.Close
    Set rsMovCupomFiscal = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Cupom_Fiscal"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
        'Call IncluiIntegracaoNuvem("EXCLUIR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_Excluir:
End Function
Public Function ExcluirNumeroData(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirNumeroData = False
    gSQL = "DELETE FROM Movimento_Cupom_Fiscal"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirNumeroData = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Cupom], Data, [Data do Cupom], Ordem, Hora, Periodo, [Tipo do Movimento], [Codigo do Cliente], [Codigo do Conveniado], [Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Forma de Pagamento], [Valor Recebido], [Numero do Cheque], Telefone, [Cupom Cancelado], [Item Cancelado], Operador, [Codigo da Aliquota], [Valor do Desconto], Nome, [CPF CNPJ], [Tipo de Combustivel], [Codigo da Ecf], [Codigo do Grupo], [Tipo do SubEstoque], [Valor Desconto Embutido] FROM Movimento_Cupom_Fiscal" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCupomFiscal = New adodb.Recordset
    rsMovCupomFiscal.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCupomFiscal.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscal.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCupomFiscal("Empresa").Value
    NumeroCupom = rsMovCupomFiscal("Numero do Cupom").Value
    Data = rsMovCupomFiscal("Data").Value
    DataCupom = rsMovCupomFiscal("Data do Cupom").Value
    Ordem = rsMovCupomFiscal("Ordem").Value
    Hora = rsMovCupomFiscal("Hora").Value
    Periodo = rsMovCupomFiscal("Periodo").Value
    TipoMovimento = rsMovCupomFiscal("Tipo do Movimento").Value
    CodigoCliente = rsMovCupomFiscal("Codigo do Cliente").Value
    CodigoConveniado = rsMovCupomFiscal("Codigo do Conveniado").Value
    CodigoProduto = rsMovCupomFiscal("Codigo do Produto").Value
    ValorUnitario = rsMovCupomFiscal("Valor Unitario").Value
    Quantidade = rsMovCupomFiscal("Quantidade").Value
    ValorTotal = rsMovCupomFiscal("Valor Total").Value
    FormaPagamento = rsMovCupomFiscal("Forma de Pagamento").Value
    ValorRecebido = rsMovCupomFiscal("Valor Recebido").Value
    NumeroCheque = rsMovCupomFiscal("Numero do Cheque").Value
    Telefone = rsMovCupomFiscal("Telefone").Value
    CupomCancelado = rsMovCupomFiscal("Cupom Cancelado").Value
    ItemCancelado = rsMovCupomFiscal("Item Cancelado").Value
    Operador = rsMovCupomFiscal("Operador").Value
    CodigoAliquota = rsMovCupomFiscal("Codigo da Aliquota").Value
    ValorDesconto = rsMovCupomFiscal("Valor do Desconto").Value
    Nome = rsMovCupomFiscal("Nome").Value
    CPFCNPJ = rsMovCupomFiscal("CPF CNPJ").Value
    TipoCombustivel = rsMovCupomFiscal("Tipo de Combustivel").Value
    CodigoECF = rsMovCupomFiscal("Codigo da Ecf").Value
    CodigoGrupo = rsMovCupomFiscal("Codigo do Grupo").Value
    TipoSubEstoque = rsMovCupomFiscal("Tipo do SubEstoque").Value
    If IsNull(rsMovCupomFiscal("Valor Desconto Embutido").Value) Then
        ValorDescontoEmbutido = 0
    Else
        ValorDescontoEmbutido = rsMovCupomFiscal("Valor Desconto Embutido").Value
    End If
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Private Sub IncluiIntegracaoNuvem(ByVal pTipoOperacao As String, ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer)
    IntegracaoNuvem.Empresa = mEmpresa
    IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
    IntegracaoNuvem.NomeTabela = "MovimentoCupomFiscal"
    IntegracaoNuvem.ChaveAcesso = pEmpresa & "|@|" & pCodigoECF & "|@|" & pNumeroCupom & "|@|" & pData & "|@|" & pOrdem & "|@|"
    IntegracaoNuvem.TipoOperacao = pTipoOperacao
    IntegracaoNuvem.IntegradoEm = CDate("00:00:00")
    If Not IntegracaoNuvem.Incluir Then
        Call CriaLogCadastroDll2("cMovimentoCupomFiscal:IncluiIntegracaoNuvem - Erro não identificado.", Err.Description, "Tabela=" & IntegracaoNuvem.NomeTabela & " - Chv.Acesso=" & IntegracaoNuvem.ChaveAcesso)
    End If
End Sub

