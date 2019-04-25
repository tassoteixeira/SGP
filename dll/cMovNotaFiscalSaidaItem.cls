VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovNotaFiscalSaidaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mSerie As String
Private mNumero As Long
Private mOrdem As Integer
Private mCodigoProduto As Long
Private mCodigoAliquota As Integer
Private mPrecoCusto As Currency
Private mPrecoVenda As Currency
Private mQuantidade As Currency
Private mTotal As Currency
Private mSubEstoque As Integer
Private mTipoCombustivel As String
Private mCancelado As Boolean
Private mCFOP As String

Private rsMovNotaFiscalSaidaItem As New adodb.Recordset


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


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get CodigoAliquota() As Integer
    CodigoAliquota = mCodigoAliquota
End Property
Public Property Let CodigoAliquota(ByVal Valor As Integer)
    mCodigoAliquota = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get Total() As Currency
    Total = mTotal
End Property
Public Property Let Total(ByVal Valor As Currency)
    mTotal = Valor
End Property


Public Property Get SubEstoque() As Integer
    SubEstoque = mSubEstoque
End Property
Public Property Let SubEstoque(ByVal Valor As Integer)
    mSubEstoque = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Cancelado() As Boolean
    Cancelado = mCancelado
End Property
Public Property Let Cancelado(ByVal Valor As Boolean)
    mCancelado = Valor
End Property


Public Property Get CFOP() As String
    CFOP = mCFOP
End Property
Public Property Let CFOP(ByVal Valor As String)
    mCFOP = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & "   AND Data = " & preparaData(pData)
    xCondicao = xCondicao & "   AND Serie = " & preparaTexto(pSerie)
    xCondicao = xCondicao & "   AND Numero = " & pNumero
    xCondicao = xCondicao & "   AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovNotaFiscalSaidaItem.Close
    Set rsMovNotaFiscalSaidaItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoNotaFiscalSaidaItem ( Empresa, Data, Serie, Numero, Ordem, "
    gSQL = gSQL & "[Codigo de Produto], [Codigo de Aliquota], [Preco de Custo], [Preco de Venda], Quantidade, "
    gSQL = gSQL & "Total, SubEstoque, [Tipo de Combustivel], Cancelado, CFOP ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlNumero(1, mCodigoAliquota, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mTotal, ", ")
    Call sqlNumero(1, mSubEstoque, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlBoolean(1, mCancelado, ", ")
    Call sqlTexto(1, mCFOP, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoNotaFiscalSaidaItem SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlNumero(2, ", Numero = ", mNumero)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlNumero(2, ", [Codigo de Produto] = ", mCodigoProduto)
    Call sqlNumero(2, ", [Codigo de Aliquota] = ", mCodigoAliquota)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", Total = ", mTotal)
    Call sqlNumero(2, ", SubEstoque = ", mSubEstoque)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlBoolean(2, ", Cancelado = ", mCancelado)
    Call sqlTexto(2, ", CFOP = ", mCFOP)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Serie = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero = " & pNumero
    gSQL = gSQL & "   AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pSerie As String, ByVal pNumero As Long, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoNotaFiscalSaidaItem"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Serie = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero = " & pNumero
    gSQL = gSQL & "   AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function QuantidadeCombustivelVendaData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pSoma5929e6929 As Boolean, ByVal pSoma5664e6664 As Boolean) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    QuantidadeCombustivelVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM MovimentoNotaFiscalSaidaItem"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Cancelado = " & preparaBooleano(False)
    If pSoma5929e6929 = False Then
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("5929")
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("6929")
    End If
    If pSoma5664e6664 = False Then
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("5664")
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("6664")
    End If
    'Não soma cfop 1??? e 2???
    xSQL = xSQL & "    AND CFOP NOT LIKE " & preparaTexto("1%")
    xSQL = xSQL & "    AND CFOP NOT LIKE " & preparaTexto("2%")
    Set rsMovNotaFiscalSaidaItem = New adodb.Recordset
    rsMovNotaFiscalSaidaItem.CursorLocation = adUseClient
    rsMovNotaFiscalSaidaItem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovNotaFiscalSaidaItem.EOF Then
        If Not IsNull(rsMovNotaFiscalSaidaItem("QuantidadeVenda").Value) Then
            QuantidadeCombustivelVendaData = rsMovNotaFiscalSaidaItem("QuantidadeVenda").Value
        End If
    End If
    rsMovNotaFiscalSaidaItem.Close
    Set rsMovNotaFiscalSaidaItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorCombustivelVendaData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pSoma5929e6929 As Boolean, ByVal pSoma5664e6664 As Boolean) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorCombustivelVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Total) AS TotalVenda"
    xSQL = xSQL & "   FROM MovimentoNotaFiscalSaidaItem"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Cancelado = " & preparaBooleano(False)
    If pSoma5929e6929 = False Then
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("5929")
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("6929")
    End If
    If pSoma5664e6664 = False Then
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("5664")
        xSQL = xSQL & "    AND CFOP <> " & preparaTexto("6664")
    End If
    'Não soma cfop 1??? e 2???
    xSQL = xSQL & "    AND CFOP NOT LIKE " & preparaTexto("1%")
    xSQL = xSQL & "    AND CFOP NOT LIKE " & preparaTexto("2%")
    
    Set rsMovNotaFiscalSaidaItem = New adodb.Recordset
    rsMovNotaFiscalSaidaItem.CursorLocation = adUseClient
    rsMovNotaFiscalSaidaItem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovNotaFiscalSaidaItem.EOF Then
        If Not IsNull(rsMovNotaFiscalSaidaItem("TotalVenda").Value) Then
            ValorCombustivelVendaData = rsMovNotaFiscalSaidaItem("TotalVenda").Value
        End If
    End If
    rsMovNotaFiscalSaidaItem.Close
    Set rsMovNotaFiscalSaidaItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Serie, Numero, Ordem, [Codigo de Produto], [Codigo de Aliquota], [Preco de Custo], [Preco de Venda], Quantidade, Total, [Tipo de Combustivel], SubEstoque, Cancelado, CFOP FROM MovimentoNotaFiscalSaidaItem" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovNotaFiscalSaidaItem = New adodb.Recordset
    rsMovNotaFiscalSaidaItem.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovNotaFiscalSaidaItem.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovNotaFiscalSaidaItem.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovNotaFiscalSaidaItem("Empresa").Value
    Data = rsMovNotaFiscalSaidaItem("Data").Value
    Serie = rsMovNotaFiscalSaidaItem("Serie").Value
    Numero = rsMovNotaFiscalSaidaItem("Numero").Value
    Ordem = rsMovNotaFiscalSaidaItem("Ordem").Value
    CodigoProduto = rsMovNotaFiscalSaidaItem("Codigo de Produto").Value
    CodigoAliquota = rsMovNotaFiscalSaidaItem("Codigo de Aliquota").Value
    PrecoCusto = rsMovNotaFiscalSaidaItem("Preco de Custo").Value
    PrecoVenda = rsMovNotaFiscalSaidaItem("Preco de Venda").Value
    Quantidade = rsMovNotaFiscalSaidaItem("Quantidade").Value
    Total = rsMovNotaFiscalSaidaItem("Total").Value
    SubEstoque = rsMovNotaFiscalSaidaItem("SubEstoque").Value
    TipoCombustivel = rsMovNotaFiscalSaidaItem("Tipo de Combustivel").Value
    Cancelado = rsMovNotaFiscalSaidaItem("Cancelado").Value
    CFOP = rsMovNotaFiscalSaidaItem("CFOP").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

