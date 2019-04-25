VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoLubrificante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As String
Private mCodigoFuncionario As Integer
Private mCodigoProduto As Long
Private mQuantidade As Currency
Private mValorCusto As Currency
Private mValorVenda As Currency
Private mValorTotal As Currency
Private mOrdemDigitacao As Integer
Private mNumeroIlha As Integer
Private mCodigoTipoSubEstoque As Integer
Private mTipoMovimento As Integer

Private rsMovLubrificante As New adodb.Recordset


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


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorCusto() As Currency
    ValorCusto = mValorCusto
End Property
Public Property Let ValorCusto(ByVal Valor As Currency)
    mValorCusto = Valor
End Property


Public Property Get ValorVenda() As Currency
    ValorVenda = mValorVenda
End Property
Public Property Let ValorVenda(ByVal Valor As Currency)
    mValorVenda = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get OrdemDigitacao() As Integer
    OrdemDigitacao = mOrdemDigitacao
End Property
Public Property Let OrdemDigitacao(ByVal Valor As Integer)
    mOrdemDigitacao = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get CodigoTipoSubEstoque() As Integer
    CodigoTipoSubEstoque = mCodigoTipoSubEstoque
End Property
Public Property Let CodigoTipoSubEstoque(ByVal Valor As Integer)
    mCodigoTipoSubEstoque = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoTipoSubEstoque As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(pPeriodo)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    xCondicao = xCondicao & " AND [Codigo do Produto2] = " & pCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Tipo do Movimento] DESC, [Codigo do Tipo do SubEstoque] DESC, [Codigo do Produto2] DESC, [Codigo do Funcionario] DESC")
    If Localizar(100) Then
        If rsMovLubrificante.RecordCount > 1 Then
            Do Until rsMovLubrificante.EOF
                If xRegIgual = False Then
                    If rsMovLubrificante("Data").Value = mData And rsMovLubrificante("Periodo").Value = mPeriodo And rsMovLubrificante("Numero da Ilha").Value = mNumeroIlha And rsMovLubrificante("Tipo do Movimento").Value = mTipoMovimento And rsMovLubrificante("Codigo do Tipo do SubEstoque").Value = mCodigoTipoSubEstoque And rsMovLubrificante("Codigo do Produto2").Value = mCodigoProduto And rsMovLubrificante("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovLubrificante.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Tipo do Movimento] ASC, [Codigo do Tipo do SubEstoque] ASC, [Codigo do Produto2] ASC, [Codigo do Funcionario] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroDataPer(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoTipoSubEstoque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroDataPer = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(pPeriodo)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto2] ASC, [Codigo do Funcionario] ASC")
    LocalizarPrimeiroDataPer = Localizar(1)
    If LocalizarPrimeiroDataPer = True Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Tipo do Movimento] ASC, [Codigo do Tipo do SubEstoque] ASC, [Codigo do Produto2] ASC, [Codigo do Funcionario] ASC")
    If Localizar(100) Then
        If rsMovLubrificante.RecordCount > 1 Then
            Do Until rsMovLubrificante.EOF
                If xRegIgual = False Then
                    If rsMovLubrificante("Data").Value = mData And rsMovLubrificante("Periodo").Value = mPeriodo And rsMovLubrificante("Numero da Ilha").Value = mNumeroIlha And rsMovLubrificante("Tipo do Movimento").Value = mTipoMovimento And rsMovLubrificante("Codigo do Tipo do SubEstoque").Value = mCodigoTipoSubEstoque And rsMovLubrificante("Codigo do Produto2").Value = mCodigoProduto And rsMovLubrificante("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovLubrificante.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Tipo do Movimento] DESC, [Codigo do Tipo do SubEstoque] DESC, [Codigo do Produto2] DESC, [Codigo do Funcionario] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mOrdemDigitacao = ProximaOrdemDigitacao(mEmpresa, mData, mPeriodo, mNumeroIlha, mTipoMovimento, mCodigoTipoSubEstoque)
    gSQL = "INSERT INTO Movimento_Lubrificante ( Empresa, Data, Periodo, [Codigo do Funcionario], [Codigo do Produto2], "
    gSQL = gSQL & "Quantidade, [Valor Custo], [Valor Venda], [Valor Total], [Ordem da Digitacao], "
    gSQL = gSQL & "[Numero da Ilha], [Codigo do Tipo do SubEstoque], [Tipo do Movimento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorCusto, ", ")
    Call sqlValor(1, mValorVenda, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlNumero(1, mOrdemDigitacao, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mCodigoTipoSubEstoque, ", ")
    Call sqlNumero(1, mTipoMovimento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoTipoSubEstoque As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Lubrificante SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Custo] = ", mValorCusto)
    Call sqlValor(2, ", [Valor Venda] = ", mValorVenda)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlNumero(2, ", [Ordem da Digitacao] = ", mOrdemDigitacao)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Codigo do Tipo do SubEstoque] = ", mCodigoTipoSubEstoque)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlteraCusto(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoProduto As Long, ByVal pPrecoCusto As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlteraCusto = False
    gSQL = "UPDATE Movimento_Lubrificante SET "
    Call sqlValor(2, "[Valor Custo] = ", pPrecoCusto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlteraCusto = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoTipoSubEstoque As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Lubrificante"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Codigo do Funcionario], [Codigo do Produto2], Quantidade, [Valor Custo], [Valor Venda], [Valor Total], [Ordem da Digitacao], [Numero da Ilha], [Codigo do Tipo do SubEstoque], [Tipo do Movimento] FROM Movimento_Lubrificante" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovLubrificante = New adodb.Recordset
    rsMovLubrificante.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovLubrificante.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovLubrificante.EOF Then
        AtualizaRecordset = True
    End If
End Function


Public Function MondaRS(ByVal pSQL As String) As adodb.Recordset
    Dim xRS As adodb.Recordset
    
    Set xRS = New adodb.Recordset
    xRS.CursorLocation = adUseClient
    xRS.Open pSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set MondaRS = xRS
End Function


Public Function TotalPeriodo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pTipoMovimento As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalPeriodo = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Lubrificante"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    If Localizar(0) Then
        If Not IsNull(rsMovLubrificante("Total").Value) Then
            TotalPeriodo = rsMovLubrificante("Total").Value
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodoFuncionario(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pTipoMovimento As Integer, ByVal pCodigoFuncionario As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalPeriodoFuncionario = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Lubrificante"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    If Localizar(0) Then
        If Not IsNull(rsMovLubrificante("Total").Value) Then
            TotalPeriodoFuncionario = rsMovLubrificante("Total").Value
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalQtd(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoProduto As Long) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalQtd = 0
    gSQL = "SELECT SUM(Quantidade) as Total FROM Movimento_Lubrificante"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    If Localizar(0) Then
        If Not IsNull(rsMovLubrificante("Total").Value) Then
            TotalQtd = rsMovLubrificante("Total").Value
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe

'Cria o número da próxima ordem da digitação
Private Function ProximaOrdemDigitacao(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoTipoSubEstoque As Integer) As Long

On Error GoTo Err_ProximaOrdemDigitacao

    ProximaOrdemDigitacao = 1
    gSQL = "SELECT [Ordem da Digitacao]"
    gSQL = gSQL & " FROM Movimento_Lubrificante"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = gSQL & " ORDER BY [Ordem da Digitacao] DESC"
    If Localizar(1) Then
        ProximaOrdemDigitacao = rsMovLubrificante("Ordem da Digitacao").Value + 1
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    Exit Function

Err_ProximaOrdemDigitacao:
End Function


Private Sub AtribuiValor()
    Empresa = rsMovLubrificante("Empresa").Value
    Data = rsMovLubrificante("Data").Value
    Periodo = rsMovLubrificante("Periodo").Value
    CodigoFuncionario = rsMovLubrificante("Codigo do Funcionario").Value
    CodigoProduto = rsMovLubrificante("Codigo do Produto2").Value
    Quantidade = rsMovLubrificante("Quantidade").Value
    ValorCusto = rsMovLubrificante("Valor Custo").Value
    ValorVenda = rsMovLubrificante("Valor Venda").Value
    ValorTotal = rsMovLubrificante("Valor Total").Value
    OrdemDigitacao = rsMovLubrificante("Ordem da Digitacao").Value
    NumeroIlha = rsMovLubrificante("Numero da Ilha").Value
    CodigoTipoSubEstoque = rsMovLubrificante("Codigo do Tipo do SubEstoque").Value
    TipoMovimento = rsMovLubrificante("Tipo do Movimento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

