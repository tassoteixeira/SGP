VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCartaFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mTipoMovimento As Integer
Private mNumero As Long
Private mCodigoCliente As Integer
Private mNome As String
Private mVeiculo As String
Private mValorCarta As Currency
Private mValorAbastecimento As Currency
Private mTrocoemDinheiroPista As Currency
Private mTrocoemDinheiroEscritorio As Currency
Private mTrocoemCheque As Currency
Private mNumeroContaCheque As String
Private mNumeroMovimentoCaixaInicial As Long
Private mNumeroMovimentoCaixaFinal As Long

Private rsMovCartaFrete As New adodb.Recordset


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


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
End Property


Public Property Get CodigoCliente() As Integer
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Integer)
    mCodigoCliente = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Veiculo() As String
    Veiculo = mVeiculo
End Property
Public Property Let Veiculo(ByVal Valor As String)
    mVeiculo = Valor
End Property


Public Property Get ValorCarta() As Currency
    ValorCarta = mValorCarta
End Property
Public Property Let ValorCarta(ByVal Valor As Currency)
    mValorCarta = Valor
End Property


Public Property Get ValorAbastecimento() As Currency
    ValorAbastecimento = mValorAbastecimento
End Property
Public Property Let ValorAbastecimento(ByVal Valor As Currency)
    mValorAbastecimento = Valor
End Property


Public Property Get TrocoemDinheiroPista() As Currency
    TrocoemDinheiroPista = mTrocoemDinheiroPista
End Property
Public Property Let TrocoemDinheiroPista(ByVal Valor As Currency)
    mTrocoemDinheiroPista = Valor
End Property


Public Property Get TrocoemDinheiroEscritorio() As Currency
    TrocoemDinheiroEscritorio = mTrocoemDinheiroEscritorio
End Property
Public Property Let TrocoemDinheiroEscritorio(ByVal Valor As Currency)
    mTrocoemDinheiroEscritorio = Valor
End Property


Public Property Get TrocoemCheque() As Currency
    TrocoemCheque = mTrocoemCheque
End Property
Public Property Let TrocoemCheque(ByVal Valor As Currency)
    mTrocoemCheque = Valor
End Property


Public Property Get NumeroContaCheque() As String
    NumeroContaCheque = mNumeroContaCheque
End Property
Public Property Let NumeroContaCheque(ByVal Valor As String)
    mNumeroContaCheque = Valor
End Property


Public Property Get NumeroMovimentoCaixaInicial() As Long
    NumeroMovimentoCaixaInicial = mNumeroMovimentoCaixaInicial
End Property
Public Property Let NumeroMovimentoCaixaInicial(ByVal Valor As Long)
    mNumeroMovimentoCaixaInicial = Valor
End Property


Public Property Get NumeroMovimentoCaixaFinal() As Long
    NumeroMovimentoCaixaFinal = mNumeroMovimentoCaixaFinal
End Property
Public Property Let NumeroMovimentoCaixaFinal(ByVal Valor As Long)
    mNumeroMovimentoCaixaFinal = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & nTipoMovimento
    xCondicao = xCondicao & " AND Numero = " & nNumero
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo < " & mPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] < " & mTipoMovimento
    xCondicao = xCondicao & " AND Numero < " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Tipo do Movimento] DESC, Numero DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo > " & mPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] > " & mTipoMovimento
    xCondicao = xCondicao & " AND Numero > " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Tipo do Movimento] ASC, Numero ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo > " & mPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] > " & mTipoMovimento
    xCondicao = xCondicao & " AND Numero > " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Tipo do Movimento] ASC, Numero ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo = " & mPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & mTipoMovimento
    xCondicao = xCondicao & " AND Numero = " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Tipo do Movimento] DESC, Numero DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoCartaFrete ( Empresa, Data, Periodo, [Tipo do Movimento], Numero, "
    gSQL = gSQL & "[Codigo do Cliente], Nome, Veiculo, [Valor da Carta], [Valor do Abastecimento], "
    gSQL = gSQL & "[Troco em Dinheiro Pista], [Troco em Dinheiro Escritorio], [Troco em Cheque], [Numero da Conta do Cheque], [Numero do Movimento do Caixa Inicial], "
    gSQL = gSQL & "[Numero do Movimento do Caixa Final] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mVeiculo, ", ")
    Call sqlValor(1, mValorCarta, ", ")
    Call sqlValor(1, mValorAbastecimento, ", ")
    Call sqlValor(1, mTrocoemDinheiroPista, ", ")
    Call sqlValor(1, mTrocoemDinheiroEscritorio, ", ")
    Call sqlValor(1, mTrocoemCheque, ", ")
    Call sqlTexto(1, mNumeroContaCheque, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaInicial, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaFinal, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoCartaFrete SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", Numero = ", mNumero)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", Veiculo = ", mVeiculo)
    Call sqlValor(2, ", [Valor da Carta] = ", mValorCarta)
    Call sqlValor(2, ", [Valor do Abastecimento] = ", mValorAbastecimento)
    Call sqlValor(2, ", [Troco em Dinheiro Pista] = ", mTrocoemDinheiroPista)
    Call sqlValor(2, ", [Troco em Dinheiro Escritorio] = ", mTrocoemDinheiroEscritorio)
    Call sqlValor(2, ", [Troco em Cheque] = ", mTrocoemCheque)
    Call sqlTexto(2, ", [Numero da Conta do Cheque] = ", mNumeroContaCheque)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Inicial] = ", mNumeroMovimentoCaixaInicial)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Final] = ", mNumeroMovimentoCaixaFinal)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Tipo do Movimento] = " & nTipoMovimento
    gSQL = gSQL & " AND Numero = " & nNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoCartaFrete"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Tipo do Movimento] = " & nTipoMovimento
    gSQL = gSQL & " AND Numero = " & nNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TotalCartaFrete(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pTipoMovimento As Integer) As Currency

On Error GoTo trata_erro
    TotalCartaFrete = 0
    gSQL = "SELECT SUM([Valor da Carta]) as Total"
    gSQL = gSQL & " FROM MovimentoCartaFrete"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    If Localizar(0) Then
        If Not IsNull(rsMovCartaFrete("Total").Value) Then
            TotalCartaFrete = rsMovCartaFrete("Total").Value
        End If
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalTrocoCheque(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pTipoMovimento As Integer) As Currency

On Error GoTo trata_erro
    TotalTrocoCheque = 0
    gSQL = "SELECT SUM([Troco em Cheque]) as Total"
    gSQL = gSQL & " FROM MovimentoCartaFrete"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    If Localizar(0) Then
        If Not IsNull(rsMovCartaFrete("Total").Value) Then
            TotalTrocoCheque = rsMovCartaFrete("Total").Value
        End If
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalTrocoEscritorio(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pTipoMovimento As Integer) As Currency

On Error GoTo trata_erro
    TotalTrocoEscritorio = 0
    gSQL = "SELECT SUM([Troco em Dinheiro Escritorio]) as Total"
    gSQL = gSQL & " FROM MovimentoCartaFrete"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    If Localizar(0) Then
        If Not IsNull(rsMovCartaFrete("Total").Value) Then
            TotalTrocoEscritorio = rsMovCartaFrete("Total").Value
        End If
    End If
    rsMovCartaFrete.Close
    Set rsMovCartaFrete = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Tipo do Movimento], Numero, [Codigo do Cliente], Nome, Veiculo, [Valor da Carta], [Valor do Abastecimento], [Troco em Dinheiro Pista], [Troco em Dinheiro Escritorio], [Troco em Cheque], [Numero da Conta do Cheque], [Numero do Movimento do Caixa Inicial], [Numero do Movimento do Caixa Final] FROM MovimentoCartaFrete" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCartaFrete = New adodb.Recordset
    rsMovCartaFrete.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCartaFrete.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCartaFrete.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCartaFrete("Empresa").Value
    Data = rsMovCartaFrete("Data").Value
    Periodo = rsMovCartaFrete("Periodo").Value
    TipoMovimento = rsMovCartaFrete("Tipo do Movimento").Value
    Numero = rsMovCartaFrete("Numero").Value
    CodigoCliente = rsMovCartaFrete("Codigo do Cliente").Value
    Nome = rsMovCartaFrete("Nome").Value
    Veiculo = rsMovCartaFrete("Veiculo").Value
    ValorCarta = rsMovCartaFrete("Valor da Carta").Value
    ValorAbastecimento = rsMovCartaFrete("Valor do Abastecimento").Value
    TrocoemDinheiroPista = rsMovCartaFrete("Troco em Dinheiro Pista").Value
    TrocoemDinheiroEscritorio = rsMovCartaFrete("Troco em Dinheiro Escritorio").Value
    TrocoemCheque = rsMovCartaFrete("Troco em Cheque").Value
    NumeroContaCheque = rsMovCartaFrete("Numero da Conta do Cheque").Value
    NumeroMovimentoCaixaInicial = rsMovCartaFrete("Numero do Movimento do Caixa Inicial").Value
    NumeroMovimentoCaixaFinal = rsMovCartaFrete("Numero do Movimento do Caixa Final").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

