VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoFaltaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As String
Private mCodigoFuncionario As Integer
Private mValor As Currency
Private mObservacao As String
Private mNumeroMovimentoCaixa As Long
Private mTipoMovimento As String
Private mValePista As Boolean
Private mNumeroRegistroBaixa As Long
Private mNumeroIlha As Integer
Private mOrdem As Integer

Private rsMovFaltaCaixa As New adodb.Recordset


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


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get TipoMovimento() As String
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As String)
    mTipoMovimento = Valor
End Property


Public Property Get ValePista() As Boolean
    ValePista = mValePista
End Property
Public Property Let ValePista(ByVal Valor As Boolean)
    mValePista = Valor
End Property


Public Property Get NumeroRegistroBaixa() As Long
    NumeroRegistroBaixa = mNumeroRegistroBaixa
End Property
Public Property Let NumeroRegistroBaixa(ByVal Valor As Long)
    mNumeroRegistroBaixa = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(pPeriodo)
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    xCondicao = xCondicao & " AND [Tipo de Movimento] = " & preparaTexto(pTipoMovimento)
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        AtribuiValor
        LocalizarCodigo = True
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
    Dim xSQL As String
    Dim xOrdem As String
    Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND Data <= " & preparaData(mData)
    xOrdem = "ORDER BY Data DESC, Periodo DESC, [Codigo do Funcionario] DESC, [Tipo de Movimento] DESC, Ordem DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(35) Then
        If rsMovFaltaCaixa.RecordCount > 1 Then
            Do Until rsMovFaltaCaixa.EOF
                If xRegIgual = False Then
                    If rsMovFaltaCaixa("Data").Value = mData And rsMovFaltaCaixa("Periodo").Value = mPeriodo And rsMovFaltaCaixa("Codigo do Funcionario").Value = mCodigoFuncionario And rsMovFaltaCaixa("Tipo de Movimento").Value = mTipoMovimento And rsMovFaltaCaixa("Ordem").Value = mOrdem Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovFaltaCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Codigo do Funcionario] ASC, [Tipo de Movimento] ASC, Ordem ASC")
    If Localizar(1) Then
        AtribuiValor
        LocalizarPrimeiro = True
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
    Dim xSQL As String
    Dim xOrdem As String
    Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND Data >= " & preparaData(mData)
    xOrdem = "ORDER BY Data, Periodo, [Codigo do Funcionario], [Tipo de Movimento], Ordem"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(35) Then
        If rsMovFaltaCaixa.RecordCount > 1 Then
            Do Until rsMovFaltaCaixa.EOF
                If xRegIgual = False Then
                    If rsMovFaltaCaixa("Data").Value = mData And rsMovFaltaCaixa("Periodo").Value = mPeriodo And rsMovFaltaCaixa("Codigo do Funcionario").Value = mCodigoFuncionario And rsMovFaltaCaixa("Tipo de Movimento").Value = mTipoMovimento And rsMovFaltaCaixa("Ordem").Value = mOrdem Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovFaltaCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, [Codigo do Funcionario] DESC, [Tipo de Movimento] DESC, Ordem DESC")
    If Localizar(1) Then
        AtribuiValor
        LocalizarUltimo = True
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoFuncionario(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoFuncionario = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Funcionario] DESC, Data DESC, Periodo DESC, [Tipo de Movimento] DESC, Ordem DESC")
    If Localizar(1) Then
        AtribuiValor
        LocalizarUltimoFuncionario = True
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    mOrdem = ProximaOrdemDigitacao(mEmpresa, mCodigoFuncionario, mData, mPeriodo, mTipoMovimento)
    gSQL = "INSERT INTO Movimento_Falta_Caixa ( Empresa, Data, Periodo, [Codigo do Funcionario], Valor, "
    gSQL = gSQL & "Observacao, [Numero do Movimento do Caixa], [Tipo de Movimento], [Vale da Pista], "
    gSQL = gSQL & "[Numero do Registro da Baixa], [Numero da Ilha], Ordem ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mObservacao, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlBoolean(1, mValePista, ", ")
    Call sqlNumero(1, mNumeroRegistroBaixa, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mOrdem, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As String, ByVal pOrdem As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Falta_Caixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlTexto(2, ", [Tipo de Movimento] = ", mTipoMovimento)
    Call sqlBoolean(2, ", [Vale da Pista] = ", mValePista)
    Call sqlNumero(2, ", [Numero do Registro da Baixa] = ", mNumeroRegistroBaixa)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto(pTipoMovimento)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As String, ByVal pOrdem As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto(pTipoMovimento)
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Codigo do Funcionario], Valor, Observacao, [Numero do Movimento do Caixa], [Tipo de Movimento], [Vale da Pista], [Numero do Registro da Baixa], [Numero da Ilha], Ordem FROM Movimento_Falta_Caixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


'Cria o número da próxima ordem da digitação
Private Function ProximaOrdemDigitacao(ByVal pEmpresa As Integer, ByVal pCodigoFuncionario As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pTipoMovimento As String) As Integer

On Error GoTo Err_ProximaOrdemDigitacao

    ProximaOrdemDigitacao = 1
    gSQL = "SELECT Ordem"
    gSQL = gSQL & " FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto(pTipoMovimento)
    gSQL = gSQL & " ORDER BY Ordem DESC"
    If Localizar(1) Then
        ProximaOrdemDigitacao = rsMovFaltaCaixa("Ordem").Value + 1
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

Err_ProximaOrdemDigitacao:
End Function


Public Function TotalFaltaCaixa(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodo As String, ByVal pIlha As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalFaltaCaixa = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    If pPeriodo <> "0" Then
        gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    End If
    If pIlha > 0 Then
        gSQL = gSQL & " AND [Numero da Ilha] = " & pIlha
    End If
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto("F")
    If Localizar(0) Then
        If Not IsNull(rsMovFaltaCaixa("Total").Value) Then
            TotalFaltaCaixa = rsMovFaltaCaixa("Total").Value
        End If
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalFaltaFuncionario(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nData As Date) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalFaltaFuncionario = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto("F")
    gSQL = gSQL & " AND YEAR(Data) = " & Year(nData)
    gSQL = gSQL & " AND MONTH(Data) = " & Month(nData)
    If Localizar(0) Then
        If Not IsNull(rsMovFaltaCaixa("Total").Value) Then
            TotalFaltaFuncionario = rsMovFaltaCaixa("Total").Value
        End If
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalValeFuncionario(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nData As Date) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalValeFuncionario = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto("V")
    gSQL = gSQL & " AND YEAR(Data) = " & Year(nData)
    gSQL = gSQL & " AND MONTH(Data) = " & Month(nData)
    If Localizar(0) Then
        If Not IsNull(rsMovFaltaCaixa("Total").Value) Then
            TotalValeFuncionario = rsMovFaltaCaixa("Total").Value
        End If
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


Public Function TotalValeCaixa(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodo As String, ByVal pIlha As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalValeCaixa = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Falta_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    If pPeriodo <> "0" Then
        gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    End If
    If pIlha > 0 Then
        gSQL = gSQL & " AND [Numero da Ilha] = " & pIlha
    End If
    gSQL = gSQL & " AND [Tipo de Movimento] = " & preparaTexto("V")
    gSQL = gSQL & " AND [Vale da Pista] = " & preparaBooleano(True)
    If Localizar(0) Then
        If Not IsNull(rsMovFaltaCaixa("Total").Value) Then
            TotalValeCaixa = rsMovFaltaCaixa("Total").Value
        End If
    End If
    rsMovFaltaCaixa.Close
    Set rsMovFaltaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovFaltaCaixa = New adodb.Recordset
    rsMovFaltaCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovFaltaCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovFaltaCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovFaltaCaixa("Empresa").Value
    Data = rsMovFaltaCaixa("Data").Value
    Periodo = rsMovFaltaCaixa("Periodo").Value
    CodigoFuncionario = rsMovFaltaCaixa("Codigo do Funcionario").Value
    Valor = rsMovFaltaCaixa("Valor").Value
    Observacao = rsMovFaltaCaixa("Observacao").Value
    NumeroMovimentoCaixa = rsMovFaltaCaixa("Numero do Movimento do Caixa").Value
    TipoMovimento = rsMovFaltaCaixa("Tipo de Movimento").Value
    ValePista = rsMovFaltaCaixa("Vale da Pista").Value
    NumeroRegistroBaixa = rsMovFaltaCaixa("Numero do Registro da Baixa").Value
    NumeroIlha = rsMovFaltaCaixa("Numero da Ilha").Value
    Ordem = rsMovFaltaCaixa("Ordem").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

