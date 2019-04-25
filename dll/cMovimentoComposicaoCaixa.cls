VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoComposicaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mTipoMovimento As Integer
Private mCodigoFuncionario As Integer
Private mCodigoComposicao As Integer
Private mValor As Currency
Private mNumeroMovimentoCaixa As Long

Private rsMovComposicaoCaixa As New adodb.Recordset


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


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get CodigoComposicao() As Integer
    CodigoComposicao = mCodigoComposicao
End Property
Public Property Let CodigoComposicao(ByVal Valor As Integer)
    mCodigoComposicao = Valor
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
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroIlha As Integer, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nCodigoComposicao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & nNumeroIlha
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & nTipoMovimento
    xCondicao = xCondicao & " AND [Codigo da Composicao] = " & nCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    'xCondicao = xCondicao & " AND [Numero da Ilha] <= " & mNumeroIlha
    'xCondicao = xCondicao & " AND Periodo <= " & mPeriodo
    'xCondicao = xCondicao & " AND [Tipo do Movimento] <= " & mTipoMovimento
    'xCondicao = xCondicao & " AND [Codigo da Composicao] <= " & mCodigoComposicao
    xOrdem = "ORDER BY Empresa DESC, Data DESC, [Numero da Ilha] DESC, Periodo DESC, [Tipo do Movimento] DESC, [Codigo da Composicao] DESC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(55) Then
        If rsMovComposicaoCaixa.RecordCount > 1 Then
            Do Until rsMovComposicaoCaixa.EOF
                If xRegIgual = False Then
                    If rsMovComposicaoCaixa("Data").Value = mData And rsMovComposicaoCaixa("Numero da Ilha").Value = mNumeroIlha And rsMovComposicaoCaixa("Periodo").Value = mPeriodo And rsMovComposicaoCaixa("Tipo do Movimento").Value = mTipoMovimento Then
                        xRegIgual = True
                    End If
                Else
                    If rsMovComposicaoCaixa("Data").Value <> mData Or rsMovComposicaoCaixa("Numero da Ilha").Value <> mNumeroIlha Or rsMovComposicaoCaixa("Periodo").Value <> mPeriodo Or rsMovComposicaoCaixa("Tipo do Movimento").Value <> mTipoMovimento Then
                        LocalizarAnterior = True
                        Exit Do
                    End If
                End If
                rsMovComposicaoCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
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
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, [Numero da Ilha] ASC, Periodo ASC, [Tipo do Movimento] ASC, [Codigo da Composicao] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroData(ByVal pEmpresa As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero da Ilha] ASC, Periodo ASC, [Tipo do Movimento] ASC, [Codigo da Composicao] ASC")
    If Localizar(1) Then
        LocalizarPrimeiroData = True
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(mData)
    'xCondicao = xCondicao & " AND [Numero da Ilha] >= " & mNumeroIlha
    'xCondicao = xCondicao & " AND Periodo >= " & mPeriodo
    'xCondicao = xCondicao & " AND [Tipo do Movimento] >= " & mTipoMovimento
    'xCondicao = xCondicao & " AND [Codigo da Composicao] >= " & mCodigoComposicao
    xOrdem = "ORDER BY Empresa ASC, Data ASC, [Numero da Ilha] ASC, Periodo ASC, [Tipo do Movimento] ASC, [Codigo da Composicao] ASC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(55) Then
        If rsMovComposicaoCaixa.RecordCount > 1 Then
            Do Until rsMovComposicaoCaixa.EOF
                If xRegIgual = False Then
                    If rsMovComposicaoCaixa("Data").Value = mData And rsMovComposicaoCaixa("Numero da Ilha").Value = mNumeroIlha And rsMovComposicaoCaixa("Periodo").Value = mPeriodo And rsMovComposicaoCaixa("Tipo do Movimento").Value = mTipoMovimento Then
                        xRegIgual = True
                    End If
                Else
                    If rsMovComposicaoCaixa("Data").Value <> mData Or rsMovComposicaoCaixa("Numero da Ilha").Value <> mNumeroIlha Or rsMovComposicaoCaixa("Periodo").Value <> mPeriodo Or rsMovComposicaoCaixa("Tipo do Movimento").Value <> mTipoMovimento Then
                        LocalizarProximo = True
                        Exit Do
                    End If
                End If
                rsMovComposicaoCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, [Numero da Ilha] DESC, Periodo DESC, [Tipo do Movimento] DESC, [Codigo da Composicao] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Composicao_Caixa ( Empresa, Data, Periodo, [Numero da Ilha], [Tipo do Movimento], "
    gSQL = gSQL & "[Codigo do Funcionario], [Codigo da Composicao], Valor, [Numero do Movimento do Caixa] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mCodigoComposicao, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroIlha As Integer, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nCodigoComposicao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Composicao_Caixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", [Codigo da Composicao] = ", mCodigoComposicao)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero da Ilha] = " & nNumeroIlha
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroIlha As Integer, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer, ByVal nCodigoComposicao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Composicao_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero da Ilha] = " & nNumeroIlha
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Tipo do Movimento] = " & nTipoMovimento
    gSQL = gSQL & " AND [Codigo da Composicao] = " & nCodigoComposicao
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluirRegistros(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroIlha As Integer, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirRegistros = False
    gSQL = "DELETE FROM Movimento_Composicao_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero da Ilha] = " & nNumeroIlha
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Tipo do Movimento] = " & nTipoMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirRegistros = True
    End If
    Exit Function

Err_Excluir:
End Function


'Inicio Métodos da Classe
Public Function ExisteRegistroData(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroIlha As Integer, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteRegistroData = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & nNumeroIlha
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & nTipoMovimento
    gSQL = PreparaSQL(xCondicao, "")
    ExisteRegistroData = Localizar(1)
    'If ExisteRegistroData Then
    '    AtribuiValor
    'End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Numero da Ilha], [Tipo do Movimento], [Codigo do Funcionario], [Codigo da Composicao], Valor, [Numero do Movimento do Caixa] FROM Movimento_Composicao_Caixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovComposicaoCaixa = New adodb.Recordset
    rsMovComposicaoCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovComposicaoCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovComposicaoCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovComposicaoCaixa("Empresa").Value
    Data = rsMovComposicaoCaixa("Data").Value
    Periodo = rsMovComposicaoCaixa("Periodo").Value
    NumeroIlha = rsMovComposicaoCaixa("Numero da Ilha").Value
    TipoMovimento = rsMovComposicaoCaixa("Tipo do Movimento").Value
    CodigoFuncionario = rsMovComposicaoCaixa("Codigo do Funcionario").Value
    CodigoComposicao = rsMovComposicaoCaixa("Codigo da Composicao").Value
    Valor = rsMovComposicaoCaixa("Valor").Value
    NumeroMovimentoCaixa = rsMovComposicaoCaixa("Numero do Movimento do Caixa").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

