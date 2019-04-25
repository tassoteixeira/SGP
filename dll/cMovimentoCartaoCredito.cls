VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCartaoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataEmissao As Date
Private mPeriodo As String
Private mTipoMovimento As String
Private mNumeroLancamento As Integer
Private mCodigoCartao As Integer
Private mDataVencimento As Date
Private mValor As Currency
Private mNumeroCartao As Integer
Private mNome As String
Private mNumeroMovimentoCaixa As Long
Private mTaxaAdministrativa As Currency
Private mNumeroIlha As Integer
Private mAutorizacao As String
Private mNSU As String
Private mCodigoFuncionario As Integer

Private rsMovCartaoCredito As New adodb.Recordset
Private IntegracaoCaixa As cIntegracaoCaixa
Private MovCaixa As cMovimentoCaixa


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
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


Public Property Get NumeroCartao() As Integer
    NumeroCartao = mNumeroCartao
End Property
Public Property Let NumeroCartao(ByVal Valor As Integer)
    mNumeroCartao = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
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


'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nNumeroLancamento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(nPeriodo)
    xCondicao = xCondicao & " AND [Numero do Lancamento] = " & nNumeroLancamento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xRegIgual As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] <= " & preparaData(mDataEmissao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data de Emissao] DESC, Periodo DESC, [Numero do Lancamento] DESC")
    If Localizar(200) Then
        If rsMovCartaoCredito.RecordCount > 1 Then
            Do Until rsMovCartaoCredito.EOF
                If xRegIgual = False Then
                    If rsMovCartaoCredito("Data de Emissao").Value = mDataEmissao And rsMovCartaoCredito("Periodo").Value = mPeriodo And rsMovCartaoCredito("Numero do Lancamento").Value = mNumeroLancamento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovCartaoCredito.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
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
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarDataPeriodoNome(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String, ByVal pTipoMovimento As String, ByVal pNome As String) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarDataPeriodoNome = False
    
    xCondicao = ""
    xCondicao = xCondicao & " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(pPeriodo)
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & preparaTexto(pTipoMovimento)
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarDataPeriodoNome = Localizar(1)
    If LocalizarDataPeriodoNome Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data de Emissao] ASC, Periodo ASC, [Numero do Lancamento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xRegIgual As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] >= " & preparaData(mDataEmissao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data de Emissao] ASC, Periodo ASC, [Numero do Lancamento] ASC")
    If Localizar(200) Then
        If rsMovCartaoCredito.RecordCount > 1 Then
            Do Until rsMovCartaoCredito.EOF
                If xRegIgual = False Then
                    If rsMovCartaoCredito("Data de Emissao").Value = mDataEmissao And rsMovCartaoCredito("Periodo").Value = mPeriodo And rsMovCartaoCredito("Numero do Lancamento").Value = mNumeroLancamento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovCartaoCredito.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data de Emissao] DESC, Periodo DESC, [Numero do Lancamento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    'mNumeroLancamento = ProximoRegistro(mEmpresa, mDataEmissao, mPeriodo)
    mNumeroLancamento = ProximoRegistro(mEmpresa, mDataEmissao)
    gSQL = "INSERT INTO Movimento_Cartao_Credito ( Empresa, [Data de Emissao], Periodo, "
    gSQL = gSQL & "[Tipo do Movimento], [Numero do Lancamento], [Codigo do Cartao], [Data do Vencimento], "
    gSQL = gSQL & "Valor, [Numero do Cartao], Nome, [Numero do Movimento do Caixa], [Taxa Administrativa], "
    gSQL = gSQL & "[Numero da Ilha], Autorizacao, NSU, [Codigo do Funcionario] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mNumeroLancamento, ", ")
    Call sqlNumero(1, mCodigoCartao, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlNumero(1, mNumeroCartao, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlValor(1, mTaxaAdministrativa, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlTexto(1, mAutorizacao, ", ")
    Call sqlTexto(1, mNSU, ", ")
    Call sqlNumero(1, mCodigoFuncionario, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nNumeroLancamento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cartao_Credito SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Numero do Lancamento] = ", mNumeroLancamento)
    Call sqlNumero(2, ", [Codigo do Cartao] = ", mCodigoCartao)
    Call sqlData(2, ", [Data do Vencimento] = ", mDataVencimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlNumero(2, ", [Numero do Cartao] = ", mNumeroCartao)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlValor(2, ", [Taxa Administrativa] = ", mTaxaAdministrativa)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlTexto(2, ", Autorizacao = ", mAutorizacao)
    Call sqlTexto(2, ", NSU = ", mNSU)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Numero do Lancamento] = " & nNumeroLancamento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nNumeroLancamento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Numero do Lancamento] = " & nNumeroLancamento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], [Numero do Lancamento], [Codigo do Cartao], [Data do Vencimento], Valor, [Numero do Cartao], Nome, [Numero do Movimento do Caixa], [Taxa Administrativa], [Numero da Ilha], Autorizacao, NSU, [Codigo do Funcionario] FROM Movimento_Cartao_Credito" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCartaoCredito = New adodb.Recordset
    rsMovCartaoCredito.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCartaoCredito.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCartaoCredito.EOF Then
        AtualizaRecordset = True
    End If
End Function


Public Function BaixaCartao(ByVal nEmpresa As Integer, ByVal pCodigoUsuario As Integer, ByVal nDataEmissao As Boolean, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nCodigoCartao As Integer, ByVal nDataAntecipacao As Date, ByVal nTaxaAdministrativa As Currency) As Boolean
    Dim xRecordsAffected As Long
    Dim xCondicao As String
    Dim xTaxaAdministrativa As Currency
    Dim xValorAdministrativo As Currency
    Dim xQtdDias As Integer
    Dim xTaxaProporcional As Currency

On Error GoTo trata_erro
    
    BaixaCartao = False
    gConn.BeginTrans
    
    
    gSQL = ""
    gSQL = gSQL & "SELECT Empresa, [Data de Emissao], Periodo, [Tipo do Movimento],"
    gSQL = gSQL & "       [Numero do Lancamento], [Codigo do Cartao], [Data do Vencimento],"
    gSQL = gSQL & "       Valor, [Numero do Cartao], Movimento_Cartao_Credito.Nome, [Numero do Movimento do Caixa],"
    gSQL = gSQL & "       [Taxa Administrativa], [Numero da Ilha], Cartao_Credito.Nome as NomeCartao,"
    gSQL = gSQL & "       Autorizacao, NSU, Movimento_Cartao_Credito.[Codigo do Funcionario]"
    gSQL = gSQL & "  FROM Movimento_Cartao_Credito, Cartao_Credito"
    xCondicao = ""
    xCondicao = xCondicao & " WHERE Empresa = " & nEmpresa
    If nDataEmissao Then
        xCondicao = xCondicao & " AND [Data de Emissao] >= " & preparaData(nDataInicial)
        xCondicao = xCondicao & " AND [Data de Emissao] <= " & preparaData(nDataFinal)
    Else
        xCondicao = xCondicao & " AND [Data do Vencimento] >= " & preparaData(nDataInicial)
        xCondicao = xCondicao & " AND [Data do Vencimento] <= " & preparaData(nDataFinal)
    End If
    If nCodigoCartao > 0 Then
        xCondicao = xCondicao & " AND [Codigo do Cartao] = " & nCodigoCartao
    End If
    gSQL = gSQL & xCondicao
    gSQL = gSQL & " AND Movimento_Cartao_Credito.[Codigo do Cartao] = Cartao_Credito.Codigo"
    If Localizar(0) Then
        With rsMovCartaoCredito
            Do Until .EOF
                If IncluiMovimentoCaixa(pCodigoUsuario) Then
                    xTaxaAdministrativa = ![Taxa Administrativa]
                    If nDataAntecipacao <> "00:00:00" Then
                        xQtdDias = DateDiff("d", nDataAntecipacao, ![Data do Vencimento])
                        xTaxaProporcional = xTaxaAdministrativa + (nTaxaAdministrativa / 30 * xQtdDias)
                        xValorAdministrativo = Format(!Valor * xTaxaAdministrativa / 100, "00000000.00")
                        xValorAdministrativo = xValorAdministrativo + Format(!Valor * xTaxaProporcional / 100, "00000000.00")
                    Else
                        xValorAdministrativo = Format(!Valor * xTaxaAdministrativa / 100, "00000000.00")
                    End If
                    gSQL = "INSERT INTO Baixa_Cartao_Credito ( Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], [Numero do Lancamento], "
                    gSQL = gSQL & "[Codigo do Cartao], [Data do Vencimento], Valor, [Numero do Cartao], Nome, "
                    gSQL = gSQL & "[Data da Baixa], [Taxa Administrativa], [Valor Administrativo], [Numero do Movimento do Caixa], "
                    gSQL = gSQL & "[Numero do Movimento do Caixa Baixa], [Numero da Ilha], Autorizacao, NSU, [Codigo do Funcionario] ) VALUES ( "
                    Call sqlNumero(1, !Empresa, ", ")
                    Call sqlData(1, ![Data de Emissao], ", ")
                    Call sqlTexto(1, !Periodo, ", ")
                    Call sqlTexto(1, ![Tipo do Movimento], ", ")
                    Call sqlNumero(1, ![Numero do Lancamento], ", ")
                    Call sqlNumero(1, ![Codigo do Cartao], ", ")
                    Call sqlData(1, ![Data do Vencimento], ", ")
                    Call sqlValor(1, !Valor, ", ")
                    Call sqlNumero(1, ![Numero do Cartao], ", ")
                    Call sqlTexto(1, !Nome, ", ")
                    If nDataAntecipacao <> "00:00:00" Then
                        Call sqlData(1, nDataAntecipacao, ", ")
                    Else
                        Call sqlData(1, ![Data do Vencimento], ", ")
                    End If
                    Call sqlValor(1, xTaxaAdministrativa, ", ")
                    Call sqlValor(1, xValorAdministrativo, ", ")
                    Call sqlNumero(1, ![Numero do Movimento do Caixa], ", ")
                    Call sqlNumero(1, MovCaixa.NumeroMovimento, ", ")
                    Call sqlNumero(1, ![Numero da Ilha], ", ")
                    Call sqlTexto(1, !Autorizacao, ", ")
                    Call sqlTexto(1, !NSU, ", ")
                    Call sqlNumero(1, ![Codigo do Funcionario], " )")
                    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
                    If xRecordsAffected = 0 Then
                        rsMovCartaoCredito.Close
                        Set rsMovCartaoCredito = Nothing
                        gConn.RollbackTrans
                        Exit Function
                    End If
                Else
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                    rsMovCartaoCredito.Close
                    Set rsMovCartaoCredito = Nothing
                    gConn.RollbackTrans
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    
    'DELETA Movimento_Cartao_Credito
    gSQL = "DELETE FROM Movimento_Cartao_Credito " & xCondicao
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected = 0 Then
        gConn.RollbackTrans
        Exit Function
    End If
    gConn.CommitTrans
    BaixaCartao = True
    Exit Function

trata_erro:
    gConn.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function EstornaCartao(ByVal nEmpresa As Integer, ByVal nDataEmissao As Boolean, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nCodigoCartao As Integer) As Boolean
    Dim xRecordsAffected As Long
    Dim xCondicao As String

On Error GoTo trata_erro
    
    EstornaCartao = False
    gConn.BeginTrans
    gSQL = ""
    gSQL = gSQL & "SELECT Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], [Numero do Lancamento],"
    gSQL = gSQL & "       [Codigo do Cartao], [Data do Vencimento], Valor, [Numero do Cartao], Baixa_Cartao_Credito.Nome,"
    gSQL = gSQL & "       [Data da Baixa], [Taxa Administrativa], [Valor Administrativo],"
    gSQL = gSQL & "       [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa],"
    gSQL = gSQL & "       [Numero da Ilha], Cartao_Credito.[Taxa de Custo], Autorizacao, NSU, Baixa_Cartao_Credito.[Codigo do Funcionario]"
    gSQL = gSQL & " FROM Baixa_Cartao_Credito, Cartao_Credito"
    xCondicao = ""
    xCondicao = xCondicao & " WHERE Empresa = " & nEmpresa
    If nDataEmissao Then
        xCondicao = xCondicao & " AND [Data de Emissao] >= " & preparaData(nDataInicial)
        xCondicao = xCondicao & " AND [Data de Emissao] <= " & preparaData(nDataFinal)
    Else
        xCondicao = xCondicao & " AND [Data do Vencimento] >= " & preparaData(nDataInicial)
        xCondicao = xCondicao & " AND [Data do Vencimento] <= " & preparaData(nDataFinal)
    End If
    If nCodigoCartao > 0 Then
        xCondicao = xCondicao & " AND [Codigo do Cartao] = " & nCodigoCartao
    End If
    gSQL = gSQL & xCondicao
    gSQL = gSQL & " AND Baixa_Cartao_Credito.[Codigo do Cartao] = Cartao_Credito.Codigo"
    If Localizar(0) Then
        With rsMovCartaoCredito
            Do Until .EOF
                If ExcluiMovimentoCaixa(!Empresa, ![Data do Vencimento], ![Numero do Movimento do Caixa]) Then
                Else
'                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
'                    rsMovCartaoCredito.Close
'                    Set rsMovCartaoCredito = Nothing
'                    gConn.RollbackTrans
'                    Exit Function
                End If
                gSQL = "INSERT INTO Movimento_Cartao_Credito ( Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], [Numero do Lancamento], "
                gSQL = gSQL & "[Codigo do Cartao], [Data do Vencimento], Valor, [Numero do Cartao], Nome, "
                gSQL = gSQL & "[Numero do Movimento do Caixa], [Taxa Administrativa], [Numero da Ilha], "
                gSQL = gSQL & "Autorizacao, NSU, [Codigo do Funcionario] ) VALUES ( "
                Call sqlNumero(1, !Empresa, ", ")
                Call sqlData(1, ![Data de Emissao], ", ")
                Call sqlTexto(1, !Periodo, ", ")
                Call sqlTexto(1, ![Tipo do Movimento], ", ")
                Call sqlNumero(1, ![Numero do Lancamento], ", ")
                Call sqlNumero(1, ![Codigo do Cartao], ", ")
                Call sqlData(1, ![Data do Vencimento], ", ")
                Call sqlValor(1, !Valor, ", ")
                Call sqlNumero(1, ![Numero do Cartao], ", ")
                Call sqlTexto(1, !Nome, ", ")
                Call sqlNumero(1, ![Numero do Movimento do Caixa], ", ")
                Call sqlValor(1, ![Taxa Administrativa], ", ")
                Call sqlNumero(1, ![Numero da Ilha], ", ")
                Call sqlTexto(1, !Autorizacao, ", ")
                Call sqlTexto(1, !NSU, ", ")
                Call sqlNumero(1, ![Codigo do Funcionario], " )")
                gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
                If xRecordsAffected = 0 Then
                    rsMovCartaoCredito.Close
                    Set rsMovCartaoCredito = Nothing
                    gConn.RollbackTrans
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    
    'DELETA Baixa_Cartao_Credito
    gSQL = "DELETE FROM Baixa_Cartao_Credito " & xCondicao
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected = 0 Then
        gConn.RollbackTrans
        Exit Function
    End If
    gConn.CommitTrans
    EstornaCartao = True
    Exit Function

trata_erro:
    gConn.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function IncluiMovimentoCaixa(ByVal pCodigoUsuario As Integer) As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    
    With rsMovCartaoCredito
        If IntegracaoCaixa.LocalizarNome(!Empresa, "BAIXA CARTAO " & !NomeCartao) Then
            xComplemento = "De " & Format(![Data de Emissao], "dd/mm/yyyy") & " TM:" & ![Tipo do Movimento] & " P:" & !Periodo & " P/ " & Format(![Data do Vencimento], "dd/mm/yyyy")
            MovCaixa.Empresa = !Empresa
            MovCaixa.Data = ![Data do Vencimento]
            MovCaixa.NumeroMovimento = 1
            MovCaixa.Valor = !Valor
            MovCaixa.NumeroDocumento = ![Numero do Cartao]
            MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixa.Complemento = Mid(xComplemento, 1, 40)
            MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
            MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
            MovCaixa.TipoMovimento = 2
            MovCaixa.FluxoCaixa = True
            MovCaixa.CodigoUsuario = pCodigoUsuario
            If MovCaixa.Incluir > 0 Then
                IncluiMovimentoCaixa = True
            Else
                MsgBox "Não foi integrado no caixa o valor=" & !Valor, vbInformation, "Erro de Integridade"
            End If
        Else
            MsgBox "Não existe a integração=" & "CARTAO " & !NomeCartao & ".", vbInformation, "Registro Inexistente"
        End If
    End With
End Function
Private Function ExcluiMovimentoCaixa(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimentoCaixa As Long) As Boolean
    ExcluiMovimentoCaixa = False
    If pNumeroMovimentoCaixa = 0 Then
        ExcluiMovimentoCaixa = True
        Exit Function
    Else
        If MovCaixa.Excluir(pEmpresa, pData, pNumeroMovimentoCaixa) Then
            ExcluiMovimentoCaixa = True
        Else
            'MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCartaoCredito("Empresa").Value
    DataEmissao = rsMovCartaoCredito("Data de Emissao").Value
    Periodo = rsMovCartaoCredito("Periodo").Value
    TipoMovimento = rsMovCartaoCredito("Tipo do Movimento").Value
    NumeroLancamento = rsMovCartaoCredito("Numero do Lancamento").Value
    CodigoCartao = rsMovCartaoCredito("Codigo do Cartao").Value
    DataVencimento = rsMovCartaoCredito("Data do Vencimento").Value
    Valor = rsMovCartaoCredito("Valor").Value
    NumeroCartao = rsMovCartaoCredito("Numero do Cartao").Value
    Nome = rsMovCartaoCredito("Nome").Value
    NumeroMovimentoCaixa = rsMovCartaoCredito("Numero do Movimento do Caixa").Value
    TaxaAdministrativa = rsMovCartaoCredito("Taxa Administrativa").Value
    NumeroIlha = rsMovCartaoCredito("Numero da Ilha").Value
    Autorizacao = rsMovCartaoCredito("Autorizacao").Value
    NSU = rsMovCartaoCredito("NSU").Value
    CodigoFuncionario = rsMovCartaoCredito("Codigo do Funcionario").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function TotalAdmEntreDatas(ByVal pEmpresa As Integer, ByVal pDataEmissao As Boolean, ByVal pNaoBaixado As Boolean, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoCartao As Integer) As Currency

On Error GoTo trata_erro
    
    TotalAdmEntreDatas = 0
    gSQL = "SELECT SUM(" & preparaArredonda("Valor * [Taxa Administrativa] / 100", 2) & ") as Total FROM "
    If pNaoBaixado Then
        gSQL = gSQL & " Movimento_Cartao_Credito"
    Else
        gSQL = gSQL & " Baixa_Cartao_Credito"
    End If
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    If pDataEmissao Then
        gSQL = gSQL & " AND [Data de Emissao] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data de Emissao] <= " & preparaData(pDataFinal)
    Else
        gSQL = gSQL & " AND [Data do Vencimento] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataFinal)
    End If
    If pCodigoCartao > 0 Then
        gSQL = gSQL & " AND [Codigo do Cartao] = " & pCodigoCartao
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCartaoCredito("Total").Value) Then
            TotalAdmEntreDatas = rsMovCartaoCredito("Total").Value
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalEntreDatas(ByVal pEmpresa As Integer, ByVal pDataEmissao As Boolean, ByVal pNaoBaixado As Boolean, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoCartao As Integer) As Currency

On Error GoTo trata_erro
    
    TotalEntreDatas = 0
    gSQL = "SELECT SUM(Valor) as Total FROM "
    If pNaoBaixado Then
        gSQL = gSQL & " Movimento_Cartao_Credito"
    Else
        gSQL = gSQL & " Baixa_Cartao_Credito"
    End If
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    If pDataEmissao Then
        gSQL = gSQL & " AND [Data de Emissao] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data de Emissao] <= " & preparaData(pDataFinal)
    Else
        gSQL = gSQL & " AND [Data do Vencimento] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataFinal)
    End If
    If pCodigoCartao > 0 Then
        gSQL = gSQL & " AND [Codigo do Cartao] = " & pCodigoCartao
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCartaoCredito("Total").Value) Then
            TotalEntreDatas = rsMovCartaoCredito("Total").Value
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String, ByVal nCodigoCartao As Integer) As Currency

On Error GoTo trata_erro
    
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    gSQL = gSQL & " AND [Codigo do Cartao] = " & nCodigoCartao
    If Localizar(0) Then
        If Not IsNull(rsMovCartaoCredito("Total").Value) Then
            TotalPeriodo = rsMovCartaoCredito("Total").Value
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodoBaixado(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String, ByVal nCodigoCartao As Integer) As Currency

On Error GoTo trata_erro
    
    TotalPeriodoBaixado = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Baixa_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    gSQL = gSQL & " AND [Codigo do Cartao] = " & nCodigoCartao
    If Localizar(0) Then
        If Not IsNull(rsMovCartaoCredito("Total").Value) Then
            TotalPeriodoBaixado = rsMovCartaoCredito("Total").Value
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalAdministrativoBaixado(ByVal nEmpresa As Integer, ByVal nDataInicial As Date, ByVal nDataFinal As Date) As Currency

On Error GoTo trata_erro
    
    TotalAdministrativoBaixado = 0
    gSQL = "SELECT SUM([Valor Administrativo]) as Total FROM Baixa_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data da Baixa] >= " & preparaData(nDataInicial)
    gSQL = gSQL & " AND [Data da Baixa] <= " & preparaData(nDataFinal)
    If Localizar(0) Then
        If Not IsNull(rsMovCartaoCredito("Total").Value) Then
            TotalAdministrativoBaixado = rsMovCartaoCredito("Total").Value
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ProximoRegistro(ByVal pEmpresa As Integer, ByVal pData As Date) As Long
'Public Function ProximoRegistro(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As String) As Long

On Error GoTo Err_ProximoRegistro
    Dim xNumero As Integer

    'Pega número do último movimento
    xNumero = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Lancamento]"
    gSQL = gSQL & "  FROM Movimento_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Data de Emissao] = " & preparaData(pData)
    'gSQL = gSQL & "   AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " ORDER BY [Numero do Lancamento] DESC"
    If Localizar(1) Then
        xNumero = rsMovCartaoCredito("Numero do Lancamento").Value + 1
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Lancamento]"
    gSQL = gSQL & "  FROM Baixa_Cartao_Credito"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Data de Emissao] = " & preparaData(pData)
    'gSQL = gSQL & "   AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = gSQL & " ORDER BY [Numero do Lancamento] DESC"
    If Localizar(1) Then
        If xNumero < rsMovCartaoCredito("Numero do Lancamento").Value Then
            xNumero = rsMovCartaoCredito("Numero do Lancamento").Value + 1
        End If
    End If
    rsMovCartaoCredito.Close
    Set rsMovCartaoCredito = Nothing
    ProximoRegistro = xNumero
    Exit Function

Err_ProximoRegistro:
End Function
'Fim Métodos da Classe

Private Sub Class_Initialize()
    Set IntegracaoCaixa = New cIntegracaoCaixa
    Set MovCaixa = New cMovimentoCaixa
End Sub

Private Sub Class_Terminate()
    Set IntegracaoCaixa = Nothing
    Set MovCaixa = Nothing
End Sub
