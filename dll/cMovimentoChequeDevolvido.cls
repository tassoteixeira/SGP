VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoChequeDevolvido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataDigitacao As Date
Private mDataEmissao As Date
Private mCodigoBanco As Integer
Private mNumeroConta As String
Private mNumeroCheque As String
Private mPeriodo As String
Private mTipoMovimento As String
Private mValor As Currency
Private mDataVencimento As Date
Private mEmitente As String
Private mOrdemDigitacao As Integer
Private mCodigoBarra1 As String
Private mCodigoBarra2 As String
Private mCodigoBarra3 As String
Private mBancoAgencia As String
Private mDataDevolucao As Date
Private mMotivoDevolucao As String
Private mSituacao As Integer
Private mCnpjCpf As String

Private rsMovimentoChequeDevolvido As New adodb.Recordset



'Inicio Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get CodigoBanco() As Integer
    CodigoBanco = mCodigoBanco
End Property
Public Property Let CodigoBanco(ByVal Valor As Integer)
    mCodigoBanco = Valor
End Property


Public Property Get NumeroConta() As String
    NumeroConta = mNumeroConta
End Property
Public Property Let NumeroConta(ByVal Valor As String)
    mNumeroConta = Valor
End Property


Public Property Get NumeroCheque() As String
    NumeroCheque = mNumeroCheque
End Property
Public Property Let NumeroCheque(ByVal Valor As String)
    mNumeroCheque = Valor
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


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal xValor As Currency)
    mValor = xValor
End Property


Public Property Get DataVencimento() As Date
    DataVencimento = mDataVencimento
End Property
Public Property Let DataVencimento(ByVal Valor As Date)
    mDataVencimento = Valor
End Property


Public Property Get Emitente() As String
    Emitente = mEmitente
End Property
Public Property Let Emitente(ByVal Valor As String)
    mEmitente = Valor
End Property


Public Property Get OrdemDigitacao() As Integer
    OrdemDigitacao = mOrdemDigitacao
End Property
Public Property Let OrdemDigitacao(ByVal Valor As Integer)
    mOrdemDigitacao = Valor
End Property


Public Property Get CodigoBarra1() As String
    CodigoBarra1 = mCodigoBarra1
End Property
Public Property Let CodigoBarra1(ByVal Valor As String)
    mCodigoBarra1 = Valor
End Property


Public Property Get CodigoBarra2() As String
    CodigoBarra2 = mCodigoBarra2
End Property
Public Property Let CodigoBarra2(ByVal Valor As String)
    mCodigoBarra2 = Valor
End Property


Public Property Get CodigoBarra3() As String
    CodigoBarra3 = mCodigoBarra3
End Property
Public Property Let CodigoBarra3(ByVal Valor As String)
    mCodigoBarra3 = Valor
End Property


Public Property Get BancoAgencia() As String
    BancoAgencia = mBancoAgencia
End Property
Public Property Let BancoAgencia(ByVal Valor As String)
    mBancoAgencia = Valor
End Property


Public Property Get DataDevolucao() As Date
    DataDevolucao = mDataDevolucao
End Property
Public Property Let DataDevolucao(ByVal Valor As Date)
    mDataDevolucao = Valor
End Property


Public Property Get MotivoDevolucao() As String
    MotivoDevolucao = mMotivoDevolucao
End Property
Public Property Let MotivoDevolucao(ByVal Valor As String)
    mMotivoDevolucao = Valor
End Property


Public Property Get Situacao() As Integer
    Situacao = mSituacao
End Property
Public Property Let Situacao(ByVal Valor As Integer)
    mSituacao = Valor
End Property


Public Property Get CnpjCpf() As String
    CnpjCpf = mCnpjCpf
End Property
Public Property Let CnpjCpf(ByVal Valor As String)
    mCnpjCpf = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function ExisteRegistro(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    ExisteRegistro = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    xSQL = xSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    xSQL = xSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    xOrdem = ""
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        ExisteRegistro = True
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarRegistro(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    LocalizarRegistro = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    xSQL = xSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    xSQL = xSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    xOrdem = ""
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        AtribuiValor
        LocalizarRegistro = True
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarOrdemDigitacao(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Integer

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    LocalizarOrdemDigitacao = 0
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    xSQL = xSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    xSQL = xSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    xOrdem = " ORDER BY Empresa, [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        LocalizarOrdemDigitacao = rsMovimentoChequeDevolvido("Ordem da Digitacao").Value
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    Dim xRegIgual As Boolean
    LocalizarAnterior = False
    xRegIgual = False
    xSQL = "WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Data de Emissao] <= " & preparaData(mDataEmissao)
    xOrdem = " ORDER BY Empresa DESC, [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC,"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC, [Numero da Conta] DESC, [Numero do Cheque] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoChequeDevolvido.RecordCount > 1 Then
            Do Until rsMovimentoChequeDevolvido.EOF
                If xRegIgual = False Then
                    If rsMovimentoChequeDevolvido("Data de Emissao").Value = mDataEmissao And rsMovimentoChequeDevolvido("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoChequeDevolvido("Periodo").Value = mPeriodo And rsMovimentoChequeDevolvido("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoChequeDevolvido("Numero da Conta").Value = mNumeroConta And rsMovimentoChequeDevolvido("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovimentoChequeDevolvido.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarConta(ByVal nConta As String) As String
Dim xData As Date

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarConta = ""
    gSQL = "SELECT Emitente, [Data de Emissao] FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE [Numero da Conta] = " & preparaTexto(nConta)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        LocalizarConta = rsMovimentoChequeDevolvido("Emitente").Value
        xData = rsMovimentoChequeDevolvido("Data de Emissao").Value
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCpfCnpj(ByVal pCpfCnpj As String) As String

On Error GoTo trata_erro
    
    LocalizarCpfCnpj = ""
    gSQL = "SELECT SUM(VALOR) AS Total, SUM(1) AS QtdCheque"
    gSQL = gSQL & "  FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE [CNPJ CPF] = " & preparaTexto(pCpfCnpj)
    If Localizar(1) Then
        If Not IsNull(rsMovimentoChequeDevolvido("Total").Value) Then
            LocalizarCpfCnpj = rsMovimentoChequeDevolvido("Total").Value & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoChequeDevolvido("QtdCheque").Value & "|@|"
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal xEmpresa As Integer) As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    LocalizarPrimeiro = False
    xSQL = "WHERE Empresa = " & xEmpresa
    xOrdem = " ORDER BY [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] ASC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        If rsMovimentoChequeDevolvido("Empresa").Value = xEmpresa Then
            AtribuiValor
            rsMovimentoChequeDevolvido.Close
        End If
    End If
    Set rsMovimentoChequeDevolvido = Nothing
    LocalizarPrimeiro = True
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    Dim xRegIgual As Boolean
    LocalizarProximo = False
    xRegIgual = False
    xSQL = "WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Data de Emissao] >= " & preparaData(mDataEmissao)
    xOrdem = " ORDER BY Empresa, [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] ASC, [Numero da Conta] ASC, [Numero do Cheque] ASC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoChequeDevolvido.RecordCount > 1 Then
            Do Until rsMovimentoChequeDevolvido.EOF
                If xRegIgual = False Then
                    If rsMovimentoChequeDevolvido("Data de Emissao").Value = mDataEmissao And rsMovimentoChequeDevolvido("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoChequeDevolvido("Periodo").Value = mPeriodo And rsMovimentoChequeDevolvido("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoChequeDevolvido("Numero da Conta").Value = mNumeroConta And rsMovimentoChequeDevolvido("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovimentoChequeDevolvido.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal xEmpresa As Integer) As Boolean

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    LocalizarUltimo = False
    xSQL = "WHERE Empresa = " & xEmpresa
    xOrdem = " ORDER BY [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC,"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        If rsMovimentoChequeDevolvido("Empresa").Value = xEmpresa Then
            AtribuiValor
            LocalizarUltimo = True
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

Dim lRecordsAffected As Long

    Incluir = False
    gSQL = "INSERT INTO Movimento_Cheque_Devolvido ( Empresa, [Data de Digitacao], [Data de Emissao], [Codigo do Banco], [Numero da Conta], [Numero do Cheque], Periodo, "
    gSQL = gSQL & "[Tipo do Movimento], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    gSQL = gSQL & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], [Data da Devolucao], [Motivo da Devolucao], Situacao, [CNPJ CPF] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlNumero(1, mCodigoBanco, ", ")
    Call sqlTexto(1, mNumeroConta, ", ")
    Call sqlTexto(1, mNumeroCheque, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlTexto(1, mEmitente, ", ")
    Call sqlNumero(1, mOrdemDigitacao, ", ")
    Call sqlTexto(1, mCodigoBarra1, ", ")
    Call sqlTexto(1, mCodigoBarra2, ", ")
    Call sqlTexto(1, mCodigoBarra3, ", ")
    Call sqlTexto(1, mBancoAgencia, ", ")
    Call sqlData(1, mDataDevolucao, ", ")
    Call sqlTexto(1, mMotivoDevolucao, ", ")
    Call sqlNumero(1, mSituacao, ", ")
    Call sqlTexto(1, mCnpjCpf, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cheque_Devolvido SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data de Digitacao] = ", mDataDigitacao)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlNumero(2, ", [Codigo do Banco] = ", mCodigoBanco)
    Call sqlTexto(2, ", [Numero da Conta] = ", mNumeroConta)
    Call sqlTexto(2, ", [Numero do Cheque] = ", mNumeroCheque)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlData(2, ", [Data do Vencimento] = ", mDataVencimento)
    Call sqlTexto(2, ", Emitente = ", mEmitente)
    Call sqlNumero(2, ", [Ordem da Digitacao] = ", mOrdemDigitacao)
    Call sqlTexto(2, ", [Codigo de Barra 1] = ", mCodigoBarra1)
    Call sqlTexto(2, ", [Codigo de Barra 2] = ", mCodigoBarra2)
    Call sqlTexto(2, ", [Codigo de Barra 3] = ", mCodigoBarra3)
    Call sqlTexto(2, ", [Banco Agencia] = ", mBancoAgencia)
    Call sqlTexto(2, ", [Data da Devolucao] = ", mDataDevolucao)
    Call sqlTexto(2, ", [Motivo da Devolucao] = ", mMotivoDevolucao)
    Call sqlNumero(2, ", Situacao = ", mSituacao)
    Call sqlTexto(2, ", [CNPJ CPF] = ", mCnpjCpf)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


Public Function TotalDataDevolucao(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Currency

On Error GoTo trata_erro
    
    TotalDataDevolucao = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Devolucao] >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND [Data da Devolucao] <= " & preparaData(pDataFinal)
    If Localizar(0) Then
        If Not IsNull(rsMovimentoChequeDevolvido("Total").Value) Then
            TotalDataDevolucao = rsMovimentoChequeDevolvido("Total").Value
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Currency

On Error GoTo trata_erro
    
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        If Not IsNull(rsMovimentoChequeDevolvido("Total").Value) Then
            TotalPeriodo = rsMovimentoChequeDevolvido("Total").Value
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalValorCpfCnpj(ByVal pCpfCnpj As String) As Currency

On Error GoTo trata_erro
    
    TotalValorCpfCnpj = 0
    gSQL = "SELECT SUM(VALOR) AS Total"
    gSQL = gSQL & "  FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE [CNPJ CPF] = " & preparaTexto(pCpfCnpj)
    If Localizar(1) Then
        If Not IsNull(rsMovimentoChequeDevolvido("Total").Value) Then
            TotalValorCpfCnpj = rsMovimentoChequeDevolvido("Total").Value
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalQtdCpfCnpj(ByVal pCpfCnpj As String) As Integer

On Error GoTo trata_erro
    
    TotalQtdCpfCnpj = 0
    gSQL = "SELECT SUM(1) AS QtdCheque"
    gSQL = gSQL & "  FROM Movimento_Cheque_Devolvido"
    gSQL = gSQL & " WHERE [CNPJ CPF] = " & preparaTexto(pCpfCnpj)
    If Localizar(1) Then
        If Not IsNull(rsMovimentoChequeDevolvido("QtdCheque").Value) Then
            TotalQtdCpfCnpj = rsMovimentoChequeDevolvido("QtdCheque").Value
        End If
    End If
    rsMovimentoChequeDevolvido.Close
    Set rsMovimentoChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


'Funções / Procedures internas
Private Function PreparaSQLGrupo(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xString As String
    xString = "Empresa, [Data de Digitacao], [Data de Emissao], [Codigo do Banco], [Numero da Conta], [Numero do Cheque], Periodo, "
    xString = xString & "[Tipo do Movimento], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    xString = xString & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], [Data da Devolucao], [Motivo da Devolucao], Situacao, [CNPJ CPF]"
    PreparaSQLGrupo = Trim("SELECT " & xString & " FROM Movimento_Cheque_Devolvido " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovimentoChequeDevolvido = New adodb.Recordset
    rsMovimentoChequeDevolvido.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovimentoChequeDevolvido.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimentoChequeDevolvido.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovimentoChequeDevolvido("Empresa").Value
    DataDigitacao = rsMovimentoChequeDevolvido("Data de Digitacao").Value
    DataEmissao = rsMovimentoChequeDevolvido("Data de Emissao").Value
    CodigoBanco = rsMovimentoChequeDevolvido("Codigo do Banco").Value
    NumeroConta = rsMovimentoChequeDevolvido("Numero da Conta").Value
    NumeroCheque = rsMovimentoChequeDevolvido("Numero do Cheque").Value
    Periodo = rsMovimentoChequeDevolvido("Periodo").Value
    TipoMovimento = rsMovimentoChequeDevolvido("Tipo do Movimento").Value
    Valor = rsMovimentoChequeDevolvido("Valor").Value
    DataVencimento = rsMovimentoChequeDevolvido("Data do Vencimento").Value
    Emitente = rsMovimentoChequeDevolvido("Emitente").Value
    OrdemDigitacao = rsMovimentoChequeDevolvido("Ordem da Digitacao").Value
    CodigoBarra1 = rsMovimentoChequeDevolvido("Codigo de Barra 1").Value
    CodigoBarra2 = rsMovimentoChequeDevolvido("Codigo de Barra 2").Value
    CodigoBarra3 = rsMovimentoChequeDevolvido("Codigo de Barra 3").Value
    BancoAgencia = rsMovimentoChequeDevolvido("Banco Agencia").Value
    DataDevolucao = rsMovimentoChequeDevolvido("Data da Devolucao").Value
    MotivoDevolucao = rsMovimentoChequeDevolvido("Motivo da Devolucao").Value
    Situacao = rsMovimentoChequeDevolvido("Situacao").Value
    CnpjCpf = rsMovimentoChequeDevolvido("CNPJ CPF").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
