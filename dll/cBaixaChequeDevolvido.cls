VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaixaChequeDevolvido"
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
Private mRecebidoPor As String
Private mDataPagamento As Date
Private mValorPagoDinheiro As Currency
Private mValorPagoChequeVista As Currency
Private mValorPagoChequePrazo As Currency
Private mSituacao As Integer
Private mCnpjCpf As String

Private rsBaixaChequeDevolvido As New adodb.Recordset



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


Public Property Get RecebidoPor() As String
    RecebidoPor = mRecebidoPor
End Property
Public Property Let RecebidoPor(ByVal Valor As String)
    mRecebidoPor = Valor
End Property


Public Property Get DataPagamento() As Date
    DataPagamento = mDataPagamento
End Property
Public Property Let DataPagamento(ByVal Valor As Date)
    mDataPagamento = Valor
End Property


Public Property Get ValorPagoDinheiro() As Currency
    ValorPagoDinheiro = mValorPagoDinheiro
End Property
Public Property Let ValorPagoDinheiro(ByVal Valor As Currency)
    mValorPagoDinheiro = Valor
End Property


Public Property Get ValorPagoChequeVista() As Currency
    ValorPagoChequeVista = mValorPagoChequeVista
End Property
Public Property Let ValorPagoChequeVista(ByVal Valor As Currency)
    mValorPagoChequeVista = Valor
End Property


Public Property Get ValorPagoChequePrazo() As Currency
    ValorPagoChequePrazo = mValorPagoChequePrazo
End Property
Public Property Let ValorPagoChequePrazo(ByVal Valor As Currency)
    mValorPagoChequePrazo = Valor
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
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
        LocalizarOrdemDigitacao = rsBaixaChequeDevolvido("Ordem da Digitacao").Value
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
        If rsBaixaChequeDevolvido.RecordCount > 1 Then
            Do Until rsBaixaChequeDevolvido.EOF
                If xRegIgual = False Then
                    If rsBaixaChequeDevolvido("Data de Emissao").Value = mDataEmissao And rsBaixaChequeDevolvido("Tipo do Movimento").Value = mTipoMovimento And rsBaixaChequeDevolvido("Periodo").Value = mPeriodo And rsBaixaChequeDevolvido("Ordem da Digitacao").Value = mOrdemDigitacao And rsBaixaChequeDevolvido("Numero da Conta").Value = mNumeroConta And rsBaixaChequeDevolvido("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsBaixaChequeDevolvido.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarConta(ByVal nConta As String) As String
Dim xData As Date

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarConta = ""
    gSQL = "SELECT Emitente, [Data de Emissao] FROM Baixa_Cheque_Devolvido"
    gSQL = gSQL & " WHERE [Numero da Conta] = " & preparaTexto(nConta)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        LocalizarConta = rsBaixaChequeDevolvido("Emitente").Value
        xData = rsBaixaChequeDevolvido("Data de Emissao").Value
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
        If rsBaixaChequeDevolvido("Empresa").Value = xEmpresa Then
            AtribuiValor
            rsBaixaChequeDevolvido.Close
        End If
    End If
    Set rsBaixaChequeDevolvido = Nothing
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
    xOrdem = " ORDER BY [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] ASC, [Numero da Conta] ASC, [Numero do Cheque] ASC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(55) Then
        If rsBaixaChequeDevolvido.RecordCount > 1 Then
            Do Until rsBaixaChequeDevolvido.EOF
                If xRegIgual = False Then
                    If rsBaixaChequeDevolvido("Data de Emissao").Value = mDataEmissao And rsBaixaChequeDevolvido("Tipo do Movimento").Value = mTipoMovimento And rsBaixaChequeDevolvido("Periodo").Value = mPeriodo And rsBaixaChequeDevolvido("Ordem da Digitacao").Value = mOrdemDigitacao And rsBaixaChequeDevolvido("Numero da Conta").Value = mNumeroConta And rsBaixaChequeDevolvido("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsBaixaChequeDevolvido.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
        If rsBaixaChequeDevolvido("Empresa").Value = xEmpresa Then
            AtribuiValor
            LocalizarUltimo = True
        End If
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

Dim lRecordsAffected As Long

    Incluir = False
    gSQL = "INSERT INTO Baixa_Cheque_Devolvido ( Empresa, [Data de Digitacao],[Data de Emissao], [Codigo do Banco], [Numero da Conta], [Numero do Cheque], Periodo, "
    gSQL = gSQL & "[Tipo do Movimento], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    gSQL = gSQL & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], "
    gSQL = gSQL & "[Data da Devolucao], [Motivo da Devolucao], [Recebido Por], [Data do Pagamento], [Valor Pago Dinheiro], "
    gSQL = gSQL & "[Valor Pago Cheque a Vista], [Valor Pago Cheque a Prazo], Situacao, [CNPJ CPF] ) VALUES ( "
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
    Call sqlTexto(1, mRecebidoPor, ", ")
    Call sqlData(1, mDataPagamento, ", ")
    Call sqlValor(1, mValorPagoDinheiro, ", ")
    Call sqlValor(1, mValorPagoChequeVista, ", ")
    Call sqlValor(1, mValorPagoChequePrazo, ", ")
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
    gSQL = "UPDATE Baixa_Cheque_Devolvido SET "
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
    Call sqlData(2, ", [Data da Devolucao] = ", mDataDevolucao)
    Call sqlTexto(2, ", [Motivo da Devolucao] = ", mMotivoDevolucao)
    Call sqlTexto(2, ", [Recebido Por] = ", mRecebidoPor)
    Call sqlData(2, ", [Data do Pagamento] = ", mDataPagamento)
    Call sqlValor(2, ", [Valor Pago Dinheiro] = ", mValorPagoDinheiro)
    Call sqlValor(2, ", [Valor Pago Cheque a Vista] = ", mValorPagoChequeVista)
    Call sqlValor(2, ", [Valor Pago Cheque a Prazo] = ", mValorPagoChequePrazo)
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
    gSQL = "DELETE FROM Baixa_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


Public Function TotalDataPagamento(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Currency

On Error GoTo trata_erro
    
    TotalDataPagamento = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM Baixa_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data do Pagamento] >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND [Data do Pagamento] <= " & preparaData(pDataFinal)
    If Localizar(0) Then
        If Not IsNull(rsBaixaChequeDevolvido("Total").Value) Then
            TotalDataPagamento = rsBaixaChequeDevolvido("Total").Value
        End If
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Currency

On Error GoTo trata_erro
    
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Baixa_Cheque_Devolvido"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        If Not IsNull(rsBaixaChequeDevolvido("Total").Value) Then
            TotalPeriodo = rsBaixaChequeDevolvido("Total").Value
        End If
    End If
    rsBaixaChequeDevolvido.Close
    Set rsBaixaChequeDevolvido = Nothing
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
    xString = xString & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], "
    xString = xString & "[Data da Devolucao], [Motivo da Devolucao], [Recebido Por], [Data do Pagamento], "
    xString = xString & "[Valor Pago Dinheiro], [Valor Pago Cheque a Vista], [Valor Pago Cheque a Prazo], Situacao, [CNPJ CPF]"
    PreparaSQLGrupo = Trim("SELECT " & xString & " FROM Baixa_Cheque_Devolvido " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBaixaChequeDevolvido = New adodb.Recordset
    rsBaixaChequeDevolvido.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBaixaChequeDevolvido.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBaixaChequeDevolvido.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsBaixaChequeDevolvido("Empresa").Value
    DataDigitacao = rsBaixaChequeDevolvido("Data de Digitacao").Value
    DataEmissao = rsBaixaChequeDevolvido("Data de Emissao").Value
    CodigoBanco = rsBaixaChequeDevolvido("Codigo do Banco").Value
    NumeroConta = rsBaixaChequeDevolvido("Numero da Conta").Value
    NumeroCheque = rsBaixaChequeDevolvido("Numero do Cheque").Value
    Periodo = rsBaixaChequeDevolvido("Periodo").Value
    TipoMovimento = rsBaixaChequeDevolvido("Tipo do Movimento").Value
    Valor = rsBaixaChequeDevolvido("Valor").Value
    DataVencimento = rsBaixaChequeDevolvido("Data do Vencimento").Value
    Emitente = rsBaixaChequeDevolvido("Emitente").Value
    OrdemDigitacao = rsBaixaChequeDevolvido("Ordem da Digitacao").Value
    CodigoBarra1 = rsBaixaChequeDevolvido("Codigo de Barra 1").Value
    CodigoBarra2 = rsBaixaChequeDevolvido("Codigo de Barra 2").Value
    CodigoBarra3 = rsBaixaChequeDevolvido("Codigo de Barra 3").Value
    BancoAgencia = rsBaixaChequeDevolvido("Banco Agencia").Value
    DataDevolucao = rsBaixaChequeDevolvido("Data da Devolucao").Value
    MotivoDevolucao = rsBaixaChequeDevolvido("Motivo da Devolucao").Value
    RecebidoPor = rsBaixaChequeDevolvido("Recebido Por").Value
    DataPagamento = rsBaixaChequeDevolvido("Data do Pagamento").Value
    ValorPagoDinheiro = rsBaixaChequeDevolvido("Valor Pago Dinheiro").Value
    ValorPagoChequeVista = rsBaixaChequeDevolvido("Valor Pago Cheque a Vista").Value
    ValorPagoChequePrazo = rsBaixaChequeDevolvido("Valor Pago Cheque a Prazo").Value
    Situacao = rsBaixaChequeDevolvido("Situacao").Value
    CnpjCpf = rsBaixaChequeDevolvido("CNPJ CPF").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
