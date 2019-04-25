VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoBaixaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataEmissao As Date
Private mNumeroConta As String
Private mNumeroCheque As String
Private mPeriodo As String
Private mTipoMovimento As String
Private mAgencia As String
Private mValor As Currency
Private mDataVencimento As Date
Private mEmitente As String
Private mOrdemDigitacao As Integer
Private mCodigoBarra1 As String
Private mCodigoBarra2 As String
Private mCodigoBarra3 As String
Private mBancoAgencia As String
Private mTelefone As String
Private mNumeroMovimentoCaixa As Long
Private mCodigoVendedor As Integer
Private mCPFCNPJ As String
Private mNumeroIlha As Integer
Private mDataCustodia As Date
Private mDataPagamento As Date
Private mPeriodoPagamento As Integer

Private rsMovimentoBaixaCheque As New adodb.Recordset
Dim lRecordsAffected As Long



'Inicio Propriedades da Classe
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


Public Property Get Agencia() As String
    Agencia = mAgencia
End Property
Public Property Let Agencia(ByVal Valor As String)
    mAgencia = Valor
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


Public Property Get Telefone() As String
    Telefone = mTelefone
End Property
Public Property Let Telefone(ByVal Valor As String)
    mTelefone = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get CodigoVendedor() As Integer
    CodigoVendedor = mCodigoVendedor
End Property
Public Property Let CodigoVendedor(ByVal Valor As Integer)
    mCodigoVendedor = Valor
End Property


Public Property Get CPFCNPJ() As String
    CPFCNPJ = mCPFCNPJ
End Property
Public Property Let CPFCNPJ(ByVal Valor As String)
    mCPFCNPJ = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get DataCustodia() As Date
    DataCustodia = mDataCustodia
End Property
Public Property Let DataCustodia(ByVal Valor As Date)
    mDataCustodia = Valor
End Property


Public Property Get DataPagamento() As Date
    DataPagamento = mDataPagamento
End Property
Public Property Let DataPagamento(ByVal Valor As Date)
    mDataPagamento = Valor
End Property


Public Property Get PeriodoPagamento() As Integer
    PeriodoPagamento = mPeriodoPagamento
End Property
Public Property Let PeriodoPagamento(ByVal Valor As Integer)
    mPeriodoPagamento = Valor
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(1) Then
        ExisteRegistro = True
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(1) Then
        AtribuiValor
        LocalizarRegistro = True
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(1) Then
        LocalizarOrdemDigitacao = rsMovimentoBaixaCheque("Ordem da Digitacao").Value
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoBaixaCheque.RecordCount > 1 Then
            Do Until rsMovimentoBaixaCheque.EOF
                If xRegIgual = False Then
                    If rsMovimentoBaixaCheque("Data de Emissao").Value = mDataEmissao And rsMovimentoBaixaCheque("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoBaixaCheque("Periodo").Value = mPeriodo And rsMovimentoBaixaCheque("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoBaixaCheque("Numero da Conta").Value = mNumeroConta And rsMovimentoBaixaCheque("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovimentoBaixaCheque.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarConta(ByVal nConta As String) As String
Dim xData As Date

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarConta = ""
    gSQL = "SELECT Emitente, [Data de Emissao] FROM Baixa_Cheque"
    gSQL = gSQL & " WHERE [Numero da Conta] = " & preparaTexto(nConta)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        LocalizarConta = rsMovimentoBaixaCheque("Emitente").Value
        xData = rsMovimentoBaixaCheque("Data de Emissao").Value
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCpfCnpj(ByVal pCpfCnpj As String) As Boolean
Dim xCondicao As String
Dim xOrdem As String

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarCpfCnpj = False
    xCondicao = "WHERE [CPF CNPJ] = " & preparaTexto(pCpfCnpj)
    xOrdem = " ORDER BY [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC, [Ordem da Digitacao] DESC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(1) Then
        LocalizarCpfCnpj = True
        AtribuiValor
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(1) Then
        LocalizarPrimeiro = True
        AtribuiValor
        rsMovimentoBaixaCheque.Close
    End If
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoBaixaCheque.RecordCount > 1 Then
            Do Until rsMovimentoBaixaCheque.EOF
                If xRegIgual = False Then
                    If rsMovimentoBaixaCheque("Data de Emissao").Value = mDataEmissao And rsMovimentoBaixaCheque("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoBaixaCheque("Periodo").Value = mPeriodo And rsMovimentoBaixaCheque("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoBaixaCheque("Numero da Conta").Value = mNumeroConta And rsMovimentoBaixaCheque("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovimentoBaixaCheque.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
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
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(1) Then
        AtribuiValor
        LocalizarUltimo = True
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir


    Incluir = False
    gSQL = "INSERT INTO Baixa_Cheque ( Empresa, [Data de Emissao], [Numero da Conta], [Numero do Cheque], Periodo, "
    gSQL = gSQL & "[Tipo do Movimento], [Agencia], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    gSQL = gSQL & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], Telefone, "
    gSQL = gSQL & "[Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], [Numero da Ilha], "
    gSQL = gSQL & "[Data da Custodia], [Data do Pagamento], [Periodo do Pagamento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlTexto(1, mNumeroConta, ", ")
    Call sqlTexto(1, mNumeroCheque, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlTexto(1, mAgencia, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlTexto(1, mEmitente, ", ")
    Call sqlNumero(1, mOrdemDigitacao, ", ")
    Call sqlTexto(1, mCodigoBarra1, ", ")
    Call sqlTexto(1, mCodigoBarra2, ", ")
    Call sqlTexto(1, mCodigoBarra3, ", ")
    Call sqlTexto(1, mBancoAgencia, ", ")
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlNumero(1, mCodigoVendedor, ", ")
    Call sqlTexto(1, mCPFCNPJ, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlData(1, mDataCustodia, ", ")
    Call sqlData(1, mDataPagamento, ", ")
    Call sqlNumero(1, mPeriodoPagamento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Baixa_Cheque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlTexto(2, ", [Numero da Conta] = ", mNumeroConta)
    Call sqlTexto(2, ", [Numero do Cheque] = ", mNumeroCheque)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlTexto(2, ", Agencia = ", mAgencia)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlData(2, ", [Data do Vencimento] = ", mDataVencimento)
    Call sqlTexto(2, ", Emitente = ", mEmitente)
    Call sqlNumero(2, ", [Ordem da Digitacao] = ", mOrdemDigitacao)
    Call sqlTexto(2, ", [Codigo de Barra 1] = ", mCodigoBarra1)
    Call sqlTexto(2, ", [Codigo de Barra 2] = ", mCodigoBarra2)
    Call sqlTexto(2, ", [Codigo de Barra 3] = ", mCodigoBarra3)
    Call sqlTexto(2, ", [Banco Agencia] = ", mBancoAgencia)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlNumero(2, ", [Codigo do Vendedor] = ", mCodigoVendedor)
    Call sqlTexto(2, ", [CPF CNPJ] = ", mCPFCNPJ)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlData(2, ", [Data da Custodia] = ", mDataCustodia)
    Call sqlData(2, ", [Data do Pagamento] = ", mDataPagamento)
    Call sqlNumero(2, ", [Periodo do Pagamento] = ", mPeriodoPagamento)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Baixa_Cheque"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(nNumeroConta)
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(nNumeroCheque)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) as Total FROM Baixa_Cheque"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        TotalPeriodo = rsMovimentoBaixaCheque("Total").Value
    End If
    rsMovimentoBaixaCheque.Close
    Set rsMovimentoBaixaCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xString As String
    xString = "Empresa, [Data de Emissao], [Numero da Conta], [Numero do Cheque], Periodo, "
    xString = xString & "[Tipo do Movimento], Agencia, Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    xString = xString & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], Telefone, "
    xString = xString & "[Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], [Numero da Ilha], "
    xString = xString & "[Data da Custodia], [Data do Pagamento], [Periodo do Pagamento] "
    PreparaSQL = Trim("SELECT " & xString & " FROM Baixa_Cheque " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovimentoBaixaCheque = New adodb.Recordset
    rsMovimentoBaixaCheque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovimentoBaixaCheque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimentoBaixaCheque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovimentoBaixaCheque("Empresa").Value
    DataEmissao = rsMovimentoBaixaCheque("Data de Emissao").Value
    NumeroConta = rsMovimentoBaixaCheque("Numero da Conta").Value
    NumeroCheque = rsMovimentoBaixaCheque("Numero do Cheque").Value
    Periodo = rsMovimentoBaixaCheque("Periodo").Value
    TipoMovimento = rsMovimentoBaixaCheque("Tipo do Movimento").Value
    Agencia = rsMovimentoBaixaCheque("Agencia").Value
    Valor = rsMovimentoBaixaCheque("Valor").Value
    DataVencimento = rsMovimentoBaixaCheque("Data do Vencimento").Value
    Emitente = rsMovimentoBaixaCheque("Emitente").Value
    OrdemDigitacao = rsMovimentoBaixaCheque("Ordem da Digitacao").Value
    CodigoBarra1 = rsMovimentoBaixaCheque("Codigo de Barra 1").Value
    CodigoBarra2 = rsMovimentoBaixaCheque("Codigo de Barra 2").Value
    CodigoBarra3 = rsMovimentoBaixaCheque("Codigo de Barra 3").Value
    BancoAgencia = rsMovimentoBaixaCheque("Banco Agencia").Value
    Telefone = rsMovimentoBaixaCheque("Telefone").Value
    NumeroMovimentoCaixa = rsMovimentoBaixaCheque("Numero do Movimento do Caixa").Value
    CodigoVendedor = rsMovimentoBaixaCheque("Codigo do Vendedor").Value
    CPFCNPJ = rsMovimentoBaixaCheque("CPF CNPJ").Value
    NumeroIlha = rsMovimentoBaixaCheque("Numero da Ilha").Value
    If IsNull(rsMovimentoBaixaCheque("Data da Custodia").Value) Then
        DataCustodia = "00:00:00"
    Else
        DataCustodia = rsMovimentoBaixaCheque("Data da Custodia").Value
    End If
    If IsNull(rsMovimentoBaixaCheque("Data do Pagamento").Value) Then
        DataPagamento = "00:00:00"
    Else
        DataPagamento = rsMovimentoBaixaCheque("Data do Pagamento").Value
    End If
    PeriodoPagamento = rsMovimentoBaixaCheque("Periodo do Pagamento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

