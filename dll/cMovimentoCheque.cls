VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCheque"
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
Private mDadosAbastecimento As String

Private rsMovimentoCheque As New adodb.Recordset


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

Public Property Get DadosAbastecimento() As String
    DadosAbastecimento = mDadosAbastecimento
End Property
Public Property Let DadosAbastecimento(ByVal Valor As String)
    mDadosAbastecimento = Valor
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
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
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
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
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
    xOrdem = " ORDER BY [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        LocalizarOrdemDigitacao = rsMovimentoCheque("Ordem da Digitacao").Value
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
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
    xSQL = "WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Data de Emissao] <= " & preparaData(mDataEmissao)
    xOrdem = " ORDER BY Empresa DESC, [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC,"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC, [Numero da Conta] DESC, [Numero do Cheque] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoCheque.RecordCount > 1 Then
            Do Until rsMovimentoCheque.EOF
                If xRegIgual = False Then
                    If rsMovimentoCheque("Data de Emissao").Value = mDataEmissao And rsMovimentoCheque("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoCheque("Periodo").Value = mPeriodo And rsMovimentoCheque("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoCheque("Numero da Conta").Value = mNumeroConta And rsMovimentoCheque("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovimentoCheque.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarConta(ByVal nConta As String) As String
Dim xData As Date

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarConta = ""
    gSQL = "SELECT Emitente, [Data de Emissao], Telefone, [Banco Agencia], [CPF CNPJ] FROM Movimento_Cheque"
    gSQL = gSQL & " WHERE [Numero da Conta] = " & preparaTexto(nConta)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        LocalizarConta = rsMovimentoCheque("Emitente").Value & "|@|"
        LocalizarConta = LocalizarConta & rsMovimentoCheque("Telefone").Value & "|@|"
        LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("Banco Agencia").Value, 1, 3) & "|@|"
        LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("Banco Agencia").Value, 4, 4) & "|@|"
        LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("CPF CNPJ").Value, 4, 4) & "|@|"
        xData = rsMovimentoCheque("Data de Emissao").Value
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    'Localiza na Baixa
    gSQL = "SELECT Emitente, [Data de Emissao], Telefone, [Banco Agencia], [CPF CNPJ] FROM Baixa_Cheque"
    gSQL = gSQL & " WHERE [Numero da Conta] = " & preparaTexto(nConta)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        If xData < rsMovimentoCheque("Data de Emissao").Value Then
            LocalizarConta = rsMovimentoCheque("Emitente").Value & "|@|"
            LocalizarConta = LocalizarConta & rsMovimentoCheque("Telefone").Value & "|@|"
            LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("Banco Agencia").Value, 1, 3) & "|@|"
            LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("Banco Agencia").Value, 4, 4) & "|@|"
            LocalizarConta = LocalizarConta & Mid(rsMovimentoCheque("CPF CNPJ").Value, 4, 4) & "|@|"
        End If
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCpfCnpj(ByVal pCpfCnpj As String) As String
Dim xCondicao As String
Dim xOrdem As String
Dim xData As Date

On Error GoTo trata_erro
    
    'Localiza no Movimento
    LocalizarCpfCnpj = ""
    gSQL = "SELECT Emitente, [Data de Emissao], Telefone, [Banco Agencia], [CPF CNPJ], [Numero da Conta]"
    gSQL = gSQL & "  FROM Movimento_Cheque"
    gSQL = gSQL & " WHERE [CPF CNPJ] = " & preparaTexto(pCpfCnpj)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        LocalizarCpfCnpj = rsMovimentoCheque("Emitente").Value & "|@|"
        LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("Telefone").Value & "|@|"
        LocalizarCpfCnpj = LocalizarCpfCnpj & Mid(rsMovimentoCheque("Banco Agencia").Value, 1, 3) & "|@|"
        LocalizarCpfCnpj = LocalizarCpfCnpj & Mid(rsMovimentoCheque("Banco Agencia").Value, 4, 4) & "|@|"
        LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("CPF CNPJ").Value & "|@|"
        LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("Numero da Conta").Value & "|@|"
        xData = rsMovimentoCheque("Data de Emissao").Value
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    'Localiza na Baixa
    gSQL = "SELECT Emitente, [Data de Emissao], Telefone, [Banco Agencia], [CPF CNPJ], [Numero da Conta]"
    gSQL = gSQL & "  FROM Baixa_Cheque"
    gSQL = gSQL & " WHERE [CPF CNPJ] = " & preparaTexto(pCpfCnpj)
    gSQL = gSQL & " ORDER BY [Data de Emissao] DESC"
    If Localizar(1) Then
        If xData < rsMovimentoCheque("Data de Emissao").Value Then
            LocalizarCpfCnpj = rsMovimentoCheque("Emitente").Value & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("Telefone").Value & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & Mid(rsMovimentoCheque("Banco Agencia").Value, 1, 3) & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & Mid(rsMovimentoCheque("Banco Agencia").Value, 4, 4) & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("CPF CNPJ").Value & "|@|"
            LocalizarCpfCnpj = LocalizarCpfCnpj & rsMovimentoCheque("Numero da Conta").Value & "|@|"
        End If
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
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
        If rsMovimentoCheque("Empresa").Value = xEmpresa Then
            AtribuiValor
            rsMovimentoCheque.Close
        End If
    End If
    Set rsMovimentoCheque = Nothing
    LocalizarPrimeiro = True
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
    xSQL = "WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Data de Emissao] >= " & preparaData(mDataEmissao)
    xOrdem = " ORDER BY Empresa, [Data de Emissao], Periodo, [Tipo do Movimento],"
    xOrdem = xOrdem & " [Ordem da Digitacao] ASC, [Numero da Conta] ASC, [Numero do Cheque] ASC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovimentoCheque.RecordCount > 1 Then
            Do Until rsMovimentoCheque.EOF
                If xRegIgual = False Then
                    If rsMovimentoCheque("Data de Emissao").Value = mDataEmissao And rsMovimentoCheque("Tipo do Movimento").Value = mTipoMovimento And rsMovimentoCheque("Periodo").Value = mPeriodo And rsMovimentoCheque("Ordem da Digitacao").Value = mOrdemDigitacao And rsMovimentoCheque("Numero da Conta").Value = mNumeroConta And rsMovimentoCheque("Numero do Cheque").Value = mNumeroCheque Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovimentoCheque.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
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
    xOrdem = " ORDER BY Empresa DESC, [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC,"
    xOrdem = xOrdem & " [Ordem da Digitacao] DESC"
    gSQL = PreparaSQLGrupo(xSQL, xOrdem)
    If Localizar(1) Then
        If rsMovimentoCheque("Empresa").Value = xEmpresa Then
            AtribuiValor
            LocalizarUltimo = True
        End If
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Cheque ( Empresa, [Data de Emissao], [Numero da Conta], [Numero do Cheque], Periodo, "
    gSQL = gSQL & "[Tipo do Movimento], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    gSQL = gSQL & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], Telefone, "
    gSQL = gSQL & "[Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], [Numero da Ilha],  "
    gSQL = gSQL & "[Data da Custodia], [Dados do Abastecimento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEmissao, ", ")
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
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlNumero(1, mCodigoVendedor, ", ")
    Call sqlTexto(1, mCPFCNPJ, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlData(1, mDataCustodia, ", ")
    Call sqlTexto(1, mDadosAbastecimento, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nNumeroConta As String, ByVal nNumeroCheque As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cheque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
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
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlNumero(2, ", [Codigo do Vendedor] = ", mCodigoVendedor)
    Call sqlTexto(2, ", [CPF CNPJ] = ", mCPFCNPJ)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlData(2, ", [Data da Custodia] = ", mDataCustodia)
    Call sqlTexto(2, ", [Dados do Abastecimento] = ", mDadosAbastecimento)
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
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir
    Excluir = False
    If LocalizarRegistro(nEmpresa, nDataEmissao, nNumeroConta, nNumeroCheque) Then
    End If
    gSQL = "DELETE FROM Movimento_Cheque"
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


Public Function TotalEmissaoPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As String, ByVal pPeriodoFinal As String, ByVal pTipoMovimento As String, ByVal pVistaPre As String) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalEmissaoPeriodo = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM Movimento_Cheque"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data de Emissao] >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND [Data de Emissao] <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & preparaTexto(pPeriodoInicial)
    gSQL = gSQL & " AND Periodo <= " & preparaTexto(pPeriodoFinal)
    If pTipoMovimento <> "0" Then
        gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(pTipoMovimento)
    End If
    If pVistaPre = "P" Then
        gSQL = gSQL & " AND [Data do Vencimento] <> [Data de Emissao]"
    ElseIf pVistaPre = "V" Then
        gSQL = gSQL & " AND [Data do Vencimento] = [Data de Emissao]"
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovimentoCheque("Total").Value) Then
            TotalEmissaoPeriodo = rsMovimentoCheque("Total").Value
        End If
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalQtdEmissaoPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As String, ByVal pPeriodoFinal As String, ByVal pTipoMovimento As String, ByVal pVistaPre As String) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalQtdEmissaoPeriodo = 0
    gSQL = "SELECT SUM(1) AS Total FROM Movimento_Cheque"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data de Emissao] >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND [Data de Emissao] <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & preparaTexto(pPeriodoInicial)
    gSQL = gSQL & " AND Periodo <= " & preparaTexto(pPeriodoFinal)
    If pTipoMovimento <> "0" Then
        gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(pTipoMovimento)
    End If
    If pVistaPre = "P" Then
        gSQL = gSQL & " AND [Data do Vencimento] <> [Data de Emissao]"
    ElseIf pVistaPre = "V" Then
        gSQL = gSQL & " AND [Data do Vencimento] = [Data de Emissao]"
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovimentoCheque("Total").Value) Then
            TotalQtdEmissaoPeriodo = rsMovimentoCheque("Total").Value
        End If
    End If
    rsMovimentoCheque.Close
    Set rsMovimentoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe



'Funções / Procedures internas
Private Function PreparaSQLGrupo(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xString As String
    xString = "Empresa, [Data de Emissao], [Numero da Conta], [Numero do Cheque], Periodo, "
    xString = xString & "[Tipo do Movimento], Valor, [Data do Vencimento], Emitente, [Ordem da Digitacao], "
    xString = xString & "[Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3], [Banco Agencia], "
    xString = xString & "Telefone, [Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], [Numero da Ilha], "
    xString = xString & "[Data da Custodia], [Dados do Abastecimento]"
    PreparaSQLGrupo = Trim("SELECT " & xString & " FROM Movimento_Cheque " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovimentoCheque = New adodb.Recordset
    rsMovimentoCheque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovimentoCheque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimentoCheque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovimentoCheque("Empresa").Value
    DataEmissao = rsMovimentoCheque("Data de Emissao").Value
    NumeroConta = rsMovimentoCheque("Numero da Conta").Value
    NumeroCheque = rsMovimentoCheque("Numero do Cheque").Value
    Periodo = rsMovimentoCheque("Periodo").Value
    TipoMovimento = rsMovimentoCheque("Tipo do Movimento").Value
    Valor = rsMovimentoCheque("Valor").Value
    DataVencimento = rsMovimentoCheque("Data do Vencimento").Value
    Emitente = rsMovimentoCheque("Emitente").Value
    OrdemDigitacao = rsMovimentoCheque("Ordem da Digitacao").Value
    CodigoBarra1 = rsMovimentoCheque("Codigo de Barra 1").Value
    CodigoBarra2 = rsMovimentoCheque("Codigo de Barra 2").Value
    CodigoBarra3 = rsMovimentoCheque("Codigo de Barra 3").Value
    BancoAgencia = rsMovimentoCheque("Banco Agencia").Value
    Telefone = rsMovimentoCheque("Telefone").Value
    NumeroMovimentoCaixa = rsMovimentoCheque("Numero do Movimento do Caixa").Value
    CodigoVendedor = rsMovimentoCheque("Codigo do Vendedor").Value
    CPFCNPJ = rsMovimentoCheque("CPF CNPJ").Value
    NumeroIlha = rsMovimentoCheque("Numero da Ilha").Value
    If IsNull(rsMovimentoCheque("Data da Custodia").Value) Then
        DataCustodia = "00:00:00"
    Else
        DataCustodia = rsMovimentoCheque("Data da Custodia").Value
    End If
    DadosAbastecimento = rsMovimentoCheque("Dados do Abastecimento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
