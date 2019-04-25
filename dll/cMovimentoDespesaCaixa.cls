VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoDespesaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mRegistro As Long
Private mDataDigitacao As Date
Private mDataMovimento As Date
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mTipoMovimento As Integer
Private mCodigoFornecedor As Integer
Private mValor As Currency
Private mNumeroDocumento As String
Private mCodigoConta As Integer
Private mComplemento As String
Private mNumeroContaContabil As String
Private mNumeroMovimentoCaixaCredito As Long
Private mNumeroMovimentoCaixaDebito As Long

Private rsMovDespesaCaixa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Registro() As Long
    Registro = mRegistro
End Property
Public Property Let Registro(ByVal Valor As Long)
    mRegistro = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get DataMovimento() As Date
    DataMovimento = mDataMovimento
End Property
Public Property Let DataMovimento(ByVal Valor As Date)
    mDataMovimento = Valor
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


Public Property Get CodigoFornecedor() As Integer
    CodigoFornecedor = mCodigoFornecedor
End Property
Public Property Let CodigoFornecedor(ByVal Valor As Integer)
    mCodigoFornecedor = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get CodigoConta() As Integer
    CodigoConta = mCodigoConta
End Property
Public Property Let CodigoConta(ByVal Valor As Integer)
    mCodigoConta = Valor
End Property


Public Property Get Complemento() As String
    Complemento = mComplemento
End Property
Public Property Let Complemento(ByVal Valor As String)
    mComplemento = Valor
End Property


Public Property Get NumeroContaContabil() As String
    NumeroContaContabil = mNumeroContaContabil
End Property
Public Property Let NumeroContaContabil(ByVal Valor As String)
    mNumeroContaContabil = Valor
End Property


Public Property Get NumeroMovimentoCaixaCredito() As Long
    NumeroMovimentoCaixaCredito = mNumeroMovimentoCaixaCredito
End Property
Public Property Let NumeroMovimentoCaixaCredito(ByVal Valor As Long)
    mNumeroMovimentoCaixaCredito = Valor
End Property


Public Property Get NumeroMovimentoCaixaDebito() As Long
    NumeroMovimentoCaixaDebito = mNumeroMovimentoCaixaDebito
End Property
Public Property Let NumeroMovimentoCaixaDebito(ByVal Valor As Long)
    mNumeroMovimentoCaixaDebito = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Registro = " & nRegistro
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Registro < " & mRegistro
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Registro DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Registro > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Registro ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Registro > " & mRegistro
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Registro ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Registro < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Registro DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoFornecedor(ByVal pEmpresa As Integer, ByVal pFornecedor As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoFornecedor = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & "   AND [Codigo do Fornecedor] = " & pFornecedor
    gSQL = PreparaSQL(xCondicao, "ORDER BY Registro DESC")
    LocalizarUltimoFornecedor = Localizar(1)
    If LocalizarUltimoFornecedor = True Then
        AtribuiValor
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ProximoRegistro(ByVal pEmpresa As Integer) As Long

On Error GoTo Err_ProximoRegistro

    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT Registro"
    gSQL = gSQL & "  FROM Movimento_Despesa_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Registro DESC"
    If Localizar(1) Then
        ProximoRegistro = rsMovDespesaCaixa("Registro").Value + 1
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT registro"
    gSQL = gSQL & "  FROM Contas_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        If ProximoRegistro <= rsMovDespesaCaixa("registro").Value Then
            ProximoRegistro = rsMovDespesaCaixa("registro").Value + 1
        End If
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT registro"
    gSQL = gSQL & "  FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        If ProximoRegistro <= rsMovDespesaCaixa("registro").Value Then
            ProximoRegistro = rsMovDespesaCaixa("registro").Value + 1
        End If
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

Err_ProximoRegistro:
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Despesa_Caixa ( Empresa, Registro, [Data da Digitacao], [Data do Movimento], Periodo, "
    gSQL = gSQL & "[Numero da Ilha], [Tipo do Movimento], [Codigo do Fornecedor], Valor, [Numero do Documento], "
    gSQL = gSQL & "[Codigo da Conta], Complemento, [Numero da Conta Contabil], [Numero do Movimento do Caixa Credito], "
    gSQL = gSQL & "[Numero do Movimento do Caixa Debito] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mRegistro, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlData(1, mDataMovimento, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoFornecedor, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoConta, ", ")
    Call sqlTexto(1, mComplemento, ", ")
    Call sqlTexto(1, mNumeroContaContabil, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaCredito, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaDebito, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Despesa_Caixa SET "
    Call sqlNumero(2, "Empresa = ", nEmpresa)
    Call sqlNumero(2, ", Registro = ", mRegistro)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlData(2, ", [Data do Movimento] = ", mDataMovimento)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Fornecedor] = ", mCodigoFornecedor)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo da Conta] = ", mCodigoConta)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    Call sqlTexto(2, ", [Numero da Conta Contabil] = ", mNumeroContaContabil)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Credito] = ", mNumeroMovimentoCaixaCredito)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Debito] = ", mNumeroMovimentoCaixaDebito)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Registro = " & nRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    If LocalizarCodigo(nEmpresa, nRegistro) Then
    End If
    gSQL = "DELETE FROM Movimento_Despesa_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Registro = " & nRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xString As String
    xString = ""
    xString = xString & "Empresa, Registro, [Data da Digitacao], [Data do Movimento], Periodo, "
    xString = xString & "[Numero da Ilha], [Tipo do Movimento], [Codigo do Fornecedor], Valor, "
    xString = xString & "[Numero do Documento], [Codigo da Conta], Complemento, [Numero da Conta Contabil], "
    xString = xString & "[Numero do Movimento do Caixa Credito], [Numero do Movimento do Caixa Debito]"
    PreparaSQL = Trim("SELECT " & xString & " FROM Movimento_Despesa_Caixa " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovDespesaCaixa = New adodb.Recordset
    rsMovDespesaCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovDespesaCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovDespesaCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovDespesaCaixa("Empresa").Value
    Registro = rsMovDespesaCaixa("Registro").Value
    DataDigitacao = rsMovDespesaCaixa("Data da Digitacao").Value
    DataMovimento = rsMovDespesaCaixa("Data do Movimento").Value
    Periodo = rsMovDespesaCaixa("Periodo").Value
    NumeroIlha = rsMovDespesaCaixa("Numero da Ilha").Value
    TipoMovimento = rsMovDespesaCaixa("Tipo do Movimento").Value
    CodigoFornecedor = rsMovDespesaCaixa("Codigo do Fornecedor").Value
    Valor = rsMovDespesaCaixa("Valor").Value
    NumeroDocumento = rsMovDespesaCaixa("Numero do Documento").Value
    CodigoConta = rsMovDespesaCaixa("Codigo da Conta").Value
    Complemento = rsMovDespesaCaixa("Complemento").Value
    NumeroContaContabil = rsMovDespesaCaixa("Numero da Conta Contabil").Value
    NumeroMovimentoCaixaCredito = rsMovDespesaCaixa("Numero do Movimento do Caixa Credito").Value
    NumeroMovimentoCaixaDebito = rsMovDespesaCaixa("Numero do Movimento do Caixa Debito").Value
End Sub


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataMovimento As Date, ByVal nPeriodo As Integer, ByVal nTipoMovimento As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM Movimento_Despesa_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data do Movimento] = " & preparaData(nDataMovimento)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Tipo do Movimento] = " & nTipoMovimento
    If Localizar(0) Then
        If Not IsNull(rsMovDespesaCaixa("Total").Value) Then
            TotalPeriodo = rsMovDespesaCaixa("Total").Value
        End If
    End If
    rsMovDespesaCaixa.Close
    Set rsMovDespesaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function



