VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaixaPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mRegistro As Long
Private mCodigoFornecedor As Integer
Private mNomeFornecedor As String
Private mDataEmissao As Date
Private mDataVencimento As Date
Private mValor As Currency
Private mNumeroDocumento As String
Private mLocalCobranca As Integer
Private mCodigoConta As Integer
Private mComplemento As String
Private mDataDigitacao As Date
Private mDataPagamento As Date
Private mValorPagamento As Currency
Private mNumeroMovimentoCaixa As Long
Private mNumeroMovimentoCaixaBaixa As Long
Private mTipoBaixa As Integer

Private rsBaixaPagar As New adodb.Recordset


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


Public Property Get CodigoFornecedor() As Integer
    CodigoFornecedor = mCodigoFornecedor
End Property
Public Property Let CodigoFornecedor(ByVal Valor As Integer)
    mCodigoFornecedor = Valor
End Property


Public Property Get NomeFornecedor() As String
    NomeFornecedor = mNomeFornecedor
End Property
Public Property Let NomeFornecedor(ByVal Valor As String)
    mNomeFornecedor = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
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


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get LocalCobranca() As Integer
    LocalCobranca = mLocalCobranca
End Property
Public Property Let LocalCobranca(ByVal Valor As Integer)
    mLocalCobranca = Valor
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


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get DataPagamento() As Date
    DataPagamento = mDataPagamento
End Property
Public Property Let DataPagamento(ByVal Valor As Date)
    mDataPagamento = Valor
End Property


Public Property Get ValorPagamento() As Currency
    ValorPagamento = mValorPagamento
End Property
Public Property Let ValorPagamento(ByVal Valor As Currency)
    mValorPagamento = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get NumeroMovimentoCaixaBaixa() As Long
    NumeroMovimentoCaixaBaixa = mNumeroMovimentoCaixaBaixa
End Property
Public Property Let NumeroMovimentoCaixaBaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixaBaixa = Valor
End Property


Public Property Get TipoBaixa() As Integer
    TipoBaixa = mTipoBaixa
End Property
Public Property Let TipoBaixa(ByVal Valor As Integer)
    mTipoBaixa = Valor
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
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY Registro DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
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
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Registro ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY Registro ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Registro DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoFornecedor(ByVal pEmpresa As Integer, ByVal pFornecedor As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoFornecedor = False
    xCondicao = " WHERE empresa = " & pEmpresa
    xCondicao = xCondicao & "   AND codigo_fornecedor = " & pFornecedor
    gSQL = PreparaSQL(xCondicao, "ORDER BY registro DESC")
    LocalizarUltimoFornecedor = Localizar(1)
    If LocalizarUltimoFornecedor = True Then
        AtribuiValor
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Baixa_Pagar ( Empresa, Registro, Codigo_Fornecedor, Nome_Fornecedor, Data_Emissao, "
    gSQL = gSQL & "Data_Vencimento, Valor, Numero_Documento, Local_Cobranca, Codigo_Conta, "
    gSQL = gSQL & "Complemento, [Data da Digitacao], Data_Pagamento, Valor_Pagamento, "
    gSQL = gSQL & "[Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], [Tipo de Baixa] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mRegistro, ", ")
    Call sqlNumero(1, mCodigoFornecedor, ", ")
    Call sqlTexto(1, mNomeFornecedor, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mLocalCobranca, ", ")
    Call sqlNumero(1, mCodigoConta, ", ")
    Call sqlTexto(1, mComplemento, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlData(1, mDataPagamento, ", ")
    Call sqlValor(1, mValorPagamento, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaBaixa, ", ")
    Call sqlNumero(1, mTipoBaixa, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
    CriaLogCadastroDll ("Classe:cBaixaPagar - Método:Incluir")
    CriaLogCadastroDll (Err & " - " & Error)
    CriaLogCadastroDll ("SQL:" & gSQL)
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Baixa_Pagar SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Registro = ", mRegistro)
    Call sqlNumero(2, ", Codigo_Fornecedor = ", mCodigoFornecedor)
    Call sqlTexto(2, ", Nome_Fornecedor = ", mNomeFornecedor)
    Call sqlData(2, ", Data_Emissao = ", mDataEmissao)
    Call sqlData(2, ", Data_Vencimento = ", mDataVencimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", Numero_Documento = ", mNumeroDocumento)
    Call sqlNumero(2, ", Local_Cobranca = ", mLocalCobranca)
    Call sqlNumero(2, ", Codigo_Conta = ", mCodigoConta)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlData(2, ", Data_Pagamento = ", mDataPagamento)
    Call sqlValor(2, ", Valor_Pagamento = ", mValorPagamento)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Baixa] = ", mNumeroMovimentoCaixaBaixa)
    Call sqlNumero(2, ", [Tipo de Baixa] = ", mTipoBaixa)
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
    gSQL = "DELETE FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Registro = " & nRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ProximoRegistro(ByVal pEmpresa As Integer) As Long

On Error GoTo Err_ProximoRegistro

    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT registro"
    gSQL = gSQL & "  FROM Contas_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        ProximoRegistro = rsBaixaPagar("registro").Value + 1
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT registro"
    gSQL = gSQL & "  FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        If ProximoRegistro <= rsBaixaPagar("registro").Value Then
            ProximoRegistro = rsBaixaPagar("registro").Value + 1
        End If
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

Err_ProximoRegistro:
End Function


Public Function TotalPeriodoConta(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pConta As Integer) As Currency

On Error GoTo trata_erro
    TotalPeriodoConta = 0
    gSQL = "SELECT SUM(Valor_Pagamento) as Total FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data_Pagamento >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data_Pagamento <= " & preparaData(pDataFinal)
    If pConta > 0 Then
        gSQL = gSQL & " AND Codigo_Conta = " & pConta
    End If
    If Localizar(0) Then
        If Not IsNull(rsBaixaPagar("Total").Value) Then
            TotalPeriodoConta = rsBaixaPagar("Total").Value
        End If
    End If
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Registro, Codigo_Fornecedor, Nome_Fornecedor, Data_Emissao, Data_Vencimento, Valor, Numero_Documento, Local_Cobranca, Codigo_Conta, Complemento, [Data da Digitacao], Data_Pagamento, Valor_Pagamento, [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], [Tipo de Baixa] FROM Baixa_Pagar" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBaixaPagar = New adodb.Recordset
    rsBaixaPagar.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBaixaPagar.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBaixaPagar.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsBaixaPagar("Empresa").Value
    Registro = rsBaixaPagar("Registro").Value
    CodigoFornecedor = rsBaixaPagar("Codigo_Fornecedor").Value
    NomeFornecedor = rsBaixaPagar("Nome_Fornecedor").Value
    DataEmissao = rsBaixaPagar("Data_Emissao").Value
    DataVencimento = rsBaixaPagar("Data_Vencimento").Value
    Valor = rsBaixaPagar("Valor").Value
    NumeroDocumento = rsBaixaPagar("Numero_Documento").Value
    LocalCobranca = rsBaixaPagar("Local_Cobranca").Value
    CodigoConta = rsBaixaPagar("Codigo_Conta").Value
    Complemento = rsBaixaPagar("Complemento").Value
    DataDigitacao = rsBaixaPagar("Data da Digitacao").Value
    DataPagamento = rsBaixaPagar("Data_Pagamento").Value
    ValorPagamento = rsBaixaPagar("Valor_Pagamento").Value
    NumeroMovimentoCaixa = rsBaixaPagar("Numero do Movimento do Caixa").Value
    NumeroMovimentoCaixaBaixa = rsBaixaPagar("Numero do Movimento do Caixa Baixa").Value
    TipoBaixa = rsBaixaPagar("Tipo de Baixa").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

