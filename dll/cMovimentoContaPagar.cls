VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoContaPagar"
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
Private mNumeroMovimentoCaixa As Long

Private rsContasPagar As New adodb.Recordset


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


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE empresa = " & nEmpresa
    xCondicao = xCondicao & " AND registro = " & nRegistro
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE empresa = " & mEmpresa
    xCondicao = xCondicao & " AND registro < " & mRegistro
    gSQL = PreparaSQL(xCondicao, "ORDER BY empresa DESC, registro DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
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
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY empresa ASC, registro ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE empresa = " & mEmpresa
    xCondicao = xCondicao & " AND registro > " & mRegistro
    gSQL = PreparaSQL(xCondicao, "ORDER BY empresa ASC, registro ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY empresa DESC, registro DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
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
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
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
        ProximoRegistro = rsContasPagar("registro").Value + 1
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT registro"
    gSQL = gSQL & "  FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        If ProximoRegistro <= rsContasPagar("registro").Value Then
            ProximoRegistro = rsContasPagar("registro").Value + 1
        End If
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

Err_ProximoRegistro:
End Function


Public Function LocalizaCobrancaConta(ByVal nEmpresa As Integer, ByVal nFornecedor As Integer) As String
Dim xRegistro As Long

On Error GoTo Err_LocalizaCobrancaConta

    LocalizaCobrancaConta = ""
    xRegistro = 0
    gSQL = ""
    gSQL = gSQL & "SELECT registro, local_cobranca, codigo_conta"
    gSQL = gSQL & "  FROM Contas_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "   AND codigo_fornecedor = " & nFornecedor
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        xRegistro = rsContasPagar("registro").Value
        LocalizaCobrancaConta = rsContasPagar("local_cobranca").Value & "|@|"
        LocalizaCobrancaConta = LocalizaCobrancaConta & rsContasPagar("codigo_conta").Value & "|@|"
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT registro, local_cobranca, codigo_conta"
    gSQL = gSQL & "  FROM Baixa_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "   AND codigo_fornecedor = " & nFornecedor
    gSQL = gSQL & " ORDER BY registro DESC"
    If Localizar(1) Then
        If xRegistro < rsContasPagar("registro").Value Then
            LocalizaCobrancaConta = rsContasPagar("local_cobranca").Value & "|@|"
            LocalizaCobrancaConta = LocalizaCobrancaConta & rsContasPagar("codigo_conta").Value & "|@|"
        End If
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

Err_LocalizaCobrancaConta:
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mRegistro = ProximoRegistro(mEmpresa)
    gSQL = "INSERT INTO Contas_Pagar ( empresa, registro, codigo_fornecedor, nome_fornecedor, data_emissao, "
    gSQL = gSQL & "data_vencimento, valor, numero_documento, local_cobranca, codigo_conta, "
    gSQL = gSQL & "complemento, [Data da Digitacao], [Numero do Movimento do Caixa] ) VALUES ( "
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
    Call sqlNumero(1, mNumeroMovimentoCaixa, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pRegistro As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Contas_Pagar SET "
    Call sqlNumero(2, "empresa = ", mEmpresa)
    Call sqlNumero(2, ", registro = ", mRegistro)
    Call sqlNumero(2, ", codigo_fornecedor = ", mCodigoFornecedor)
    Call sqlTexto(2, ", nome_fornecedor = ", mNomeFornecedor)
    Call sqlData(2, ", data_emissao = ", mDataEmissao)
    Call sqlData(2, ", data_vencimento = ", mDataVencimento)
    Call sqlValor(2, ", valor = ", mValor)
    Call sqlTexto(2, ", numero_documento = ", mNumeroDocumento)
    Call sqlNumero(2, ", local_cobranca = ", mLocalCobranca)
    Call sqlNumero(2, ", codigo_conta = ", mCodigoConta)
    Call sqlTexto(2, ", complemento = ", mComplemento)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    gSQL = gSQL & " WHERE empresa = " & pEmpresa
    gSQL = gSQL & " AND registro = " & pRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlteraNumeroRegistroAutomatico(ByVal pEmpresa As Integer, ByVal pRegistro As Long) As Long
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    AlteraNumeroRegistroAutomatico = 0
    mRegistro = ProximoRegistro(pEmpresa)
    gSQL = "UPDATE Contas_Pagar SET "
    Call sqlNumero(2, "registro = ", mRegistro)
    gSQL = gSQL & " WHERE empresa = " & pEmpresa
    gSQL = gSQL & " AND registro = " & pRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlteraNumeroRegistroAutomatico = mRegistro
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nRegistro As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Contas_Pagar"
    gSQL = gSQL & " WHERE empresa = " & nEmpresa
    gSQL = gSQL & " AND registro = " & nRegistro
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT empresa, registro, codigo_fornecedor, nome_fornecedor, data_emissao, data_vencimento, valor, numero_documento, local_cobranca, codigo_conta, complemento, [Data da Digitacao], [Numero do Movimento do Caixa] FROM Contas_Pagar" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsContasPagar = New adodb.Recordset
    rsContasPagar.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsContasPagar.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsContasPagar.EOF Then
        AtualizaRecordset = True
    End If
End Function


Public Function PrimeiraDataVencimento(ByVal pEmpresa As Integer) As Date

On Error GoTo Err_Rotina
    PrimeiraDataVencimento = Date
    
    gSQL = ""
    gSQL = gSQL & "SELECT data_vencimento"
    gSQL = gSQL & "  FROM Contas_Pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY data_vencimento ASC"
    If Localizar(1) Then
        If Not IsNull(rsContasPagar("data_vencimento").Value) Then
            PrimeiraDataVencimento = rsContasPagar("data_vencimento").Value
        End If
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

Err_Rotina:
    MsgBox "Erro na Rotina: PrimeiraDataVencimento", vbInformation
    Exit Function
End Function


Private Sub AtribuiValor()
    Empresa = rsContasPagar("empresa").Value
    Registro = rsContasPagar("registro").Value
    CodigoFornecedor = rsContasPagar("codigo_fornecedor").Value
    NomeFornecedor = rsContasPagar("nome_fornecedor").Value
    DataEmissao = rsContasPagar("data_emissao").Value
    DataVencimento = rsContasPagar("data_vencimento").Value
    Valor = rsContasPagar("valor").Value
    NumeroDocumento = rsContasPagar("numero_documento").Value
    LocalCobranca = rsContasPagar("local_cobranca").Value
    CodigoConta = rsContasPagar("codigo_conta").Value
    Complemento = rsContasPagar("complemento").Value
    DataDigitacao = rsContasPagar("Data da Digitacao").Value
    NumeroMovimentoCaixa = rsContasPagar("Numero do Movimento do Caixa").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

Public Function TotalEntreDatas(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Currency

On Error GoTo trata_erro
    
    TotalEntreDatas = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM contas_pagar"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND data_vencimento >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND data_vencimento <= " & preparaData(pDataFinal)
    If Localizar(0) Then
        If Not IsNull(rsContasPagar("Total").Value) Then
            TotalEntreDatas = rsContasPagar("Total").Value
        End If
    End If
    rsContasPagar.Close
    Set rsContasPagar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

