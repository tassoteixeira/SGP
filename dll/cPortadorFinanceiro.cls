VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPortadorFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mEmpresa As Integer
Private mNumeroContaContabil As String
Private mInstituicaoFinanceira As Boolean
Private mCodigoBanco As Integer
Private mNumeroContaBancaria As String
Private mCodigoAgencia As String
Private mNomeAgencia As String
Private mCodigoCedente As String

Private rsPortadorFinanceiro As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroContaContabil() As String
    NumeroContaContabil = mNumeroContaContabil
End Property
Public Property Let NumeroContaContabil(ByVal Valor As String)
    mNumeroContaContabil = Valor
End Property


Public Property Get InstituicaoFinanceira() As Boolean
    InstituicaoFinanceira = mInstituicaoFinanceira
End Property
Public Property Let InstituicaoFinanceira(ByVal Valor As Boolean)
    mInstituicaoFinanceira = Valor
End Property


Public Property Get CodigoBanco() As Integer
    CodigoBanco = mCodigoBanco
End Property
Public Property Let CodigoBanco(ByVal Valor As Integer)
    mCodigoBanco = Valor
End Property


Public Property Get NumeroContaBancaria() As String
    NumeroContaBancaria = mNumeroContaBancaria
End Property
Public Property Let NumeroContaBancaria(ByVal Valor As String)
    mNumeroContaBancaria = Valor
End Property


Public Property Get CodigoAgencia() As String
    CodigoAgencia = mCodigoAgencia
End Property
Public Property Let CodigoAgencia(ByVal Valor As String)
    mCodigoAgencia = Valor
End Property


Public Property Get NomeAgencia() As String
    NomeAgencia = mNomeAgencia
End Property
Public Property Let NomeAgencia(ByVal Valor As String)
    mNomeAgencia = Valor
End Property


Public Property Get CodigoCedente() As String
    CodigoCedente = mCodigoCedente
End Property
Public Property Let CodigoCedente(ByVal Valor As String)
    mCodigoCedente = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarContaContabil(ByVal pContaContabil As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarContaContabil = False
    gSQL = PreparaSQL(" WHERE [Numero da Conta Contabil] = " & preparaTexto(pContaContabil), "")
    If Localizar(1) Then
        LocalizarContaContabil = True
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroContaBancaria(ByVal pNumeroContaBancaria As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroContaBancaria = False
    xCondicao = " WHERE [Numero da Conta Bancaria] = " & preparaTexto(pNumeroContaBancaria)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarNumeroContaBancaria = True
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO PortadorFinanceiro ( Codigo, Nome, Empresa, [Numero da Conta Contabil], [Instituicao Financeira], "
    gSQL = gSQL & "[Codigo do Banco], [Numero da Conta Bancaria], [Codigo da Agencia], [Nome da Agencia], [Codigo do Cedente] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mNumeroContaContabil, ", ")
    Call sqlBoolean(1, mInstituicaoFinanceira, ", ")
    Call sqlNumero(1, mCodigoBanco, ", ")
    Call sqlTexto(1, mNumeroContaBancaria, ", ")
    Call sqlTexto(1, mCodigoAgencia, ", ")
    Call sqlTexto(1, mNomeAgencia, ", ")
    Call sqlTexto(1, mCodigoCedente, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE PortadorFinanceiro SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlNumero(2, ", Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Numero da Conta Contabil] = ", mNumeroContaContabil)
    Call sqlBoolean(2, ", [Instituicao Financeira] = ", mInstituicaoFinanceira)
    Call sqlNumero(2, ", [Codigo do Banco] = ", mCodigoBanco)
    Call sqlTexto(2, ", [Numero da Conta Bancaria] = ", mNumeroContaBancaria)
    Call sqlTexto(2, ", [Codigo da Agencia] = ", mCodigoAgencia)
    Call sqlTexto(2, ", [Nome da Agencia] = ", mNomeAgencia)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM PortadorFinanceiro"
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, Empresa, [Numero da Conta Contabil], [Instituicao Financeira], [Codigo do Banco], [Numero da Conta Bancaria], [Codigo da Agencia], [Nome da Agencia] FROM PortadorFinanceiro" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPortadorFinanceiro = New adodb.Recordset
    rsPortadorFinanceiro.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPortadorFinanceiro.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPortadorFinanceiro.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsPortadorFinanceiro("Codigo").Value
    Nome = rsPortadorFinanceiro("Nome").Value
    Empresa = rsPortadorFinanceiro("Empresa").Value
    NumeroContaContabil = rsPortadorFinanceiro("Numero da Conta Contabil").Value
    InstituicaoFinanceira = rsPortadorFinanceiro("Instituicao Financeira").Value
    CodigoBanco = rsPortadorFinanceiro("Codigo do Banco").Value
    NumeroContaBancaria = rsPortadorFinanceiro("Numero da Conta Bancaria").Value
    CodigoAgencia = rsPortadorFinanceiro("Codigo da Agencia").Value
    NomeAgencia = rsPortadorFinanceiro("Nome da Agencia").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

