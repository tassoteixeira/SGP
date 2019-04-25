VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As Integer
Private mNome As String
Private mEndereco As String
Private mBairro As String
Private mCidade As String
Private mUF As String
Private mCEP As String
Private mTelefone As String
Private mTelefone2 As String
Private mFax As String
Private mVendedor As String
Private mCGC As String
Private mInscricaoEstadual As String
Private mCodigoConta As Integer
Private mContaContabil As String
Private mDataAlteracao As Date
Private mInativo As Boolean

Private rsFornecedor As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


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


Public Property Get Endereco() As String
    Endereco = mEndereco
End Property
Public Property Let Endereco(ByVal Valor As String)
    mEndereco = Valor
End Property


Public Property Get Bairro() As String
    Bairro = mBairro
End Property
Public Property Let Bairro(ByVal Valor As String)
    mBairro = Valor
End Property


Public Property Get Cidade() As String
    Cidade = mCidade
End Property
Public Property Let Cidade(ByVal Valor As String)
    mCidade = Valor
End Property


Public Property Get UF() As String
    UF = mUF
End Property
Public Property Let UF(ByVal Valor As String)
    mUF = Valor
End Property


Public Property Get CEP() As String
    CEP = mCEP
End Property
Public Property Let CEP(ByVal Valor As String)
    mCEP = Valor
End Property


Public Property Get Telefone() As String
    Telefone = mTelefone
End Property
Public Property Let Telefone(ByVal Valor As String)
    mTelefone = Valor
End Property


Public Property Get Telefone2() As String
    Telefone2 = mTelefone2
End Property
Public Property Let Telefone2(ByVal Valor As String)
    mTelefone2 = Valor
End Property


Public Property Get Fax() As String
    Fax = mFax
End Property
Public Property Let Fax(ByVal Valor As String)
    mFax = Valor
End Property


Public Property Get Vendedor() As String
    Vendedor = mVendedor
End Property
Public Property Let Vendedor(ByVal Valor As String)
    mVendedor = Valor
End Property


Public Property Get CGC() As String
    CGC = mCGC
End Property
Public Property Let CGC(ByVal Valor As String)
    mCGC = Valor
End Property


Public Property Get InscricaoEstadual() As String
    InscricaoEstadual = mInscricaoEstadual
End Property
Public Property Let InscricaoEstadual(ByVal Valor As String)
    mInscricaoEstadual = Valor
End Property


Public Property Get CodigoConta() As Integer
    CodigoConta = mCodigoConta
End Property
Public Property Let CodigoConta(ByVal Valor As Integer)
    mCodigoConta = Valor
End Property


Public Property Get ContaContabil() As String
    ContaContabil = mContaContabil
End Property
Public Property Let ContaContabil(ByVal Valor As String)
    mContaContabil = Valor
End Property


Public Property Get DataAlteracao() As Date
    DataAlteracao = mDataAlteracao
End Property
Public Property Let DataAlteracao(ByVal Valor As Date)
    mDataAlteracao = Valor
End Property


Public Property Get Inativo() As Boolean
    Inativo = mInativo
End Property
Public Property Let Inativo(ByVal Valor As Boolean)
    mInativo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nEmpresa As Integer, ByVal nNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(nNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Fornecedor ( Empresa, Codigo, Nome, Endereco, Bairro, "
    gSQL = gSQL & "Cidade, UF, CEP, Telefone, [Telefone 2], "
    gSQL = gSQL & "Fax, Vendedor, CGC, [Inscricao Estadual], [Codigo da Conta], "
    gSQL = gSQL & "[Conta Contabil], [Data da Alteracao], Inativo ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mEndereco, ", ")
    Call sqlTexto(1, mBairro, ", ")
    Call sqlTexto(1, mCidade, ", ")
    Call sqlTexto(1, mUF, ", ")
    Call sqlTexto(1, mCEP, ", ")
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlTexto(1, mTelefone2, ", ")
    Call sqlTexto(1, mFax, ", ")
    Call sqlTexto(1, mVendedor, ", ")
    Call sqlTexto(1, mCGC, ", ")
    Call sqlTexto(1, mInscricaoEstadual, ", ")
    Call sqlNumero(1, mCodigoConta, ", ")
    Call sqlTexto(1, mContaContabil, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlBoolean(1, mInativo, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Fornecedor SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", Endereco = ", mEndereco)
    Call sqlTexto(2, ", Bairro = ", mBairro)
    Call sqlTexto(2, ", Cidade = ", mCidade)
    Call sqlTexto(2, ", UF = ", mUF)
    Call sqlTexto(2, ", CEP = ", mCEP)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlTexto(2, ", [Telefone 2] = ", mTelefone2)
    Call sqlTexto(2, ", Fax = ", mFax)
    Call sqlTexto(2, ", Vendedor = ", mVendedor)
    Call sqlTexto(2, ", CGC = ", mCGC)
    Call sqlTexto(2, ", [Inscricao Estadual] = ", mInscricaoEstadual)
    Call sqlNumero(2, ", [Codigo da Conta] = ", mCodigoConta)
    Call sqlTexto(2, ", [Conta Contabil] = ", mContaContabil)
    Call sqlData(2, ", [Data da Acerto] = ", mDataAlteracao)
    Call sqlBoolean(2, ", Inativo = ", mInativo)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Fornecedor"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, Nome, Endereco, Bairro, Cidade, UF, CEP, Telefone, [Telefone 2], Fax, Vendedor, CGC, [Inscricao Estadual], [Codigo da Conta], [Conta Contabil], [Data da Alteracao], Inativo FROM Fornecedor" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsFornecedor = New adodb.Recordset
    rsFornecedor.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsFornecedor.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsFornecedor.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsFornecedor("Empresa").Value
    Codigo = rsFornecedor("Codigo").Value
    Nome = rsFornecedor("Nome").Value
    Endereco = rsFornecedor("Endereco").Value
    Bairro = rsFornecedor("Bairro").Value
    Cidade = rsFornecedor("Cidade").Value
    UF = rsFornecedor("UF").Value
    CEP = rsFornecedor("CEP").Value
    Telefone = rsFornecedor("Telefone").Value
    Telefone2 = rsFornecedor("Telefone 2").Value
    Fax = rsFornecedor("Fax").Value
    Vendedor = rsFornecedor("Vendedor").Value
    CGC = rsFornecedor("CGC").Value
    InscricaoEstadual = rsFornecedor("Inscricao Estadual").Value
    CodigoConta = rsFornecedor("Codigo da Conta").Value
    ContaContabil = rsFornecedor("Conta Contabil").Value
    If IsNull(rsFornecedor("Data da Alteracao").Value) Then
        DataAlteracao = "00:00:00"
    Else
        DataAlteracao = rsFornecedor("Data da Alteracao").Value
    End If
    Inativo = rsFornecedor("Inativo").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsFornecedor.Close
    Set rsFornecedor = Nothing
End Function

