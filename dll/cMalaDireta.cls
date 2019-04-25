VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNomeRazaoSocial As String
Private mEndereco As String
Private mBairro As String
Private mCidade As String
Private mUnidadeFederativa As String
Private mCEP As String
Private mTelefone1 As String
Private mDataNascimento As Date
Private mCPFCNPJ As String
Private mTelefone2 As String
Private mFax As String
Private mPessoaparaContato As String
Private mEmail As String

Private rsMalaDireta As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get NomeRazaoSocial() As String
    NomeRazaoSocial = mNomeRazaoSocial
End Property
Public Property Let NomeRazaoSocial(ByVal Valor As String)
    mNomeRazaoSocial = Valor
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


Public Property Get UnidadeFederativa() As String
    UnidadeFederativa = mUnidadeFederativa
End Property
Public Property Let UnidadeFederativa(ByVal Valor As String)
    mUnidadeFederativa = Valor
End Property


Public Property Get CEP() As String
    CEP = mCEP
End Property
Public Property Let CEP(ByVal Valor As String)
    mCEP = Valor
End Property


Public Property Get Telefone1() As String
    Telefone1 = mTelefone1
End Property
Public Property Let Telefone1(ByVal Valor As String)
    mTelefone1 = Valor
End Property


Public Property Get DataNascimento() As Date
    DataNascimento = mDataNascimento
End Property
Public Property Let DataNascimento(ByVal Valor As Date)
    mDataNascimento = Valor
End Property


Public Property Get CPFCNPJ() As String
    CPFCNPJ = mCPFCNPJ
End Property
Public Property Let CPFCNPJ(ByVal Valor As String)
    mCPFCNPJ = Valor
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


Public Property Get PessoaparaContato() As String
    PessoaparaContato = mPessoaparaContato
End Property
Public Property Let PessoaparaContato(ByVal Valor As String)
    mPessoaparaContato = Valor
End Property


Public Property Get Email() As String
    Email = mEmail
End Property
Public Property Let Email(ByVal Valor As String)
    mEmail = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
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
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE [Nome Razao Social] = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > " & 0
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
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
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMalaDireta.Close
    Set rsMalaDireta = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Mala_Direta ( Codigo, [Nome Razao Social], Endereco, Bairro, Cidade, "
    gSQL = gSQL & "[Unidade Federativa], CEP, [Telefone 1], [Data de Nascimento], [CPF CNPJ], "
    gSQL = gSQL & "[Telefone 2], Fax, [Pessoa para Contato], Email ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNomeRazaoSocial, ", ")
    Call sqlTexto(1, mEndereco, ", ")
    Call sqlTexto(1, mBairro, ", ")
    Call sqlTexto(1, mCidade, ", ")
    Call sqlTexto(1, mUnidadeFederativa, ", ")
    Call sqlTexto(1, mCEP, ", ")
    Call sqlTexto(1, mTelefone1, ", ")
    Call sqlData(1, mDataNascimento, ", ")
    Call sqlTexto(1, mCPFCNPJ, ", ")
    Call sqlTexto(1, mTelefone2, ", ")
    Call sqlTexto(1, mFax, ", ")
    Call sqlTexto(1, mPessoaparaContato, ", ")
    Call sqlTexto(1, mEmail, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Mala_Direta SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", [Nome Razao Social] = ", mNomeRazaoSocial)
    Call sqlTexto(2, ", Endereco = ", mEndereco)
    Call sqlTexto(2, ", Bairro = ", mBairro)
    Call sqlTexto(2, ", Cidade = ", mCidade)
    Call sqlTexto(2, ", [Unidade Federativa] = ", mUnidadeFederativa)
    Call sqlTexto(2, ", CEP = ", mCEP)
    Call sqlTexto(2, ", [Telefone 1] = ", mTelefone1)
    Call sqlData(2, ", [Data de Nascimento] = ", mDataNascimento)
    Call sqlTexto(2, ", [CPF CNPJ] = ", mCPFCNPJ)
    Call sqlTexto(2, ", [Telefone 2] = ", mTelefone2)
    Call sqlTexto(2, ", Fax = ", mFax)
    Call sqlTexto(2, ", [Pessoa para Contato] = ", mPessoaparaContato)
    Call sqlTexto(2, ", Email = ", mEmail)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Mala_Direta"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, [Nome Razao Social], Endereco, Bairro, Cidade, [Unidade Federativa], CEP, [Telefone 1], [Data de Nascimento], [CPF CNPJ], [Telefone 2], Fax, [Pessoa para Contato], Email FROM Mala_Direta" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMalaDireta = New adodb.Recordset
    rsMalaDireta.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMalaDireta.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMalaDireta.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsMalaDireta("Codigo").Value
    NomeRazaoSocial = rsMalaDireta("Nome Razao Social").Value
    Endereco = rsMalaDireta("Endereco").Value
    Bairro = rsMalaDireta("Bairro").Value
    Cidade = rsMalaDireta("Cidade").Value
    UnidadeFederativa = rsMalaDireta("Unidade Federativa").Value
    CEP = rsMalaDireta("CEP").Value
    Telefone1 = rsMalaDireta("Telefone 1").Value
    If IsNull(rsMalaDireta("Data de Nascimento").Value) Then
        DataNascimento = "00:00:00"
    Else
        DataNascimento = rsMalaDireta("Data de Nascimento").Value
    End If
    CPFCNPJ = rsMalaDireta("CPF CNPJ").Value
    Telefone2 = rsMalaDireta("Telefone 2").Value
    Fax = rsMalaDireta("Fax").Value
    PessoaparaContato = rsMalaDireta("Pessoa para Contato").Value
    Email = rsMalaDireta("Email").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

