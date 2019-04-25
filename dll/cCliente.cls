VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Long
Private mRazaoSocial As String
Private mEndereco As String
Private mBairro As String
Private mCidade As String
Private mUF As String
Private mCEP As String
Private mTelefone As String
Private mCPF As String
Private mCGC As String
Private mIdentidade As String
Private mOrgaoEmissor As String
Private mInscricaoEstadual As String
Private mCodigoVencimento As Integer
Private mCodigoConvenio As Integer
Private mEmpresa As Integer
Private mInativo As Boolean
Private mImprimeDuplicata As Boolean
Private mEmail As String
Private mFormaPagamento As String
Private mGeraNotaAbastecimento As Boolean
Private mImprimeDadosECF As Boolean
Private mCodigoGrupoCliente As Integer
Private mNomeConjuge As String
Private mCPFConjuge As String
Private mNomeMae As String
Private mTelefone2 As String
Private mFax As String
Private mCelular As String
Private mDataAlteracao As Date
Private mDescontoEspecial As Boolean
Private mContribuinte As Boolean

Private rsCliente As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Long)
    mCodigo = Valor
End Property


Public Property Get RazaoSocial() As String
    RazaoSocial = mRazaoSocial
End Property
Public Property Let RazaoSocial(ByVal Valor As String)
    mRazaoSocial = Valor
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


Public Property Get CPF() As String
    CPF = mCPF
End Property
Public Property Let CPF(ByVal Valor As String)
    mCPF = Valor
End Property


Public Property Get CGC() As String
    CGC = mCGC
End Property
Public Property Let CGC(ByVal Valor As String)
    mCGC = Valor
End Property


Public Property Get Identidade() As String
    Identidade = mIdentidade
End Property
Public Property Let Identidade(ByVal Valor As String)
    mIdentidade = Valor
End Property


Public Property Get OrgaoEmissor() As String
    OrgaoEmissor = mOrgaoEmissor
End Property
Public Property Let OrgaoEmissor(ByVal Valor As String)
    mOrgaoEmissor = Valor
End Property


Public Property Get InscricaoEstadual() As String
    InscricaoEstadual = mInscricaoEstadual
End Property
Public Property Let InscricaoEstadual(ByVal Valor As String)
    mInscricaoEstadual = Valor
End Property


Public Property Get CodigoVencimento() As Integer
    CodigoVencimento = mCodigoVencimento
End Property
Public Property Let CodigoVencimento(ByVal Valor As Integer)
    mCodigoVencimento = Valor
End Property


Public Property Get CodigoConvenio() As Integer
    CodigoConvenio = mCodigoConvenio
End Property
Public Property Let CodigoConvenio(ByVal Valor As Integer)
    mCodigoConvenio = Valor
End Property


Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Inativo() As Boolean
    Inativo = mInativo
End Property
Public Property Let Inativo(ByVal Valor As Boolean)
    mInativo = Valor
End Property


Public Property Get ImprimeDuplicata() As Boolean
    ImprimeDuplicata = mImprimeDuplicata
End Property
Public Property Let ImprimeDuplicata(ByVal Valor As Boolean)
    mImprimeDuplicata = Valor
End Property


Public Property Get Email() As String
    Email = mEmail
End Property
Public Property Let Email(ByVal Valor As String)
    mEmail = Valor
End Property


Public Property Get FormaPagamento() As String
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As String)
    mFormaPagamento = Valor
End Property


Public Property Get GeraNotaAbastecimento() As Boolean
    GeraNotaAbastecimento = mGeraNotaAbastecimento
End Property
Public Property Let GeraNotaAbastecimento(ByVal Valor As Boolean)
    mGeraNotaAbastecimento = Valor
End Property


Public Property Get ImprimeDadosECF() As Boolean
    ImprimeDadosECF = mImprimeDadosECF
End Property
Public Property Let ImprimeDadosECF(ByVal Valor As Boolean)
    mImprimeDadosECF = Valor
End Property


Public Property Get CodigoGrupoCliente() As Integer
    CodigoGrupoCliente = mCodigoGrupoCliente
End Property
Public Property Let CodigoGrupoCliente(ByVal Valor As Integer)
    mCodigoGrupoCliente = Valor
End Property


Public Property Get NomeConjuge() As String
    NomeConjuge = mNomeConjuge
End Property
Public Property Let NomeConjuge(ByVal Valor As String)
    mNomeConjuge = Valor
End Property


Public Property Get CPFConjuge() As String
    CPFConjuge = mCPFConjuge
End Property
Public Property Let CPFConjuge(ByVal Valor As String)
    mCPFConjuge = Valor
End Property


Public Property Get NomeMae() As String
    NomeMae = mNomeMae
End Property
Public Property Let NomeMae(ByVal Valor As String)
    mNomeMae = Valor
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


Public Property Get Celular() As String
    Celular = mCelular
End Property
Public Property Let Celular(ByVal Valor As String)
    mCelular = Valor
End Property


Public Property Get DataAlteracao() As Date
    DataAlteracao = mDataAlteracao
End Property
Public Property Let DataAlteracao(ByVal Valor As Date)
    mDataAlteracao = Valor
End Property


Public Property Get DescontoEspecial() As Boolean
    DescontoEspecial = mDescontoEspecial
End Property
Public Property Let DescontoEspecial(ByVal Valor As Boolean)
    mDescontoEspecial = Valor
End Property


Public Property Get Contribuinte() As Boolean
    Contribuinte = mContribuinte
End Property
Public Property Let Contribuinte(ByVal Valor As Boolean)
    mContribuinte = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & pCodigo
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
    xCondicao = " WHERE Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE [Razao Social] = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
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
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < " & 99999999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Cliente ( Codigo, [Razao Social], Endereco, Bairro, Cidade, "
    gSQL = gSQL & "UF, CEP, Telefone, CPF, CGC, "
    gSQL = gSQL & "Identidade, [Orgao Emissor], [Inscricao Estadual], [Codigo do Vencimento], [Codigo do Convenio], "
    gSQL = gSQL & "Empresa, Inativo, [Imprime Duplicata], Email, [Forma de Pagamento], [Gera Nota de Abastecimento], "
    gSQL = gSQL & "[Imprime Dados no ECF], [Codigo do Grupo de Cliente], [Nome Conjuge], [CPF Conjuge], "
    gSQL = gSQL & "[Nome da Mae], [Telefone 2], FAX, Celular, [Data da Alteracao], "
    gSQL = gSQL & "[Desconto Especial], Contribuinte ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mRazaoSocial, ", ")
    Call sqlTexto(1, mEndereco, ", ")
    Call sqlTexto(1, mBairro, ", ")
    Call sqlTexto(1, mCidade, ", ")
    Call sqlTexto(1, mUF, ", ")
    Call sqlTexto(1, mCEP, ", ")
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlTexto(1, mCPF, ", ")
    Call sqlTexto(1, mCGC, ", ")
    Call sqlTexto(1, mIdentidade, ", ")
    Call sqlTexto(1, mOrgaoEmissor, ", ")
    Call sqlTexto(1, mInscricaoEstadual, ", ")
    Call sqlNumero(1, mCodigoVencimento, ", ")
    Call sqlNumero(1, mCodigoConvenio, ", ")
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlBoolean(1, mInativo, ", ")
    Call sqlBoolean(1, mImprimeDuplicata, ", ")
    Call sqlTexto(1, mEmail, ", ")
    Call sqlTexto(1, mFormaPagamento, ", ")
    Call sqlBoolean(1, mGeraNotaAbastecimento, ", ")
    Call sqlBoolean(1, mImprimeDadosECF, ", ")
    Call sqlNumero(1, mCodigoGrupoCliente, ", ")
    Call sqlTexto(1, mNomeConjuge, ", ")
    Call sqlTexto(1, mCPFConjuge, ", ")
    Call sqlTexto(1, mNomeMae, ", ")
    Call sqlTexto(1, mTelefone2, ", ")
    Call sqlTexto(1, mFax, ", ")
    Call sqlTexto(1, mCelular, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlBoolean(1, mDescontoEspecial, ", ")
    Call sqlBoolean(1, mContribuinte, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    Else
        Call CriaLogCadastroDll2("Classe:cCliente - Método:Incluir", "Não foi possível incluir registro.", gSQL)
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("Classe:cCliente - Método:Incluir", Err.Description, gSQL)
End Function


Public Function Alterar(ByVal pCodigo As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Cliente SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", [Razao Social] = ", mRazaoSocial)
    Call sqlTexto(2, ", Endereco = ", mEndereco)
    Call sqlTexto(2, ", Bairro = ", mBairro)
    Call sqlTexto(2, ", Cidade = ", mCidade)
    Call sqlTexto(2, ", UF = ", mUF)
    Call sqlTexto(2, ", CEP = ", mCEP)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlTexto(2, ", CPF = ", mCPF)
    Call sqlTexto(2, ", CGC = ", mCGC)
    Call sqlTexto(2, ", Identidade = ", mIdentidade)
    Call sqlTexto(2, ", [Orgao Emissor] = ", mOrgaoEmissor)
    Call sqlTexto(2, ", [Inscricao Estadual] = ", mInscricaoEstadual)
    Call sqlNumero(2, ", [Codigo do Vencimento] = ", mCodigoVencimento)
    Call sqlNumero(2, ", [Codigo do Convenio] = ", mCodigoConvenio)
    Call sqlNumero(2, ", Empresa = ", mEmpresa)
    Call sqlBoolean(2, ", Inativo = ", mInativo)
    Call sqlBoolean(2, ", [Imprime Duplicata] = ", mImprimeDuplicata)
    Call sqlTexto(2, ", Email = ", mEmail)
    Call sqlTexto(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlBoolean(2, ", [Gera Nota de Abastecimento] = ", mGeraNotaAbastecimento)
    Call sqlBoolean(2, ", [Imprime Dados no ECF] = ", mImprimeDadosECF)
    Call sqlNumero(2, ", [Codigo do Grupo de Cliente] = ", mCodigoGrupoCliente)
    Call sqlTexto(2, ", [Nome Conjuge] = ", mNomeConjuge)
    Call sqlTexto(2, ", [CPF Conjuge] = ", mCPFConjuge)
    Call sqlTexto(2, ", [Nome da Mae] = ", mNomeMae)
    Call sqlTexto(2, ", [Telefone 2] = ", mTelefone2)
    Call sqlTexto(2, ", FAX = ", mFax)
    Call sqlTexto(2, ", Celular = ", mCelular)
    Call sqlData(2, ", [Data da Acerto] = ", mDataAlteracao)
    Call sqlBoolean(2, ", [Desconto Especial] = ", mDescontoEspecial)
    Call sqlBoolean(2, ", Contribuinte = ", mContribuinte)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigo As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Cliente"
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ProximoCodigo() As Long
Dim xCondicao As String

On Error GoTo trata_erro

    ProximoCodigo = 1
    xCondicao = "SELECT TOP 1 Codigo FROM Cliente WHERE Codigo < " & 50000 & " ORDER BY Codigo DESC"
    Set rsCliente = New adodb.Recordset
    rsCliente.CursorLocation = adUseClient
    rsCliente.Open xCondicao, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCliente.EOF Then
        ProximoCodigo = rsCliente("Codigo").Value + 1
    End If
    rsCliente.Close
    Set rsCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, [Razao Social], Endereco, Bairro, Cidade, UF, CEP, Telefone, CPF, CGC, Identidade, [Orgao Emissor], [Inscricao Estadual], [Codigo do Vencimento], [Codigo do Convenio], Empresa, Inativo, [Imprime Duplicata], Email, [Forma de Pagamento], [Gera Nota de Abastecimento], [Imprime Dados no ECF], [Codigo do Grupo de Cliente], [Nome Conjuge], [CPF Conjuge], [Nome da Mae], [Telefone 2], FAX, Celular, [Data da Alteracao], [Desconto Especial], Contribuinte FROM Cliente" & " " & pCondicao & " " & pTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    
    AtualizaRecordset = False
    Set rsCliente = New adodb.Recordset
    rsCliente.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCliente.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCliente.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsCliente("Codigo").Value
    RazaoSocial = rsCliente("Razao Social").Value
    Endereco = rsCliente("Endereco").Value
    Bairro = rsCliente("Bairro").Value
    Cidade = rsCliente("Cidade").Value
    UF = rsCliente("UF").Value
    CEP = rsCliente("CEP").Value
    Telefone = rsCliente("Telefone").Value
    CPF = rsCliente("CPF").Value
    CGC = rsCliente("CGC").Value
    Identidade = rsCliente("Identidade").Value
    OrgaoEmissor = rsCliente("Orgao Emissor").Value
    InscricaoEstadual = rsCliente("Inscricao Estadual").Value
    CodigoVencimento = rsCliente("Codigo do Vencimento").Value
    CodigoConvenio = rsCliente("Codigo do Convenio").Value
    Empresa = rsCliente("Empresa").Value
    Inativo = rsCliente("Inativo").Value
    ImprimeDuplicata = rsCliente("Imprime Duplicata").Value
    Email = rsCliente("Email").Value
    FormaPagamento = rsCliente("Forma de Pagamento").Value
    GeraNotaAbastecimento = rsCliente("Gera Nota de Abastecimento").Value
    ImprimeDadosECF = rsCliente("Imprime Dados no ECF").Value
    CodigoGrupoCliente = rsCliente("Codigo do Grupo de Cliente").Value
    NomeConjuge = rsCliente("Nome Conjuge").Value
    CPFConjuge = rsCliente("CPF Conjuge").Value
    NomeMae = rsCliente("Nome da Mae").Value
    Telefone2 = rsCliente("Telefone 2").Value
    Fax = rsCliente("FAX").Value
    Celular = rsCliente("Celular").Value
    If IsNull(rsCliente("Data da Alteracao").Value) Then
        DataAlteracao = "00:00:00"
    Else
        DataAlteracao = rsCliente("Data da Alteracao").Value
    End If
    DescontoEspecial = rsCliente("Desconto Especial").Value
    Contribuinte = rsCliente("Contribuinte").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsCliente.Close
    Set rsCliente = Nothing
End Function

