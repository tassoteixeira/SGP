VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFuncionario"
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
Private mCPF As String
Private mIdentidade As String
Private mOrgaoEmissor As String
Private mTelefone As String
Private mCargo As String
Private mSituacao As String
Private mPessoalBairro1 As String
Private mPessoalCidade1 As String
Private mPessoalEndereco1 As String
Private mPessoalNome1 As String
Private mPessoalTelefone1 As String
Private mPessoalGrauRelacionamento1 As String
Private mPessoalNome2 As String
Private mPessoalEndereco2 As String
Private mPessoalBairro2 As String
Private mPessoalCidade2 As String
Private mPessoalTelefone2 As String
Private mPessoalGrauRelacionamento2 As String
Private mPessoalNome3 As String
Private mPessoalEndereco3 As String
Private mPessoalBairro3 As String
Private mPessoalCidade3 As String
Private mPessoalTelefone3 As String
Private mPessoalGrauRelacionamento3 As String
Private mPIS As String
Private mCarteiraTrabalho As String
Private mSerieCarteiraTrabalho As String
Private mCNH As String
Private mTituloEleitor As String
Private mNumeroReservista As String
Private mSexo As String
Private mGrauInstrucao As String
Private mDataNascimento As Date
Private mDataAdmissao As Date
Private mDataDemissao As Date
Private mEstadoCivil As String
Private mDependente As Integer
Private mSalarioBase As Currency
Private mAdicionalPericulosidade As Boolean
Private mAdicionalNoturno As Boolean
Private mCestaBasica As Boolean
Private mValeTransporte As Boolean
Private mSeguroVida As Boolean
Private mPeriodo As Integer
Private mSenha As String
Private mCodigoUsuario As Integer

Private rsFuncionario As New adodb.Recordset


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


Public Property Get CPF() As String
    CPF = mCPF
End Property
Public Property Let CPF(ByVal Valor As String)
    mCPF = Valor
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


Public Property Get Telefone() As String
    Telefone = mTelefone
End Property
Public Property Let Telefone(ByVal Valor As String)
    mTelefone = Valor
End Property


Public Property Get Cargo() As String
    Cargo = mCargo
End Property
Public Property Let Cargo(ByVal Valor As String)
    mCargo = Valor
End Property


Public Property Get Situacao() As String
    Situacao = mSituacao
End Property
Public Property Let Situacao(ByVal Valor As String)
    mSituacao = Valor
End Property


Public Property Get PessoalBairro1() As String
    PessoalBairro1 = mPessoalBairro1
End Property
Public Property Let PessoalBairro1(ByVal Valor As String)
    mPessoalBairro1 = Valor
End Property


Public Property Get PessoalCidade1() As String
    PessoalCidade1 = mPessoalCidade1
End Property
Public Property Let PessoalCidade1(ByVal Valor As String)
    mPessoalCidade1 = Valor
End Property


Public Property Get PessoalEndereco1() As String
    PessoalEndereco1 = mPessoalEndereco1
End Property
Public Property Let PessoalEndereco1(ByVal Valor As String)
    mPessoalEndereco1 = Valor
End Property


Public Property Get PessoalNome1() As String
    PessoalNome1 = mPessoalNome1
End Property
Public Property Let PessoalNome1(ByVal Valor As String)
    mPessoalNome1 = Valor
End Property


Public Property Get PessoalTelefone1() As String
    PessoalTelefone1 = mPessoalTelefone1
End Property
Public Property Let PessoalTelefone1(ByVal Valor As String)
    mPessoalTelefone1 = Valor
End Property


Public Property Get PessoalGrauRelacionamento1() As String
    PessoalGrauRelacionamento1 = mPessoalGrauRelacionamento1
End Property
Public Property Let PessoalGrauRelacionamento1(ByVal Valor As String)
    mPessoalGrauRelacionamento1 = Valor
End Property


Public Property Get PessoalNome2() As String
    PessoalNome2 = mPessoalNome2
End Property
Public Property Let PessoalNome2(ByVal Valor As String)
    mPessoalNome2 = Valor
End Property


Public Property Get PessoalEndereco2() As String
    PessoalEndereco2 = mPessoalEndereco2
End Property
Public Property Let PessoalEndereco2(ByVal Valor As String)
    mPessoalEndereco2 = Valor
End Property


Public Property Get PessoalBairro2() As String
    PessoalBairro2 = mPessoalBairro2
End Property
Public Property Let PessoalBairro2(ByVal Valor As String)
    mPessoalBairro2 = Valor
End Property


Public Property Get PessoalCidade2() As String
    PessoalCidade2 = mPessoalCidade2
End Property
Public Property Let PessoalCidade2(ByVal Valor As String)
    mPessoalCidade2 = Valor
End Property


Public Property Get PessoalTelefone2() As String
    PessoalTelefone2 = mPessoalTelefone2
End Property
Public Property Let PessoalTelefone2(ByVal Valor As String)
    mPessoalTelefone2 = Valor
End Property


Public Property Get PessoalGrauRelacionamento2() As String
    PessoalGrauRelacionamento2 = mPessoalGrauRelacionamento2
End Property
Public Property Let PessoalGrauRelacionamento2(ByVal Valor As String)
    mPessoalGrauRelacionamento2 = Valor
End Property


Public Property Get PessoalNome3() As String
    PessoalNome3 = mPessoalNome3
End Property
Public Property Let PessoalNome3(ByVal Valor As String)
    mPessoalNome3 = Valor
End Property


Public Property Get PessoalEndereco3() As String
    PessoalEndereco3 = mPessoalEndereco3
End Property
Public Property Let PessoalEndereco3(ByVal Valor As String)
    mPessoalEndereco3 = Valor
End Property


Public Property Get PessoalBairro3() As String
    PessoalBairro3 = mPessoalBairro3
End Property
Public Property Let PessoalBairro3(ByVal Valor As String)
    mPessoalBairro3 = Valor
End Property


Public Property Get PessoalCidade3() As String
    PessoalCidade3 = mPessoalCidade3
End Property
Public Property Let PessoalCidade3(ByVal Valor As String)
    mPessoalCidade3 = Valor
End Property


Public Property Get PessoalTelefone3() As String
    PessoalTelefone3 = mPessoalTelefone3
End Property
Public Property Let PessoalTelefone3(ByVal Valor As String)
    mPessoalTelefone3 = Valor
End Property


Public Property Get PessoalGrauRelacionamento3() As String
    PessoalGrauRelacionamento3 = mPessoalGrauRelacionamento3
End Property
Public Property Let PessoalGrauRelacionamento3(ByVal Valor As String)
    mPessoalGrauRelacionamento3 = Valor
End Property


Public Property Get PIS() As String
    PIS = mPIS
End Property
Public Property Let PIS(ByVal Valor As String)
    mPIS = Valor
End Property


Public Property Get CarteiraTrabalho() As String
    CarteiraTrabalho = mCarteiraTrabalho
End Property
Public Property Let CarteiraTrabalho(ByVal Valor As String)
    mCarteiraTrabalho = Valor
End Property


Public Property Get SerieCarteiraTrabalho() As String
    SerieCarteiraTrabalho = mSerieCarteiraTrabalho
End Property
Public Property Let SerieCarteiraTrabalho(ByVal Valor As String)
    mSerieCarteiraTrabalho = Valor
End Property


Public Property Get CNH() As String
    CNH = mCNH
End Property
Public Property Let CNH(ByVal Valor As String)
    mCNH = Valor
End Property


Public Property Get TituloEleitor() As String
    TituloEleitor = mTituloEleitor
End Property
Public Property Let TituloEleitor(ByVal Valor As String)
    mTituloEleitor = Valor
End Property


Public Property Get NumeroReservista() As String
    NumeroReservista = mNumeroReservista
End Property
Public Property Let NumeroReservista(ByVal Valor As String)
    mNumeroReservista = Valor
End Property


Public Property Get Sexo() As String
    Sexo = mSexo
End Property
Public Property Let Sexo(ByVal Valor As String)
    mSexo = Valor
End Property


Public Property Get GrauInstrucao() As String
    GrauInstrucao = mGrauInstrucao
End Property
Public Property Let GrauInstrucao(ByVal Valor As String)
    mGrauInstrucao = Valor
End Property


Public Property Get DataNascimento() As Date
    DataNascimento = mDataNascimento
End Property
Public Property Let DataNascimento(ByVal Valor As Date)
    mDataNascimento = Valor
End Property


Public Property Get DataAdmissao() As Date
    DataAdmissao = mDataAdmissao
End Property
Public Property Let DataAdmissao(ByVal Valor As Date)
    mDataAdmissao = Valor
End Property


Public Property Get DataDemissao() As Date
    DataDemissao = mDataDemissao
End Property
Public Property Let DataDemissao(ByVal Valor As Date)
    mDataDemissao = Valor
End Property


Public Property Get EstadoCivil() As String
    EstadoCivil = mEstadoCivil
End Property
Public Property Let EstadoCivil(ByVal Valor As String)
    mEstadoCivil = Valor
End Property


Public Property Get Dependente() As Integer
    Dependente = mDependente
End Property
Public Property Let Dependente(ByVal Valor As Integer)
    mDependente = Valor
End Property


Public Property Get SalarioBase() As Currency
    SalarioBase = mSalarioBase
End Property
Public Property Let SalarioBase(ByVal Valor As Currency)
    mSalarioBase = Valor
End Property


Public Property Get AdicionalPericulosidade() As Boolean
    AdicionalPericulosidade = mAdicionalPericulosidade
End Property
Public Property Let AdicionalPericulosidade(ByVal Valor As Boolean)
    mAdicionalPericulosidade = Valor
End Property


Public Property Get AdicionalNoturno() As Boolean
    AdicionalNoturno = mAdicionalNoturno
End Property
Public Property Let AdicionalNoturno(ByVal Valor As Boolean)
    mAdicionalNoturno = Valor
End Property


Public Property Get CestaBasica() As Boolean
    CestaBasica = mCestaBasica
End Property
Public Property Let CestaBasica(ByVal Valor As Boolean)
    mCestaBasica = Valor
End Property


Public Property Get ValeTransporte() As Boolean
    ValeTransporte = mValeTransporte
End Property
Public Property Let ValeTransporte(ByVal Valor As Boolean)
    mValeTransporte = Valor
End Property


Public Property Get SeguroVida() As Boolean
    SeguroVida = mSeguroVida
End Property
Public Property Let SeguroVida(ByVal Valor As Boolean)
    mSeguroVida = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get Senha() As String
    Senha = mSenha
End Property
Public Property Let Senha(ByVal Valor As String)
    mSenha = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
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
    Call CriaLogCadastroDll2("cFuncionario:LocalizarCodigo Erro.", Err.Description, gSQL)
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


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
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

Public Function LocalizarFuncionarioDoUsuario(ByVal pCodigoUsuario As Integer, ByVal pCodigoEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarFuncionarioDoUsuario = False
    xCondicao = " WHERE Empresa = " & pCodigoEmpresa
    xCondicao = xCondicao & " AND Codigo > 0"
    xCondicao = xCondicao & " AND [Codigo do Usuario] = " & pCodigoUsuario
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarFuncionarioDoUsuario = Localizar
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
    gSQL = "INSERT INTO Funcionario ( Empresa, Codigo, Nome, Endereco, Bairro, "
    gSQL = gSQL & "Cidade, UF, CEP, CPF, Identidade, "
    gSQL = gSQL & "[Orgao Emissor], Telefone, Cargo, Situacao, [Pessoal Bairro 1], "
    gSQL = gSQL & "[Pessoal Cidade 1], [Pessoal Endereco 1], [Pessoal Nome 1], [Pessoal Telefone 1], [Pessoal Grau Relacionamento 1], "
    gSQL = gSQL & "[Pessoal Nome 2], [Pessoal Endereco 2], [Pessoal Bairro 2], [Pessoal Cidade 2], [Pessoal Telefone 2], "
    gSQL = gSQL & "[Pessoal Grau Relacionamento 2], [Pessoal Nome 3], [Pessoal Endereco 3], [Pessoal Bairro 3], [Pessoal Cidade 3], "
    gSQL = gSQL & "[Pessoal Telefone 3], [Pessoal Grau Relacionamento 3], PIS, [Carteira de Trabalho], [Serie da Carteira de Trabalho], "
    gSQL = gSQL & "CNH, [Titulo de Eleitor], [Numero da Reservista], Sexo, [Grau de Instrucao], "
    gSQL = gSQL & "[Data de Nascimento], [Data de Admissao], [Data de Demissao], [Estado Civil], Dependente, "
    gSQL = gSQL & "[Salario Base], [Adicional de Periculosidade], [Adicional Noturno], [Cesta Basica], [Vale Transporte], "
    gSQL = gSQL & "[Seguro de Vida], Periodo, Senha, [Codigo do Usuario] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mEndereco, ", ")
    Call sqlTexto(1, mBairro, ", ")
    Call sqlTexto(1, mCidade, ", ")
    Call sqlTexto(1, mUF, ", ")
    Call sqlTexto(1, mCEP, ", ")
    Call sqlTexto(1, mCPF, ", ")
    Call sqlTexto(1, mIdentidade, ", ")
    Call sqlTexto(1, mOrgaoEmissor, ", ")
    Call sqlTexto(1, mTelefone, ", ")
    Call sqlTexto(1, mCargo, ", ")
    Call sqlTexto(1, mSituacao, ", ")
    Call sqlTexto(1, mPessoalBairro1, ", ")
    Call sqlTexto(1, mPessoalCidade1, ", ")
    Call sqlTexto(1, mPessoalEndereco1, ", ")
    Call sqlTexto(1, mPessoalNome1, ", ")
    Call sqlTexto(1, mPessoalTelefone1, ", ")
    Call sqlTexto(1, mPessoalGrauRelacionamento1, ", ")
    Call sqlTexto(1, mPessoalNome2, ", ")
    Call sqlTexto(1, mPessoalEndereco2, ", ")
    Call sqlTexto(1, mPessoalBairro2, ", ")
    Call sqlTexto(1, mPessoalCidade2, ", ")
    Call sqlTexto(1, mPessoalTelefone2, ", ")
    Call sqlTexto(1, mPessoalGrauRelacionamento2, ", ")
    Call sqlTexto(1, mPessoalNome3, ", ")
    Call sqlTexto(1, mPessoalEndereco3, ", ")
    Call sqlTexto(1, mPessoalBairro3, ", ")
    Call sqlTexto(1, mPessoalCidade3, ", ")
    Call sqlTexto(1, mPessoalTelefone3, ", ")
    Call sqlTexto(1, mPessoalGrauRelacionamento3, ", ")
    Call sqlTexto(1, mPIS, ", ")
    Call sqlTexto(1, mCarteiraTrabalho, ", ")
    Call sqlTexto(1, mSerieCarteiraTrabalho, ", ")
    Call sqlTexto(1, mCNH, ", ")
    Call sqlTexto(1, mTituloEleitor, ", ")
    Call sqlTexto(1, mNumeroReservista, ", ")
    Call sqlTexto(1, mSexo, ", ")
    Call sqlTexto(1, mGrauInstrucao, ", ")
    Call sqlData(1, mDataNascimento, ", ")
    Call sqlData(1, mDataAdmissao, ", ")
    Call sqlData(1, mDataDemissao, ", ")
    Call sqlTexto(1, mEstadoCivil, ", ")
    Call sqlNumero(1, mDependente, ", ")
    Call sqlValor(1, mSalarioBase, ", ")
    Call sqlBoolean(1, mAdicionalPericulosidade, ", ")
    Call sqlBoolean(1, mAdicionalNoturno, ", ")
    Call sqlBoolean(1, mCestaBasica, ", ")
    Call sqlBoolean(1, mValeTransporte, ", ")
    Call sqlBoolean(1, mSeguroVida, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlTexto(1, mSenha, ", ")
    Call sqlNumero(1, mCodigoUsuario, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Funcionario SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", Endereco = ", mEndereco)
    Call sqlTexto(2, ", Bairro = ", mBairro)
    Call sqlTexto(2, ", Cidade = ", mCidade)
    Call sqlTexto(2, ", UF = ", mUF)
    Call sqlTexto(2, ", CEP = ", mCEP)
    Call sqlTexto(2, ", CPF = ", mCPF)
    Call sqlTexto(2, ", Identidade = ", mIdentidade)
    Call sqlTexto(2, ", [Orgao Emissor] = ", mOrgaoEmissor)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    Call sqlTexto(2, ", Cargo = ", mCargo)
    Call sqlTexto(2, ", Situacao = ", mSituacao)
    Call sqlTexto(2, ", [Pessoal Bairro 1] = ", mPessoalBairro1)
    Call sqlTexto(2, ", [Pessoal Cidade 1] = ", mPessoalCidade1)
    Call sqlTexto(2, ", [Pessoal Endereco 1] = ", mPessoalEndereco1)
    Call sqlTexto(2, ", [Pessoal Nome 1] = ", mPessoalNome1)
    Call sqlTexto(2, ", [Pessoal Telefone 1] = ", mPessoalTelefone1)
    Call sqlTexto(2, ", [Pessoal Grau Relacionamento 1] = ", mPessoalGrauRelacionamento1)
    Call sqlTexto(2, ", [Pessoal Nome 2] = ", mPessoalNome2)
    Call sqlTexto(2, ", [Pessoal Endereco 2] = ", mPessoalEndereco2)
    Call sqlTexto(2, ", [Pessoal Bairro 2] = ", mPessoalBairro2)
    Call sqlTexto(2, ", [Pessoal Cidade 2] = ", mPessoalCidade2)
    Call sqlTexto(2, ", [Pessoal Telefone 2] = ", mPessoalTelefone2)
    Call sqlTexto(2, ", [Pessoal Grau Relacionamento 2] = ", mPessoalGrauRelacionamento2)
    Call sqlTexto(2, ", [Pessoal Nome 3] = ", mPessoalNome3)
    Call sqlTexto(2, ", [Pessoal Endereco 3] = ", mPessoalEndereco3)
    Call sqlTexto(2, ", [Pessoal Bairro 3] = ", mPessoalBairro3)
    Call sqlTexto(2, ", [Pessoal Cidade 3] = ", mPessoalCidade3)
    Call sqlTexto(2, ", [Pessoal Telefone 3] = ", mPessoalTelefone3)
    Call sqlTexto(2, ", [Pessoal Grau Relacionamento 3] = ", mPessoalGrauRelacionamento3)
    Call sqlTexto(2, ", PIS = ", mPIS)
    Call sqlTexto(2, ", [Carteira de Trabalho] = ", mCarteiraTrabalho)
    Call sqlTexto(2, ", [Serie da Carteira de Trabalho] = ", mSerieCarteiraTrabalho)
    Call sqlTexto(2, ", CNH = ", mCNH)
    Call sqlTexto(2, ", [Titulo de Eleitor] = ", mTituloEleitor)
    Call sqlTexto(2, ", [Numero da Reservista] = ", mNumeroReservista)
    Call sqlTexto(2, ", Sexo = ", mSexo)
    Call sqlTexto(2, ", [Grau de Instrucao] = ", mGrauInstrucao)
    Call sqlData(2, ", [Data de Nascimento] = ", mDataNascimento)
    Call sqlData(2, ", [Data de Admissao] = ", mDataAdmissao)
    Call sqlData(2, ", [Data de Demissao] = ", mDataDemissao)
    Call sqlTexto(2, ", [Estado Civil] = ", mEstadoCivil)
    Call sqlNumero(2, ", Dependente = ", mDependente)
    Call sqlValor(2, ", [Salario Base] = ", mSalarioBase)
    Call sqlBoolean(2, ", [Adicional de Periculosidade] = ", mAdicionalPericulosidade)
    Call sqlBoolean(2, ", [Adicional Noturno] = ", mAdicionalNoturno)
    Call sqlBoolean(2, ", [Cesta Basica] = ", mCestaBasica)
    Call sqlBoolean(2, ", [Vale Transporte] = ", mValeTransporte)
    Call sqlBoolean(2, ", [Seguro de Vida] = ", mSeguroVida)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", Senha = ", mSenha)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
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
    gSQL = "DELETE FROM Funcionario"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, Nome, Endereco, Bairro, Cidade, UF, CEP, CPF, Identidade, [Orgao Emissor], Telefone, Cargo, Situacao, [Pessoal Bairro 1], [Pessoal Cidade 1], [Pessoal Endereco 1], [Pessoal Nome 1], [Pessoal Telefone 1], [Pessoal Grau Relacionamento 1], [Pessoal Nome 2], [Pessoal Endereco 2], [Pessoal Bairro 2], [Pessoal Cidade 2], [Pessoal Telefone 2], [Pessoal Grau Relacionamento 2], [Pessoal Nome 3], [Pessoal Endereco 3], [Pessoal Bairro 3], [Pessoal Cidade 3], [Pessoal Telefone 3], [Pessoal Grau Relacionamento 3], PIS, [Carteira de Trabalho], [Serie da Carteira de Trabalho], CNH, [Titulo de Eleitor], [Numero da Reservista], Sexo, [Grau de Instrucao], [Data de Nascimento], [Data de Admissao], [Data de Demissao], [Estado Civil], Dependente, [Salario Base], [Adicional de Periculosidade], [Adicional Noturno], [Cesta Basica], [Vale Transporte], [Seguro de Vida], Periodo, Senha, [Codigo do Usuario] FROM Funcionario" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsFuncionario = New adodb.Recordset
    rsFuncionario.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsFuncionario.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsFuncionario.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsFuncionario("Empresa").Value
    Codigo = rsFuncionario("Codigo").Value
    Nome = rsFuncionario("Nome").Value
    Endereco = rsFuncionario("Endereco").Value
    Bairro = rsFuncionario("Bairro").Value
    Cidade = rsFuncionario("Cidade").Value
    UF = rsFuncionario("UF").Value
    CEP = rsFuncionario("CEP").Value
    CPF = rsFuncionario("CPF").Value
    Identidade = rsFuncionario("Identidade").Value
    OrgaoEmissor = rsFuncionario("Orgao Emissor").Value
    Telefone = rsFuncionario("Telefone").Value
    Cargo = rsFuncionario("Cargo").Value
    Situacao = rsFuncionario("Situacao").Value
    PessoalBairro1 = rsFuncionario("Pessoal Bairro 1").Value
    PessoalCidade1 = rsFuncionario("Pessoal Cidade 1").Value
    PessoalEndereco1 = rsFuncionario("Pessoal Endereco 1").Value
    PessoalNome1 = rsFuncionario("Pessoal Nome 1").Value
    PessoalTelefone1 = rsFuncionario("Pessoal Telefone 1").Value
    PessoalGrauRelacionamento1 = rsFuncionario("Pessoal Grau Relacionamento 1").Value
    PessoalNome2 = rsFuncionario("Pessoal Nome 2").Value
    PessoalEndereco2 = rsFuncionario("Pessoal Endereco 2").Value
    PessoalBairro2 = rsFuncionario("Pessoal Bairro 2").Value
    PessoalCidade2 = rsFuncionario("Pessoal Cidade 2").Value
    PessoalTelefone2 = rsFuncionario("Pessoal Telefone 2").Value
    PessoalGrauRelacionamento2 = rsFuncionario("Pessoal Grau Relacionamento 2").Value
    PessoalNome3 = rsFuncionario("Pessoal Nome 3").Value
    PessoalEndereco3 = rsFuncionario("Pessoal Endereco 3").Value
    PessoalBairro3 = rsFuncionario("Pessoal Bairro 3").Value
    PessoalCidade3 = rsFuncionario("Pessoal Cidade 3").Value
    PessoalTelefone3 = rsFuncionario("Pessoal Telefone 3").Value
    PessoalGrauRelacionamento3 = rsFuncionario("Pessoal Grau Relacionamento 3").Value
    PIS = rsFuncionario("PIS").Value
    CarteiraTrabalho = rsFuncionario("Carteira de Trabalho").Value
    SerieCarteiraTrabalho = rsFuncionario("Serie da Carteira de Trabalho").Value
    CNH = rsFuncionario("CNH").Value
    TituloEleitor = rsFuncionario("Titulo de Eleitor").Value
    NumeroReservista = rsFuncionario("Numero da Reservista").Value
    Sexo = rsFuncionario("Sexo").Value
    GrauInstrucao = rsFuncionario("Grau de Instrucao").Value
    If IsNull(rsFuncionario("Data de Nascimento").Value) Then
        DataNascimento = "00:00:00"
    Else
        DataNascimento = rsFuncionario("Data de Nascimento").Value
    End If
    If IsNull(rsFuncionario("Data de Admissao").Value) Then
        DataAdmissao = "00:00:00"
    Else
        DataAdmissao = rsFuncionario("Data de Admissao").Value
    End If
    If IsNull(rsFuncionario("Data de Demissao").Value) Then
        DataDemissao = "00:00:00"
    Else
        DataDemissao = rsFuncionario("Data de Demissao").Value
    End If
    EstadoCivil = rsFuncionario("Estado Civil").Value
    Dependente = rsFuncionario("Dependente").Value
    SalarioBase = rsFuncionario("Salario Base").Value
    AdicionalPericulosidade = rsFuncionario("Adicional de Periculosidade").Value
    AdicionalNoturno = rsFuncionario("Adicional Noturno").Value
    CestaBasica = rsFuncionario("Cesta Basica").Value
    ValeTransporte = rsFuncionario("Vale Transporte").Value
    SeguroVida = rsFuncionario("Seguro de Vida").Value
    Periodo = rsFuncionario("Periodo").Value
    Senha = rsFuncionario("Senha").Value
    CodigoUsuario = rsFuncionario("Codigo do Usuario").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsFuncionario.Close
    Set rsFuncionario = Nothing
End Function

