VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mQuantidadePeriodos As Integer
Private mQuantidadeBico As Integer
Private mPularBomba As Integer
Private mCustoDuplicata As Currency
Private mValorSuperior As Currency
Private mValorEsquerda As Currency
Private mExtenso1Superior As Currency
Private mExtenso1Esquerda As Currency
Private mExtenso2Superior As Currency
Private mExtenso2Esquerda As Currency
Private mFavorecidoSuperior As Currency
Private mFavorecidoEsquerda As Currency
Private mCidadeSuperior As Currency
Private mCidadeEsquerda As Currency
Private mDiaSuperior As Currency
Private mDiaEsquerda As Currency
Private mMesSuperior As Currency
Private mMesEsquerda As Currency
Private mAnoSuperior As Currency
Private mAnoEsquerda As Currency
Private mOutrasConfiguracoes As String
Private mMensagemCobranca As String
Private mQuantidadeIlha As Integer
Private mProgramacaoAntiga As Boolean
Private mHoraFechamento1 As Date
Private mHoraFechamento2 As Date
Private mHoraFechamento3 As Date
Private mHoraFechamento4 As Date
Private mHoraFechamento5 As Date
Private mHoraFechamento6 As Date
Private mHoraFechamento7 As Date
Private mHoraFechamento8 As Date
Private mImprimirReducaoZ As Boolean
Private mQuantidadeViasTEF As Integer
Private mControleSolicitacaoTEF As Long
Private mIntegraMovimentoBombaCaixa As Boolean
Private mAlteraAberturaBomba As Boolean
Private mAlteraPrecoMovimentoBomba As Boolean
Private mECFBaixaEstoque As Boolean
Private mNomeclaturaCaixa As String
Private mAlteracaoCaixaPeloResponsavel As Boolean
Private mAlteraPrecoProdutoPelaVenda As Boolean
Private mInverteEncerrantenaPlanilha As Boolean
Private mIdentificaFuncionarioaCadaCupom As Boolean
Private mRelacaoNotasnoCaixa As Boolean
Private mBloqueiaVendaPeloEstoque As Boolean
Private mBloqueiaVendaPeloSubEstoque As Boolean
Private mNumeroDuplicata As Long

Private rsConfiguracao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get QuantidadePeriodos() As Integer
    QuantidadePeriodos = mQuantidadePeriodos
End Property
Public Property Let QuantidadePeriodos(ByVal Valor As Integer)
    mQuantidadePeriodos = Valor
End Property


Public Property Get QuantidadeBico() As Integer
    QuantidadeBico = mQuantidadeBico
End Property
Public Property Let QuantidadeBico(ByVal Valor As Integer)
    mQuantidadeBico = Valor
End Property


Public Property Get PularBomba() As Integer
    PularBomba = mPularBomba
End Property
Public Property Let PularBomba(ByVal Valor As Integer)
    mPularBomba = Valor
End Property


Public Property Get CustoDuplicata() As Currency
    CustoDuplicata = mCustoDuplicata
End Property
Public Property Let CustoDuplicata(ByVal Valor As Currency)
    mCustoDuplicata = Valor
End Property


Public Property Get ValorSuperior() As Currency
    ValorSuperior = mValorSuperior
End Property
Public Property Let ValorSuperior(ByVal Valor As Currency)
    mValorSuperior = Valor
End Property


Public Property Get ValorEsquerda() As Currency
    ValorEsquerda = mValorEsquerda
End Property
Public Property Let ValorEsquerda(ByVal Valor As Currency)
    mValorEsquerda = Valor
End Property


Public Property Get Extenso1Superior() As Currency
    Extenso1Superior = mExtenso1Superior
End Property
Public Property Let Extenso1Superior(ByVal Valor As Currency)
    mExtenso1Superior = Valor
End Property


Public Property Get Extenso1Esquerda() As Currency
    Extenso1Esquerda = mExtenso1Esquerda
End Property
Public Property Let Extenso1Esquerda(ByVal Valor As Currency)
    mExtenso1Esquerda = Valor
End Property


Public Property Get Extenso2Superior() As Currency
    Extenso2Superior = mExtenso2Superior
End Property
Public Property Let Extenso2Superior(ByVal Valor As Currency)
    mExtenso2Superior = Valor
End Property


Public Property Get Extenso2Esquerda() As Currency
    Extenso2Esquerda = mExtenso2Esquerda
End Property
Public Property Let Extenso2Esquerda(ByVal Valor As Currency)
    mExtenso2Esquerda = Valor
End Property


Public Property Get FavorecidoSuperior() As Currency
    FavorecidoSuperior = mFavorecidoSuperior
End Property
Public Property Let FavorecidoSuperior(ByVal Valor As Currency)
    mFavorecidoSuperior = Valor
End Property


Public Property Get FavorecidoEsquerda() As Currency
    FavorecidoEsquerda = mFavorecidoEsquerda
End Property
Public Property Let FavorecidoEsquerda(ByVal Valor As Currency)
    mFavorecidoEsquerda = Valor
End Property


Public Property Get CidadeSuperior() As Currency
    CidadeSuperior = mCidadeSuperior
End Property
Public Property Let CidadeSuperior(ByVal Valor As Currency)
    mCidadeSuperior = Valor
End Property


Public Property Get CidadeEsquerda() As Currency
    CidadeEsquerda = mCidadeEsquerda
End Property
Public Property Let CidadeEsquerda(ByVal Valor As Currency)
    mCidadeEsquerda = Valor
End Property


Public Property Get DiaSuperior() As Currency
    DiaSuperior = mDiaSuperior
End Property
Public Property Let DiaSuperior(ByVal Valor As Currency)
    mDiaSuperior = Valor
End Property


Public Property Get DiaEsquerda() As Currency
    DiaEsquerda = mDiaEsquerda
End Property
Public Property Let DiaEsquerda(ByVal Valor As Currency)
    mDiaEsquerda = Valor
End Property


Public Property Get MesSuperior() As Currency
    MesSuperior = mMesSuperior
End Property
Public Property Let MesSuperior(ByVal Valor As Currency)
    mMesSuperior = Valor
End Property


Public Property Get MesEsquerda() As Currency
    MesEsquerda = mMesEsquerda
End Property
Public Property Let MesEsquerda(ByVal Valor As Currency)
    mMesEsquerda = Valor
End Property


Public Property Get AnoSuperior() As Currency
    AnoSuperior = mAnoSuperior
End Property
Public Property Let AnoSuperior(ByVal Valor As Currency)
    mAnoSuperior = Valor
End Property


Public Property Get AnoEsquerda() As Currency
    AnoEsquerda = mAnoEsquerda
End Property
Public Property Let AnoEsquerda(ByVal Valor As Currency)
    mAnoEsquerda = Valor
End Property


Public Property Get OutrasConfiguracoes() As String
    OutrasConfiguracoes = mOutrasConfiguracoes
End Property
Public Property Let OutrasConfiguracoes(ByVal Valor As String)
    mOutrasConfiguracoes = Valor
End Property


Public Property Get MensagemCobranca() As String
    MensagemCobranca = mMensagemCobranca
End Property
Public Property Let MensagemCobranca(ByVal Valor As String)
    mMensagemCobranca = Valor
End Property


Public Property Get QuantidadeIlha() As Integer
    QuantidadeIlha = mQuantidadeIlha
End Property
Public Property Let QuantidadeIlha(ByVal Valor As Integer)
    mQuantidadeIlha = Valor
End Property


Public Property Get ProgramacaoAntiga() As Boolean
    ProgramacaoAntiga = mProgramacaoAntiga
End Property
Public Property Let ProgramacaoAntiga(ByVal Valor As Boolean)
    mProgramacaoAntiga = Valor
End Property


Public Property Get HoraFechamento1() As Date
    HoraFechamento1 = mHoraFechamento1
End Property
Public Property Let HoraFechamento1(ByVal Valor As Date)
    mHoraFechamento1 = Valor
End Property


Public Property Get HoraFechamento2() As Date
    HoraFechamento2 = mHoraFechamento2
End Property
Public Property Let HoraFechamento2(ByVal Valor As Date)
    mHoraFechamento2 = Valor
End Property


Public Property Get HoraFechamento3() As Date
    HoraFechamento3 = mHoraFechamento3
End Property
Public Property Let HoraFechamento3(ByVal Valor As Date)
    mHoraFechamento3 = Valor
End Property


Public Property Get HoraFechamento4() As Date
    HoraFechamento4 = mHoraFechamento4
End Property
Public Property Let HoraFechamento4(ByVal Valor As Date)
    mHoraFechamento4 = Valor
End Property


Public Property Get HoraFechamento5() As Date
    HoraFechamento5 = mHoraFechamento5
End Property
Public Property Let HoraFechamento5(ByVal Valor As Date)
    mHoraFechamento5 = Valor
End Property


Public Property Get HoraFechamento6() As Date
    HoraFechamento6 = mHoraFechamento6
End Property
Public Property Let HoraFechamento6(ByVal Valor As Date)
    mHoraFechamento6 = Valor
End Property


Public Property Get HoraFechamento7() As Date
    HoraFechamento7 = mHoraFechamento7
End Property
Public Property Let HoraFechamento7(ByVal Valor As Date)
    mHoraFechamento7 = Valor
End Property


Public Property Get HoraFechamento8() As Date
    HoraFechamento8 = mHoraFechamento8
End Property
Public Property Let HoraFechamento8(ByVal Valor As Date)
    mHoraFechamento8 = Valor
End Property


Public Property Get ImprimirReducaoZ() As Boolean
    ImprimirReducaoZ = mImprimirReducaoZ
End Property
Public Property Let ImprimirReducaoZ(ByVal Valor As Boolean)
    mImprimirReducaoZ = Valor
End Property


Public Property Get QuantidadeViasTEF() As Integer
    QuantidadeViasTEF = mQuantidadeViasTEF
End Property
Public Property Let QuantidadeViasTEF(ByVal Valor As Integer)
    mQuantidadeViasTEF = Valor
End Property


Public Property Get ControleSolicitacaoTEF() As Long
    ControleSolicitacaoTEF = mControleSolicitacaoTEF
End Property
Public Property Let ControleSolicitacaoTEF(ByVal Valor As Long)
    mControleSolicitacaoTEF = Valor
End Property


Public Property Get IntegraMovimentoBombaCaixa() As Boolean
    IntegraMovimentoBombaCaixa = mIntegraMovimentoBombaCaixa
End Property
Public Property Let IntegraMovimentoBombaCaixa(ByVal Valor As Boolean)
    mIntegraMovimentoBombaCaixa = Valor
End Property


Public Property Get AlteraAberturaBomba() As Boolean
    AlteraAberturaBomba = mAlteraAberturaBomba
End Property
Public Property Let AlteraAberturaBomba(ByVal Valor As Boolean)
    mAlteraAberturaBomba = Valor
End Property


Public Property Get AlteraPrecoMovimentoBomba() As Boolean
    AlteraPrecoMovimentoBomba = mAlteraPrecoMovimentoBomba
End Property
Public Property Let AlteraPrecoMovimentoBomba(ByVal Valor As Boolean)
    mAlteraPrecoMovimentoBomba = Valor
End Property

Public Property Get ECFBaixaEstoque() As Boolean
    ECFBaixaEstoque = mECFBaixaEstoque
End Property
Public Property Let ECFBaixaEstoque(ByVal Valor As Boolean)
    mECFBaixaEstoque = Valor
End Property


Public Property Get NomeclaturaCaixa() As String
    NomeclaturaCaixa = mNomeclaturaCaixa
End Property
Public Property Let NomeclaturaCaixa(ByVal Valor As String)
    mNomeclaturaCaixa = Valor
End Property


Public Property Get AlteracaoCaixaPeloResponsavel() As Boolean
    AlteracaoCaixaPeloResponsavel = mAlteracaoCaixaPeloResponsavel
End Property
Public Property Let AlteracaoCaixaPeloResponsavel(ByVal Valor As Boolean)
    mAlteracaoCaixaPeloResponsavel = Valor
End Property


Public Property Get AlteraPrecoProdutoPelaVenda() As Boolean
    AlteraPrecoProdutoPelaVenda = mAlteraPrecoProdutoPelaVenda
End Property
Public Property Let AlteraPrecoProdutoPelaVenda(ByVal Valor As Boolean)
    mAlteraPrecoProdutoPelaVenda = Valor
End Property


Public Property Get InverteEncerrantenaPlanilha() As Boolean
    InverteEncerrantenaPlanilha = mInverteEncerrantenaPlanilha
End Property
Public Property Let InverteEncerrantenaPlanilha(ByVal Valor As Boolean)
    mInverteEncerrantenaPlanilha = Valor
End Property


Public Property Get IdentificaFuncionarioaCadaCupom() As Boolean
    IdentificaFuncionarioaCadaCupom = mIdentificaFuncionarioaCadaCupom
End Property
Public Property Let IdentificaFuncionarioaCadaCupom(ByVal Valor As Boolean)
    mIdentificaFuncionarioaCadaCupom = Valor
End Property


Public Property Get RelacaoNotasnoCaixa() As Boolean
    RelacaoNotasnoCaixa = mRelacaoNotasnoCaixa
End Property
Public Property Let RelacaoNotasnoCaixa(ByVal Valor As Boolean)
    mRelacaoNotasnoCaixa = Valor
End Property


Public Property Get BloqueiaVendaPeloEstoque() As Boolean
    BloqueiaVendaPeloEstoque = mBloqueiaVendaPeloEstoque
End Property
Public Property Let BloqueiaVendaPeloEstoque(ByVal Valor As Boolean)
    mBloqueiaVendaPeloEstoque = Valor
End Property


Public Property Get BloqueiaVendaPeloSubEstoque() As Boolean
    BloqueiaVendaPeloSubEstoque = mBloqueiaVendaPeloSubEstoque
End Property
Public Property Let BloqueiaVendaPeloSubEstoque(ByVal Valor As Boolean)
    mBloqueiaVendaPeloSubEstoque = Valor
End Property


Public Property Get NumeroDuplicata() As Long
    NumeroDuplicata = mNumeroDuplicata
End Property
Public Property Let NumeroDuplicata(ByVal Valor As Long)
    mNumeroDuplicata = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        AtribuiValor
        LocalizarCodigo = True
    End If
    rsConfiguracao.Close
    Set rsConfiguracao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ProximaSolicitacaoTEF(ByVal nEmpresa As Integer) As Long
    Dim xSQL As String
    Dim lRecordsAffected As Long
On Error GoTo trata_erro
    ProximaSolicitacaoTEF = 1
    'Adiciona 1 para o "Controle de Solicitacao TEF"
    gSQL = "UPDATE Configuracao SET "
    Call sqlNumero(2, "[Controle de Solicitacao TEF] = [Controle de Solicitacao TEF] + ", 1)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    'Busca "Controle de Solicitacao TEF"
    gSQL = "SELECT [Controle de Solicitacao TEF] FROM Configuracao WHERE Empresa = " & nEmpresa
    If Localizar(1) Then
        ProximaSolicitacaoTEF = rsConfiguracao("Controle de Solicitacao TEF").Value
    End If
    rsConfiguracao.Close
    Set rsConfiguracao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Configuracao ( Empresa, [Quantidade de Periodos], [Quantidade de Bico], [Pular Bomba], [Custo por Duplicata], "
    gSQL = gSQL & "[Valor Superior], [Valor Esquerda], [Extenso 1 Superior], [Extenso 1 Esquerda], [Extenso 2 Superior], "
    gSQL = gSQL & "[Extenso 2 Esquerda], [Favorecido Superior], [Favorecido Esquerda], [Cidade Superior], [Cidade Esquerda], "
    gSQL = gSQL & "[Dia Superior], [Dia Esquerda], [Mes Superior], [Mes Esquerda], [Ano Superior], "
    gSQL = gSQL & "[Ano Esquerda], [Outras Configuracoes], [Mensagem Cobranca], [Quantidade de Ilha], [Programacao Antiga], "
    gSQL = gSQL & "[Hora Fechamento 1], [Hora Fechamento 2], [Hora Fechamento 3], [Hora Fechamento 4], [Hora Fechamento 5], "
    gSQL = gSQL & "[Hora Fechamento 6], [Hora Fechamento 7], [Hora Fechamento 8], [Imprimir Reducao Z], [Quantidade de Vias TEF], "
    gSQL = gSQL & "[Controle de Solicitacao TEF], [Integra Movimento de Bomba no Caixa], [Altera Abertura Bomba], "
    gSQL = gSQL & "[Altera Preco no Movimento de Bomba], [ECF Baixa Estoque], [Nomeclatura do Caixa], [Alteracao no Caixa Pelo Responsavel], "
    gSQL = gSQL & "[Altera Preco Produto Pela Venda], [Inverte Encerrante na Planilha], [Identifica Funcionario a Cada Cupom], "
    gSQL = gSQL & "[Relacao Notas no Caixa], [Bloqueia Venda Pelo Estoque], [Bloqueia Venda Pelo SubEstoque], [Numero da Duplicata] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mQuantidadePeriodos, ", ")
    Call sqlNumero(1, mQuantidadeBico, ", ")
    Call sqlNumero(1, mPularBomba, ", ")
    Call sqlValor(1, mCustoDuplicata, ", ")
    Call sqlValor(1, mValorSuperior, ", ")
    Call sqlValor(1, mValorEsquerda, ", ")
    Call sqlValor(1, mExtenso1Superior, ", ")
    Call sqlValor(1, mExtenso1Esquerda, ", ")
    Call sqlValor(1, mExtenso2Superior, ", ")
    Call sqlValor(1, mExtenso2Esquerda, ", ")
    Call sqlValor(1, mFavorecidoSuperior, ", ")
    Call sqlValor(1, mFavorecidoEsquerda, ", ")
    Call sqlValor(1, mCidadeSuperior, ", ")
    Call sqlValor(1, mCidadeEsquerda, ", ")
    Call sqlValor(1, mDiaSuperior, ", ")
    Call sqlValor(1, mDiaEsquerda, ", ")
    Call sqlValor(1, mMesSuperior, ", ")
    Call sqlValor(1, mMesEsquerda, ", ")
    Call sqlValor(1, mAnoSuperior, ", ")
    Call sqlValor(1, mAnoEsquerda, ", ")
    Call sqlTexto(1, mOutrasConfiguracoes, ", ")
    Call sqlTexto(1, mMensagemCobranca, ", ")
    Call sqlNumero(1, mQuantidadeIlha, ", ")
    Call sqlBoolean(1, mProgramacaoAntiga, ", ")
    Call sqlHora(1, mHoraFechamento1, ", ")
    Call sqlHora(1, mHoraFechamento2, ", ")
    Call sqlHora(1, mHoraFechamento3, ", ")
    Call sqlHora(1, mHoraFechamento4, ", ")
    Call sqlHora(1, mHoraFechamento5, ", ")
    Call sqlHora(1, mHoraFechamento6, ", ")
    Call sqlHora(1, mHoraFechamento7, ", ")
    Call sqlHora(1, mHoraFechamento8, ", ")
    Call sqlBoolean(1, mImprimirReducaoZ, ", ")
    Call sqlNumero(1, mQuantidadeViasTEF, ", ")
    Call sqlNumero(1, mControleSolicitacaoTEF, ", ")
    Call sqlBoolean(1, mIntegraMovimentoBombaCaixa, ", ")
    Call sqlBoolean(1, mAlteraAberturaBomba, ", ")
    Call sqlBoolean(1, mAlteraPrecoMovimentoBomba, ", ")
    Call sqlBoolean(1, mECFBaixaEstoque, ", ")
    Call sqlTexto(1, mNomeclaturaCaixa, ", ")
    Call sqlBoolean(1, mAlteracaoCaixaPeloResponsavel, ", ")
    Call sqlBoolean(1, mAlteraPrecoProdutoPelaVenda, ", ")
    Call sqlBoolean(1, mInverteEncerrantenaPlanilha, ", ")
    Call sqlBoolean(1, mIdentificaFuncionarioaCadaCupom, ", ")
    Call sqlBoolean(1, mRelacaoNotasnoCaixa, ", ")
    Call sqlBoolean(1, mBloqueiaVendaPeloEstoque, ", ")
    Call sqlBoolean(1, mBloqueiaVendaPeloSubEstoque, ", ")
    Call sqlNumero(1, mNumeroDuplicata, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Configuracao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Quantidade de Periodos] = ", mQuantidadePeriodos)
    Call sqlNumero(2, ", [Quantidade de Bico] = ", mQuantidadeBico)
    Call sqlNumero(2, ", [Pular Bomba] = ", mPularBomba)
    Call sqlValor(2, ", [Custo por Duplicata] = ", mCustoDuplicata)
    Call sqlValor(2, ", [Valor Superior] = ", mValorSuperior)
    Call sqlValor(2, ", [Valor Esquerda] = ", mValorEsquerda)
    Call sqlValor(2, ", [Extenso 1 Superior] = ", mExtenso1Superior)
    Call sqlValor(2, ", [Extenso 1 Esquerda] = ", mExtenso1Esquerda)
    Call sqlValor(2, ", [Extenso 2 Superior] = ", mExtenso2Superior)
    Call sqlValor(2, ", [Extenso 2 Esquerda] = ", mExtenso2Esquerda)
    Call sqlValor(2, ", [Favorecido Superior] = ", mFavorecidoSuperior)
    Call sqlValor(2, ", [Favorecido Esquerda] = ", mFavorecidoEsquerda)
    Call sqlValor(2, ", [Cidade Superior] = ", mCidadeSuperior)
    Call sqlValor(2, ", [Cidade Esquerda] = ", mCidadeEsquerda)
    Call sqlValor(2, ", [Dia Superior] = ", mDiaSuperior)
    Call sqlValor(2, ", [Dia Esquerda] = ", mDiaEsquerda)
    Call sqlValor(2, ", [Mes Superior] = ", mMesSuperior)
    Call sqlValor(2, ", [Mes Esquerda] = ", mMesEsquerda)
    Call sqlValor(2, ", [Ano Superior] = ", mAnoSuperior)
    Call sqlValor(2, ", [Ano Esquerda] = ", mAnoEsquerda)
    Call sqlTexto(2, ", [Outras Configuracoes] = ", mOutrasConfiguracoes)
    Call sqlTexto(2, ", [Mensagem Cobranca] = ", mMensagemCobranca)
    Call sqlNumero(2, ", [Quantidade de Ilha] = ", mQuantidadeIlha)
    Call sqlBoolean(2, ", [Programacao Antiga] = ", mProgramacaoAntiga)
    Call sqlHora(2, ", [Hora Fechamento 1] = ", mHoraFechamento1)
    Call sqlHora(2, ", [Hora Fechamento 2] = ", mHoraFechamento2)
    Call sqlHora(2, ", [Hora Fechamento 3] = ", mHoraFechamento3)
    Call sqlHora(2, ", [Hora Fechamento 4] = ", mHoraFechamento4)
    Call sqlHora(2, ", [Hora Fechamento 5] = ", mHoraFechamento5)
    Call sqlHora(2, ", [Hora Fechamento 6] = ", mHoraFechamento6)
    Call sqlHora(2, ", [Hora Fechamento 7] = ", mHoraFechamento7)
    Call sqlHora(2, ", [Hora Fechamento 8] = ", mHoraFechamento8)
    Call sqlBoolean(2, ", [Imprimir Reducao Z] = ", mImprimirReducaoZ)
    Call sqlNumero(2, ", [Quantidade de Vias TEF] = ", mQuantidadeViasTEF)
    Call sqlNumero(2, ", [Controle de Solicitacao TEF] = ", mControleSolicitacaoTEF)
    Call sqlBoolean(2, ", [Integra Movimento de Bomba no Caixa] = ", mIntegraMovimentoBombaCaixa)
    Call sqlBoolean(2, ", [Altera Abertura Bomba] = ", mAlteraAberturaBomba)
    Call sqlBoolean(2, ", [Altera Preco no Movimento de Bomba] = ", mAlteraPrecoMovimentoBomba)
    Call sqlBoolean(2, ", [ECF Baixa Estoque] = ", mECFBaixaEstoque)
    Call sqlTexto(2, ", [Nomeclatura do Caixa] = ", mNomeclaturaCaixa)
    Call sqlBoolean(2, ", [Alteracao no Caixa Pelo Responsavel] = ", mAlteracaoCaixaPeloResponsavel)
    Call sqlBoolean(2, ", [Altera Preco Produto Pela Venda] = ", mAlteraPrecoProdutoPelaVenda)
    Call sqlBoolean(2, ", [Inverte Encerrante na Planilha] = ", mInverteEncerrantenaPlanilha)
    Call sqlBoolean(2, ", [Identifica Funcionario a Cada Cupom] = ", mIdentificaFuncionarioaCadaCupom)
    Call sqlBoolean(2, ", [Relacao Notas no Caixa] = ", mRelacaoNotasnoCaixa)
    Call sqlBoolean(2, ", [Bloqueia Venda Pelo Estoque] = ", mBloqueiaVendaPeloEstoque)
    Call sqlBoolean(2, ", [Bloqueia Venda Pelo SubEstoque] = ", mBloqueiaVendaPeloSubEstoque)
    Call sqlNumero(2, ", [Numero da Duplicata] = ", mNumeroDuplicata)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Configuracao"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = "SELECT Empresa, [Quantidade de Periodos], [Quantidade de Bico], [Pular Bomba], [Custo por Duplicata], [Valor Superior], [Valor Esquerda], [Extenso 1 Superior], [Extenso 1 Esquerda], [Extenso 2 Superior], [Extenso 2 Esquerda], [Favorecido Superior], [Favorecido Esquerda], [Cidade Superior], [Cidade Esquerda], [Dia Superior], [Dia Esquerda], [Mes Superior], [Mes Esquerda], [Ano Superior], [Ano Esquerda], [Outras Configuracoes], [Mensagem Cobranca], [Quantidade de Ilha], [Programacao Antiga], [Hora Fechamento 1], [Hora Fechamento 2], [Hora Fechamento 3], [Hora Fechamento 4], [Hora Fechamento 5], [Hora Fechamento 6], [Hora Fechamento 7], [Hora Fechamento 8], [Imprimir Reducao Z], [Quantidade de Vias TEF], [Controle de Solicitacao TEF], [Integra Movimento de Bomba no Caixa], [Altera Abertura Bomba], [Altera Preco no Movimento de Bomba], [ECF Baixa Estoque], [Nomeclatura do Caixa], "
    PreparaSQL = PreparaSQL & "[Alteracao no Caixa Pelo Responsavel], [Altera Preco Produto Pela Venda], [Inverte Encerrante na Planilha], [Identifica Funcionario a Cada Cupom], [Relacao Notas no Caixa], [Bloqueia Venda Pelo Estoque], [Bloqueia Venda Pelo SubEstoque], [Numero da Duplicata] FROM Configuracao" & " " & xCondicao & " " & xTipoOrdem
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConfiguracao = New adodb.Recordset
    rsConfiguracao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConfiguracao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConfiguracao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsConfiguracao("Empresa").Value
    QuantidadePeriodos = rsConfiguracao("Quantidade de Periodos").Value
    QuantidadeBico = rsConfiguracao("Quantidade de Bico").Value
    PularBomba = rsConfiguracao("Pular Bomba").Value
    CustoDuplicata = rsConfiguracao("Custo por Duplicata").Value
    ValorSuperior = rsConfiguracao("Valor Superior").Value
    ValorEsquerda = rsConfiguracao("Valor Esquerda").Value
    Extenso1Superior = rsConfiguracao("Extenso 1 Superior").Value
    Extenso1Esquerda = rsConfiguracao("Extenso 1 Esquerda").Value
    Extenso2Superior = rsConfiguracao("Extenso 2 Superior").Value
    Extenso2Esquerda = rsConfiguracao("Extenso 2 Esquerda").Value
    FavorecidoSuperior = rsConfiguracao("Favorecido Superior").Value
    FavorecidoEsquerda = rsConfiguracao("Favorecido Esquerda").Value
    CidadeSuperior = rsConfiguracao("Cidade Superior").Value
    CidadeEsquerda = rsConfiguracao("Cidade Esquerda").Value
    DiaSuperior = rsConfiguracao("Dia Superior").Value
    DiaEsquerda = rsConfiguracao("Dia Esquerda").Value
    MesSuperior = rsConfiguracao("Mes Superior").Value
    MesEsquerda = rsConfiguracao("Mes Esquerda").Value
    AnoSuperior = rsConfiguracao("Ano Superior").Value
    AnoEsquerda = rsConfiguracao("Ano Esquerda").Value
    OutrasConfiguracoes = rsConfiguracao("Outras Configuracoes").Value
    MensagemCobranca = rsConfiguracao("Mensagem Cobranca").Value
    QuantidadeIlha = rsConfiguracao("Quantidade de Ilha").Value
    ProgramacaoAntiga = rsConfiguracao("Programacao Antiga").Value
    If IsNull(rsConfiguracao("Hora Fechamento 1").Value) Then
        HoraFechamento1 = "00:00:00"
    Else
        HoraFechamento1 = rsConfiguracao("Hora Fechamento 1").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 2").Value) Then
        HoraFechamento2 = "00:00:00"
    Else
        HoraFechamento2 = rsConfiguracao("Hora Fechamento 2").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 3").Value) Then
        HoraFechamento3 = "00:00:00"
    Else
        HoraFechamento3 = rsConfiguracao("Hora Fechamento 3").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 4").Value) Then
        HoraFechamento4 = "00:00:00"
    Else
        HoraFechamento4 = rsConfiguracao("Hora Fechamento 4").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 5").Value) Then
        HoraFechamento5 = "00:00:00"
    Else
        HoraFechamento5 = rsConfiguracao("Hora Fechamento 5").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 6").Value) Then
        HoraFechamento6 = "00:00:00"
    Else
        HoraFechamento6 = rsConfiguracao("Hora Fechamento 6").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 7").Value) Then
        HoraFechamento7 = "00:00:00"
    Else
        HoraFechamento7 = rsConfiguracao("Hora Fechamento 7").Value
    End If
    If IsNull(rsConfiguracao("Hora Fechamento 8").Value) Then
        HoraFechamento8 = "00:00:00"
    Else
        HoraFechamento8 = rsConfiguracao("Hora Fechamento 8").Value
    End If
    ImprimirReducaoZ = rsConfiguracao("Imprimir Reducao Z").Value
    QuantidadeViasTEF = rsConfiguracao("Quantidade de Vias TEF").Value
    ControleSolicitacaoTEF = rsConfiguracao("Controle de Solicitacao TEF").Value
    IntegraMovimentoBombaCaixa = rsConfiguracao("Integra Movimento de Bomba no Caixa").Value
    AlteraAberturaBomba = rsConfiguracao("Altera Abertura Bomba").Value
    AlteraPrecoMovimentoBomba = rsConfiguracao("Altera Preco no Movimento de Bomba").Value
    ECFBaixaEstoque = rsConfiguracao("ECF Baixa Estoque").Value
    NomeclaturaCaixa = rsConfiguracao("Nomeclatura do Caixa").Value
    AlteracaoCaixaPeloResponsavel = rsConfiguracao("Alteracao no Caixa Pelo Responsavel").Value
    AlteraPrecoProdutoPelaVenda = rsConfiguracao("Altera Preco Produto Pela Venda").Value
    InverteEncerrantenaPlanilha = rsConfiguracao("Inverte Encerrante na Planilha").Value
    IdentificaFuncionarioaCadaCupom = rsConfiguracao("Identifica Funcionario a Cada Cupom").Value
    RelacaoNotasnoCaixa = rsConfiguracao("Relacao Notas no Caixa").Value
    BloqueiaVendaPeloEstoque = rsConfiguracao("Bloqueia Venda Pelo Estoque").Value
    BloqueiaVendaPeloSubEstoque = rsConfiguracao("Bloqueia Venda Pelo SubEstoque").Value
    NumeroDuplicata = rsConfiguracao("Numero da Duplicata").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

