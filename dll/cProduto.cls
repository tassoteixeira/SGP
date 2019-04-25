VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoGrupo As Integer
Private mCodigo As Long
Private mNome As String
Private mUnidade As String
Private mPrecoVenda As Currency
Private mPrecoCustoAnterior As Currency
Private mPrecoCusto As Currency
Private mMargemLucroDesejada As Currency
Private mPrecoCustoMedio As Currency
Private mEstoqueMinimo As Currency
Private mEstoqueMaximo As Currency
Private mEstoqueFisico As Currency
Private mDataUltimaCompra As Date
Private mDataAcerto As Date
Private mComissao As Boolean
Private mInativo As Boolean
Private mCodigoAliquota As Integer
Private mCodigoSubGrupo As Integer
Private mCodigoBarra As String
Private mExclusivoPosto As Boolean
Private mExclusivoLoja As Boolean
Private mTipoCombustivel As String
Private mDescontoEspecial As Currency
Private mPercentualComissao As Currency
Private mImprimeCupomFiscal As Boolean
Private mComissaoEmPercentual As Boolean
Private mCSTICMS As Integer
Private mCSTIPI As Integer
Private mCSTCOFINS As Integer
Private mCSTPIS As Integer
Private mCodigoNCM As String
Private mTipoProduto As Integer
Private mCodigoServico As Integer
Private mDataAlteracao As Date
Private mCodigoANP As String
Private mCEST As String

Private rsProduto As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property


Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Long)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Unidade() As String
    Unidade = mUnidade
End Property
Public Property Let Unidade(ByVal Valor As String)
    mUnidade = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get PrecoCustoAnterior() As Currency
    PrecoCustoAnterior = mPrecoCustoAnterior
End Property
Public Property Let PrecoCustoAnterior(ByVal Valor As Currency)
    mPrecoCustoAnterior = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get MargemLucroDesejada() As Currency
    MargemLucroDesejada = mMargemLucroDesejada
End Property
Public Property Let MargemLucroDesejada(ByVal Valor As Currency)
    mMargemLucroDesejada = Valor
End Property


Public Property Get PrecoCustoMedio() As Currency
    PrecoCustoMedio = mPrecoCustoMedio
End Property
Public Property Let PrecoCustoMedio(ByVal Valor As Currency)
    mPrecoCustoMedio = Valor
End Property


Public Property Get EstoqueMinimo() As Currency
    EstoqueMinimo = mEstoqueMinimo
End Property
Public Property Let EstoqueMinimo(ByVal Valor As Currency)
    mEstoqueMinimo = Valor
End Property


Public Property Get EstoqueMaximo() As Currency
    EstoqueMaximo = mEstoqueMaximo
End Property
Public Property Let EstoqueMaximo(ByVal Valor As Currency)
    mEstoqueMaximo = Valor
End Property


Public Property Get EstoqueFisico() As Currency
    EstoqueFisico = mEstoqueFisico
End Property
Public Property Let EstoqueFisico(ByVal Valor As Currency)
    mEstoqueFisico = Valor
End Property


Public Property Get DataUltimaCompra() As Date
    DataUltimaCompra = mDataUltimaCompra
End Property
Public Property Let DataUltimaCompra(ByVal Valor As Date)
    mDataUltimaCompra = Valor
End Property


Public Property Get DataAcerto() As Date
    DataAcerto = mDataAcerto
End Property
Public Property Let DataAcerto(ByVal Valor As Date)
    mDataAcerto = Valor
End Property


Public Property Get Comissao() As Boolean
    Comissao = mComissao
End Property
Public Property Let Comissao(ByVal Valor As Boolean)
    mComissao = Valor
End Property


Public Property Get Inativo() As Boolean
    Inativo = mInativo
End Property
Public Property Let Inativo(ByVal Valor As Boolean)
    mInativo = Valor
End Property


Public Property Get CodigoAliquota() As Integer
    CodigoAliquota = mCodigoAliquota
End Property
Public Property Let CodigoAliquota(ByVal Valor As Integer)
    mCodigoAliquota = Valor
End Property


Public Property Get CodigoSubGrupo() As Integer
    CodigoSubGrupo = mCodigoSubGrupo
End Property
Public Property Let CodigoSubGrupo(ByVal Valor As Integer)
    mCodigoSubGrupo = Valor
End Property


Public Property Get CodigoBarra() As String
    CodigoBarra = mCodigoBarra
End Property
Public Property Let CodigoBarra(ByVal Valor As String)
    mCodigoBarra = Valor
End Property


Public Property Get ExclusivoPosto() As Boolean
    ExclusivoPosto = mExclusivoPosto
End Property
Public Property Let ExclusivoPosto(ByVal Valor As Boolean)
    mExclusivoPosto = Valor
End Property


Public Property Get ExclusivoLoja() As Boolean
    ExclusivoLoja = mExclusivoLoja
End Property
Public Property Let ExclusivoLoja(ByVal Valor As Boolean)
    mExclusivoLoja = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get DescontoEspecial() As Currency
    DescontoEspecial = mDescontoEspecial
End Property
Public Property Let DescontoEspecial(ByVal Valor As Currency)
    mDescontoEspecial = Valor
End Property


Public Property Get PercentualComissao() As Currency
    PercentualComissao = mPercentualComissao
End Property
Public Property Let PercentualComissao(ByVal Valor As Currency)
    mPercentualComissao = Valor
End Property


Public Property Get ImprimeCupomFiscal() As Boolean
    ImprimeCupomFiscal = mImprimeCupomFiscal
End Property
Public Property Let ImprimeCupomFiscal(ByVal Valor As Boolean)
    mImprimeCupomFiscal = Valor
End Property


Public Property Get ComissaoEmPercentual() As Boolean
    ComissaoEmPercentual = mComissaoEmPercentual
End Property
Public Property Let ComissaoEmPercentual(ByVal Valor As Boolean)
    mComissaoEmPercentual = Valor
End Property


Public Property Get CSTICMS() As Integer
    CSTICMS = mCSTICMS
End Property
Public Property Let CSTICMS(ByVal Valor As Integer)
    mCSTICMS = Valor
End Property


Public Property Get CSTIPI() As Integer
    CSTIPI = mCSTIPI
End Property
Public Property Let CSTIPI(ByVal Valor As Integer)
    mCSTIPI = Valor
End Property


Public Property Get CSTCOFINS() As Integer
    CSTCOFINS = mCSTCOFINS
End Property
Public Property Let CSTCOFINS(ByVal Valor As Integer)
    mCSTCOFINS = Valor
End Property


Public Property Get CSTPIS() As Integer
    CSTPIS = mCSTPIS
End Property
Public Property Let CSTPIS(ByVal Valor As Integer)
    mCSTPIS = Valor
End Property


Public Property Get CodigoNCM() As String
    CodigoNCM = mCodigoNCM
End Property
Public Property Let CodigoNCM(ByVal Valor As String)
    mCodigoNCM = Valor
End Property


Public Property Get TipoProduto() As Integer
    TipoProduto = mTipoProduto
End Property
Public Property Let TipoProduto(ByVal Valor As Integer)
    mTipoProduto = Valor
End Property


Public Property Get CodigoServico() As Integer
    CodigoServico = mCodigoServico
End Property
Public Property Let CodigoServico(ByVal Valor As Integer)
    mCodigoServico = Valor
End Property


Public Property Get DataAlteracao() As Date
    DataAlteracao = mDataAlteracao
End Property
Public Property Let DataAlteracao(ByVal Valor As Date)
    mDataAlteracao = Valor
End Property

Public Property Get CodigoANP() As String
    CodigoANP = mCodigoANP
End Property
Public Property Let CodigoANP(ByVal Valor As String)
    mCodigoANP = Valor
End Property

Public Property Get CEST() As String
    CEST = mCEST
End Property
Public Property Let CEST(ByVal Valor As String)
    mCEST = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigoBarra(ByVal nCodigoBarra As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigoBarra = False
    xCondicao = " WHERE [Codigo de Barra] = " & preparaTexto(nCodigoBarra)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigoBarra = Localizar
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


Public Function IncluiEstoque() As Boolean
Dim lRecordsAffected As Long
Dim xEmpresa As Integer
Dim xUltimaEmpresa As Integer

On Error GoTo Err_IncluiEstoque

    IncluiEstoque = False
    xUltimaEmpresa = UltimaEmpresa

    'Inclui Estoque
    For xEmpresa = 1 To xUltimaEmpresa
        gSQL = "INSERT INTO Estoque ( Empresa, [Grupo do Produto], [Codigo do Produto2], Quantidade, [Preco de Venda], [Quantidade Temporaria] ) VALUES ( "
        Call sqlNumero(1, xEmpresa, ", ")
        Call sqlNumero(1, mCodigoGrupo, ", ")
        Call sqlNumero(1, mCodigo, ", ")
        Call sqlValor(1, 0, ", ")
        Call sqlValor(1, mPrecoVenda, ", ")
        Call sqlValor(1, 0, " )")
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected = 0 Then
            IncluiEstoque = False
            Exit Function
        End If
        IncluiEstoque = True
    Next
    Exit Function

Err_IncluiEstoque:
    MsgBox "Erro na Rotina", vbInformation, "IncluiEstoque"
    Exit Function
End Function


Public Function IncluiEstoquePadrao() As Boolean
Dim lRecordsAffected As Long
Dim xEmpresa As Integer
Dim xSubEstoque As Integer
Dim xUltimaEmpresa As Integer
Dim xUltimoSubEstoque As Integer

On Error GoTo Err_IncluiEstoquePadrao

    IncluiEstoquePadrao = False
    xUltimaEmpresa = UltimaEmpresa
    xUltimoSubEstoque = UltimoSubEstoque
    'Inclui EstoquePadrao
    For xEmpresa = 1 To xUltimaEmpresa
        For xSubEstoque = 1 To xUltimoSubEstoque
            gSQL = "INSERT INTO EstoquePadrao ( Empresa, [Codigo do Produto], [Codigo do Tipo do SubEstoque], Quantidade, Observacao ) VALUES ( "
            Call sqlNumero(1, xEmpresa, ", ")
            Call sqlNumero(1, mCodigo, ", ")
            Call sqlNumero(1, xSubEstoque, ", ")
            Call sqlValor(1, 0, ", ")
            Call sqlTexto(1, "", " )")
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected = 0 Then
                IncluiEstoquePadrao = False
                Exit Function
            End If
            IncluiEstoquePadrao = True
        Next
    Next
    Exit Function

Err_IncluiEstoquePadrao:
    MsgBox "Erro na Rotina", vbInformation, "IncluiEstoquePadrao"
    Exit Function
End Function


Public Function IncluiSubEstoque() As Boolean
Dim lRecordsAffected As Long
Dim xEmpresa As Integer
Dim xSubEstoque As Integer
Dim xUltimaEmpresa As Integer
Dim xUltimoSubEstoque As Integer

On Error GoTo Err_IncluiSubEstoque

    IncluiSubEstoque = False
    xUltimaEmpresa = UltimaEmpresa
    xUltimoSubEstoque = UltimoSubEstoque
    'Inclui SubEstoque
    For xEmpresa = 1 To xUltimaEmpresa
        For xSubEstoque = 1 To xUltimoSubEstoque
            gSQL = "INSERT INTO SubEstoque ( Empresa, [Codigo do Produto], [Codigo do Tipo do SubEstoque], Quantidade ) VALUES ( "
            Call sqlNumero(1, xEmpresa, ", ")
            Call sqlNumero(1, mCodigo, ", ")
            Call sqlNumero(1, xSubEstoque, ", ")
            Call sqlValor(1, 0, " )")
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected = 0 Then
                IncluiSubEstoque = False
                Exit Function
            End If
            IncluiSubEstoque = True
        Next
    Next
    Exit Function

Err_IncluiSubEstoque:
    MsgBox "Erro na Rotina", vbInformation, "IncluiSubEstoque"
    Exit Function
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Produto ( [Codigo do Grupo], Codigo, Nome, Unidade, [Preco de Venda], "
    gSQL = gSQL & "[Preco de Custo Anterior], [Preco de Custo], [Margem de Lucro Desejada], "
    gSQL = gSQL & "[Preco de Custo Medio], [Estoque Minimo], [Estoque Maximo], [Estoque Fisico], "
    gSQL = gSQL & "[Data da Ultima Compra], [Data do Acerto], Comissao, Inativo, [Codigo da Aliquota], "
    gSQL = gSQL & "[Codigo do SubGrupo], [Codigo de Barra], [Exclusivo Posto], [Exclusivo Loja], "
    gSQL = gSQL & "[Tipo de Combustivel], [Desconto Especial], [Percentual da Comissao], [Imprime Cupom Fiscal], "
    gSQL = gSQL & "[Comissao Em Percentual], [CST ICMS], [CST IPI], [CST COFINS], [CST PIS], "
    gSQL = gSQL & "[Codigo NCM], [Tipo de Produto], [Codigo do Servico], [Data da Alteracao], [Codigo ANP], CEST ) VALUES ( "
    Call sqlNumero(1, mCodigoGrupo, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mUnidade, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mPrecoCustoAnterior, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mMargemLucroDesejada, ", ")
    Call sqlValor(1, mPrecoCustoMedio, ", ")
    Call sqlValor(1, mEstoqueMinimo, ", ")
    Call sqlValor(1, mEstoqueMaximo, ", ")
    Call sqlValor(1, mEstoqueFisico, ", ")
    Call sqlData(1, mDataUltimaCompra, ", ")
    Call sqlData(1, mDataAcerto, ", ")
    Call sqlBoolean(1, mComissao, ", ")
    Call sqlBoolean(1, mInativo, ", ")
    Call sqlNumero(1, mCodigoAliquota, ", ")
    Call sqlNumero(1, mCodigoSubGrupo, ", ")
    Call sqlTexto(1, mCodigoBarra, ", ")
    Call sqlBoolean(1, mExclusivoPosto, ", ")
    Call sqlBoolean(1, mExclusivoLoja, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlValor(1, mDescontoEspecial, ", ")
    Call sqlValor(1, mPercentualComissao, ", ")
    Call sqlBoolean(1, mImprimeCupomFiscal, ", ")
    Call sqlBoolean(1, mComissaoEmPercentual, ", ")
    Call sqlNumero(1, mCSTICMS, ", ")
    Call sqlNumero(1, mCSTIPI, ", ")
    Call sqlNumero(1, mCSTCOFINS, ", ")
    Call sqlNumero(1, mCSTPIS, ", ")
    Call sqlTexto(1, mCodigoNCM, ", ")
    Call sqlNumero(1, mTipoProduto, ", ")
    Call sqlNumero(1, mCodigoServico, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlTexto(1, mCodigoANP, ", ")
    Call sqlTexto(1, mCEST, " )")
    gConn.BeginTrans
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If IncluiEstoque Then
            If IncluiSubEstoque Then
                If IncluiEstoquePadrao Then
                    Incluir = True
                Else
                    MsgBox "Erro na inclusão do EstoquePadrao."
                    gConn.RollbackTrans
                    Exit Function
                End If
            Else
                MsgBox "Erro na inclusão do Sub-Estoque."
                gConn.RollbackTrans
                Exit Function
            End If
        Else
            MsgBox "Erro na inclusão do estoque."
            gConn.RollbackTrans
            Exit Function
        End If
    Else
        Call CriaLogCadastroDll2("Classe:cProduto - Método:Incluir", "Não foi possível incluir registro.", gSQL)
    End If
    gConn.CommitTrans
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("Classe:cProduto - Método:Incluir", Err.Description, gSQL)
    gConn.RollbackTrans
End Function


Public Function Alterar(ByVal pCodigo As Long, ByVal pEmpresa As Integer) As Boolean
Dim lRecordsAffected As Long
Dim xFase As Integer

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Produto SET "
    Call sqlNumero(2, "[Codigo do Grupo] = ", mCodigoGrupo)
    xFase = 1
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    xFase = 2
    Call sqlTexto(2, ", Nome = ", mNome)
    xFase = 3
    Call sqlTexto(2, ", Unidade = ", mUnidade)
    xFase = 4
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    xFase = 5
    Call sqlValor(2, ", [Preco de Custo Anterior] = ", mPrecoCustoAnterior)
    xFase = 6
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    xFase = 7
    Call sqlValor(2, ", [Margem de Lucro Desejada] = ", mMargemLucroDesejada)
    xFase = 8
    Call sqlValor(2, ", [Preco de Custo Medio] = ", mPrecoCustoMedio)
    xFase = 9
    Call sqlValor(2, ", [Estoque Minimo] = ", mEstoqueMinimo)
    xFase = 10
    Call sqlValor(2, ", [Estoque Maximo] = ", mEstoqueMaximo)
    xFase = 11
    Call sqlValor(2, ", [Estoque Fisico] = ", mEstoqueFisico)
    xFase = 12
    Call sqlData(2, ", [Data da Ultima Compra] = ", mDataUltimaCompra)
    xFase = 13
    Call sqlData(2, ", [Data do Acerto] = ", mDataAcerto)
    xFase = 14
    Call sqlBoolean(2, ", Comissao = ", mComissao)
    xFase = 15
    Call sqlBoolean(2, ", Inativo = ", mInativo)
    xFase = 16
    Call sqlNumero(2, ", [Codigo da Aliquota] = ", mCodigoAliquota)
    xFase = 17
    Call sqlNumero(2, ", [Codigo do SubGrupo] = ", mCodigoSubGrupo)
    xFase = 18
    Call sqlTexto(2, ", [Codigo de Barra] = ", mCodigoBarra)
    xFase = 19
    Call sqlBoolean(2, ", [Exclusivo Posto] = ", mExclusivoPosto)
    xFase = 20
    Call sqlBoolean(2, ", [Exclusivo Loja] = ", mExclusivoLoja)
    xFase = 21
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    xFase = 22
    Call sqlValor(2, ", [Desconto Especial] = ", mDescontoEspecial)
    xFase = 23
    Call sqlValor(2, ", [Percentual da Comissao] = ", mPercentualComissao)
    xFase = 24
    Call sqlBoolean(2, ", [Imprime Cupom Fiscal] = ", mImprimeCupomFiscal)
    xFase = 25
    Call sqlBoolean(2, ", [Comissao Em Percentual] = ", mComissaoEmPercentual)
    xFase = 26
    Call sqlNumero(2, ", [CST ICMS] = ", mCSTICMS)
    xFase = 27
    Call sqlNumero(2, ", [CST IPI] = ", mCSTIPI)
    xFase = 28
    Call sqlNumero(2, ", [CST COFINS] = ", mCSTCOFINS)
    xFase = 29
    Call sqlNumero(2, ", [CST PIS] = ", mCSTPIS)
    xFase = 30
    Call sqlTexto(2, ", [Codigo NCM] = ", mCodigoNCM)
    xFase = 31
    Call sqlNumero(2, ", [Tipo de Produto] = ", mTipoProduto)
    xFase = 32
    Call sqlNumero(2, ", [Codigo do Servico] = ", mCodigoServico)
    xFase = 33
    Call sqlData(2, ", [Data da Alteracao] = ", mDataAlteracao)
    xFase = 34
    Call sqlTexto(2, ", [Codigo ANP] = ", mCodigoANP)
    xFase = 35
    Call sqlTexto(2, ", CEST = ", mCEST)
    xFase = 36
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    xFase = 37
    'gConn.BeginTrans
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Call CriaLogCadastroDll2("cProduto:Alterar - ** Investigação Especial - Registro alterado p/:", "Investigação Especial", gSQL)
        If AlteraPrecoVendaTabEstoque(pEmpresa) Then
            Alterar = True
        Else
            'gConn.RollbackTrans
            MsgBox "Erro na alteração do estoque."
            Exit Function
        End If
    End If
    'gConn.CommitTrans
    Exit Function

Err_Alterar:
    Call CriaLogCadastroDll2("Classe:cProduto - Método:Alterar Fase: " & xFase, Err.Description, gSQL)
    gConn.RollbackTrans
End Function


Public Function AlterarCusto(ByVal pCodigo As Long, ByVal pEmpresa As Integer, ByVal pPrecoCusto As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarCusto = False
    gSQL = "UPDATE Produto SET "
    Call sqlValor(2, "[Preco de Custo] = ", pPrecoCusto)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    'gConn.BeginTrans
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarCusto = True
        'If AlteraPrecoCustoTabEstoque(pCodigo, pEmpresa, pPrecoCusto) Then
        '    AlterarCusto = True
        'Else
        '    'gConn.RollbackTrans
        '    MsgBox "Erro na alteração do estoque."
        '    Exit Function
        'End If
    End If
    'gConn.CommitTrans
    Exit Function

Err_Alterar:
    gConn.RollbackTrans
End Function


Public Function ExcluiEstoque() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_ExcluiEstoque

    ExcluiEstoque = False

    'Exclui Estoque
    gSQL = ""
    gSQL = gSQL & "DELETE * "
    gSQL = gSQL & "  FROM Estoque "
    gSQL = gSQL & " WHERE [Codigo do Produto2] = " & mCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluiEstoque = True
    End If
    Exit Function

Err_ExcluiEstoque:
    MsgBox "Erro na Rotina", vbInformation, "ExcluiEstoque"
    Exit Function
End Function


Public Function ExcluiSubEstoque() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_ExcluiSubEstoque

    ExcluiSubEstoque = False

    'Exclui SubEstoque
    gSQL = ""
    gSQL = gSQL & "DELETE * "
    gSQL = gSQL & "  FROM SubEstoque "
    gSQL = gSQL & " WHERE [Codigo do Produto] = " & mCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluiSubEstoque = True
    End If
    Exit Function

Err_ExcluiSubEstoque:
    MsgBox "Erro na Rotina", vbInformation, "ExcluiSubEstoque"
    Exit Function
End Function


Public Function Excluir(ByVal nCodigo As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Produto"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.BeginTrans
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If ExcluiEstoque Then
            If ExcluiSubEstoque Then
                Excluir = True
            Else
                gConn.RollbackTrans
                MsgBox "Erro na exclusão do Sub-Estoque."
                Exit Function
            End If
        Else
            gConn.RollbackTrans
            MsgBox "Erro na exclusão do estoque."
            Exit Function
        End If
    End If
    gConn.CommitTrans
    Exit Function

Err_Excluir:
    gConn.RollbackTrans
End Function


'Funções / Procedures internas
Private Function AlteraPrecoCustoTabEstoque(ByVal pCodigo As Long, ByVal pEmpresa As Integer, ByVal pPrecoCusto As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlteraPrecoVendaTabEstoque

    AlteraPrecoCustoTabEstoque = False

    'Altera Estoque
    gSQL = ""
    gSQL = gSQL & "UPDATE Estoque SET "
    Call sqlValor(2, "[Preco de Custo] = ", pPrecoCusto)
    gSQL = gSQL & " WHERE [Codigo do Produto2] = " & pCodigo
    If pEmpresa > 0 Then
        gSQL = gSQL & " AND Empresa = " & pEmpresa
    End If
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected = 0 Then
        Exit Function
    End If
    AlteraPrecoCustoTabEstoque = True
    Exit Function

Err_AlteraPrecoVendaTabEstoque:
    MsgBox "Erro na Rotina", vbInformation, "AlteraPrecoCustoTabEstoque"
    Exit Function
End Function


Private Function AlteraPrecoVendaTabEstoque(ByVal pEmpresa As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlteraPrecoVendaTabEstoque

    AlteraPrecoVendaTabEstoque = False

    'Altera Estoque
    gSQL = ""
    gSQL = gSQL & "UPDATE Estoque SET "
    Call sqlValor(2, "[Preco de Venda] = ", mPrecoVenda)
    gSQL = gSQL & " WHERE [Codigo do Produto2] = " & mCodigo
    If pEmpresa > 0 Then
        gSQL = gSQL & " AND Empresa = " & pEmpresa
    End If
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected = 0 Then
        Exit Function
    End If
    AlteraPrecoVendaTabEstoque = True
    Exit Function

Err_AlteraPrecoVendaTabEstoque:
    MsgBox "Erro na Rotina", vbInformation, "AlteraPrecoVendaTabEstoque"
    Exit Function
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsProduto = New adodb.Recordset
    rsProduto.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsProduto.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Grupo], Codigo, Nome, Unidade, [Preco de Venda], [Preco de Custo Anterior], [Preco de Custo], [Margem de Lucro Desejada], [Preco de Custo Medio], [Estoque Minimo], [Estoque Maximo], [Estoque Fisico], [Data da Ultima Compra], [Data do Acerto], Comissao, Inativo, [Codigo da Aliquota], [Codigo do SubGrupo], [Codigo de Barra], [Exclusivo Posto], [Exclusivo Loja], [Tipo de Combustivel], [Desconto Especial], [Percentual da Comissao], [Imprime Cupom Fiscal], [Comissao Em Percentual], [CST ICMS], [CST IPI], [CST COFINS], [CST PIS], [Codigo NCM], [Tipo de Produto], [Codigo do Servico], [Data da Alteracao], [Codigo ANP], CEST FROM Produto" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function UltimaEmpresa() As Integer

On Error GoTo trata_erro

    UltimaEmpresa = 1
    gSQL = "SELECT TOP 1 Codigo FROM Empresas ORDER BY Codigo DESC"
    Set rsProduto = New adodb.Recordset
    rsProduto.CursorLocation = adUseClient
    rsProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsProduto.EOF Then
        UltimaEmpresa = rsProduto("Codigo").Value
    End If
    rsProduto.Close
    Set rsProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function UltimoSubEstoque() As Integer

On Error GoTo trata_erro

    UltimoSubEstoque = 1
    gSQL = "SELECT TOP 1 Codigo FROM TipoSubEstoque ORDER BY Codigo DESC"
    Set rsProduto = New adodb.Recordset
    rsProduto.CursorLocation = adUseClient
    rsProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsProduto.EOF Then
        UltimoSubEstoque = rsProduto("Codigo").Value
    End If
    rsProduto.Close
    Set rsProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Sub AtribuiValor()
    CodigoGrupo = rsProduto("Codigo do Grupo").Value
    Codigo = rsProduto("Codigo").Value
    Nome = rsProduto("Nome").Value
    Unidade = rsProduto("Unidade").Value
    PrecoVenda = rsProduto("Preco de Venda").Value
    PrecoCustoAnterior = rsProduto("Preco de Custo Anterior").Value
    PrecoCusto = rsProduto("Preco de Custo").Value
    MargemLucroDesejada = rsProduto("Margem de Lucro Desejada").Value
    PrecoCustoMedio = rsProduto("Preco de Custo Medio").Value
    EstoqueMinimo = rsProduto("Estoque Minimo").Value
    EstoqueMaximo = rsProduto("Estoque Maximo").Value
    EstoqueFisico = rsProduto("Estoque Fisico").Value
    If IsNull(rsProduto("Data da Ultima Compra").Value) Then
        DataUltimaCompra = "00:00:00"
    Else
        DataUltimaCompra = rsProduto("Data da Ultima Compra").Value
    End If
    If IsNull(rsProduto("Data do Acerto").Value) Then
        DataAcerto = "00:00:00"
    Else
        DataAcerto = rsProduto("Data do Acerto").Value
    End If
    Comissao = rsProduto("Comissao").Value
    Inativo = rsProduto("Inativo").Value
    CodigoAliquota = rsProduto("Codigo da Aliquota").Value
    CodigoSubGrupo = rsProduto("Codigo do SubGrupo").Value
    CodigoBarra = rsProduto("Codigo de Barra").Value
    ExclusivoPosto = rsProduto("Exclusivo Posto").Value
    ExclusivoLoja = rsProduto("Exclusivo Loja").Value
    TipoCombustivel = rsProduto("Tipo de Combustivel").Value
    DescontoEspecial = rsProduto("Desconto Especial").Value
    PercentualComissao = rsProduto("Percentual da Comissao").Value
    ImprimeCupomFiscal = rsProduto("Imprime Cupom Fiscal").Value
    ComissaoEmPercentual = rsProduto("Comissao Em Percentual").Value
    CSTICMS = rsProduto("CST ICMS").Value
    CSTIPI = rsProduto("CST IPI").Value
    CSTCOFINS = rsProduto("CST COFINS").Value
    CSTPIS = rsProduto("CST PIS").Value
    CodigoNCM = rsProduto("Codigo NCM").Value
    TipoProduto = rsProduto("Tipo de Produto").Value
    CodigoServico = rsProduto("Codigo do Servico").Value
    CodigoANP = rsProduto("Codigo ANP").Value
    CEST = rsProduto("CEST").Value
    If IsNull(rsProduto("Data da Alteracao").Value) Then
        DataAlteracao = "00:00:00"
    Else
        DataAlteracao = rsProduto("Data da Alteracao").Value
    End If
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsProduto.Close
    Set rsProduto = Nothing
End Function

