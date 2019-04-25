VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntradaCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNomeTabela As String
Private mEmpresa As Integer
Private mData As Date
Private mTipoCombustivel As String
Private mNumeroNota As String
Private mValorLitro As Currency
Private mQuantidade As Currency
Private mValorEntrada As Currency
Private mCodigoFornecedor As Integer
Private mTipoTransporte As Integer
Private mItem As Integer
Private mModelo As String
Private mSerie As String
Private mChaveAcesso As String
Private mNumeroTanque As Integer
Private mFormaPagamento As String
Private mValorFrete As Currency
Private mOrdem As Integer
Private mCST As String
Private mCFOP As String
Private mCodigoProduto As Integer
Private mValorBCICMS As Currency
Private mAliquotaICMS As Currency
Private mValorICMS As Currency
Private mValorBCICMSST As Currency
Private mValorICMSST As Currency
Private mValorNaoTributadoReducaoBCICMS As Currency
Private mDataEmissao As Date

Private rsEntradaCombustivel As New adodb.Recordset
Private IntegracaoNuvem As New cIntegracaoNuvem


'Inicio das Propriedades da Classe
Public Property Get NomeTabela() As String
    NomeTabela = mNomeTabela
End Property
Public Property Let NomeTabela(ByVal Valor As String)
    mNomeTabela = Valor
End Property


Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get NumeroNota() As String
    NumeroNota = mNumeroNota
End Property
Public Property Let NumeroNota(ByVal Valor As String)
    mNumeroNota = Valor
End Property


Public Property Get ValorLitro() As Currency
    ValorLitro = mValorLitro
End Property
Public Property Let ValorLitro(ByVal Valor As Currency)
    mValorLitro = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorEntrada() As Currency
    ValorEntrada = mValorEntrada
End Property
Public Property Let ValorEntrada(ByVal Valor As Currency)
    mValorEntrada = Valor
End Property


Public Property Get CodigoFornecedor() As Integer
    CodigoFornecedor = mCodigoFornecedor
End Property
Public Property Let CodigoFornecedor(ByVal Valor As Integer)
    mCodigoFornecedor = Valor
End Property


Public Property Get TipoTransporte() As Integer
    TipoTransporte = mTipoTransporte
End Property
Public Property Let TipoTransporte(ByVal Valor As Integer)
    mTipoTransporte = Valor
End Property


Public Property Get Item() As Integer
    Item = mItem
End Property
Public Property Let Item(ByVal Valor As Integer)
    mItem = Valor
End Property


Public Property Get Modelo() As String
    Modelo = mModelo
End Property
Public Property Let Modelo(ByVal Valor As String)
    mModelo = Valor
End Property


Public Property Get Serie() As String
    Serie = mSerie
End Property
Public Property Let Serie(ByVal Valor As String)
    mSerie = Valor
End Property


Public Property Get ChaveAcesso() As String
    ChaveAcesso = mChaveAcesso
End Property
Public Property Let ChaveAcesso(ByVal Valor As String)
    mChaveAcesso = Valor
End Property


Public Property Get NumeroTanque() As Integer
    NumeroTanque = mNumeroTanque
End Property
Public Property Let NumeroTanque(ByVal Valor As Integer)
    mNumeroTanque = Valor
End Property


Public Property Get FormaPagamento() As String
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As String)
    mFormaPagamento = Valor
End Property


Public Property Get ValorFrete() As Currency
    ValorFrete = mValorFrete
End Property
Public Property Let ValorFrete(ByVal Valor As Currency)
    mValorFrete = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get CST() As String
    CST = mCST
End Property
Public Property Let CST(ByVal Valor As String)
    mCST = Valor
End Property


Public Property Get CFOP() As String
    CFOP = mCFOP
End Property
Public Property Let CFOP(ByVal Valor As String)
    mCFOP = Valor
End Property


Public Property Get CodigoProduto() As Integer
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Integer)
    mCodigoProduto = Valor
End Property


Public Property Get ValorBCICMS() As Currency
    ValorBCICMS = mValorBCICMS
End Property
Public Property Let ValorBCICMS(ByVal Valor As Currency)
    mValorBCICMS = Valor
End Property


Public Property Get AliquotaICMS() As Currency
    AliquotaICMS = mAliquotaICMS
End Property
Public Property Let AliquotaICMS(ByVal Valor As Currency)
    mAliquotaICMS = Valor
End Property


Public Property Get ValorICMS() As Currency
    ValorICMS = mValorICMS
End Property
Public Property Let ValorICMS(ByVal Valor As Currency)
    mValorICMS = Valor
End Property


Public Property Get ValorBCICMSST() As Currency
    ValorBCICMSST = mValorBCICMSST
End Property
Public Property Let ValorBCICMSST(ByVal Valor As Currency)
    mValorBCICMSST = Valor
End Property


Public Property Get ValorICMSST() As Currency
    ValorICMSST = mValorICMSST
End Property
Public Property Let ValorICMSST(ByVal Valor As Currency)
    mValorICMSST = Valor
End Property


Public Property Get ValorNaoTributadoReducaoBCICMS() As Currency
    ValorNaoTributadoReducaoBCICMS = mValorNaoTributadoReducaoBCICMS
End Property
Public Property Let ValorNaoTributadoReducaoBCICMS(ByVal Valor As Currency)
    mValorNaoTributadoReducaoBCICMS = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Numero da Nota] = " & preparaTexto(pNumeroNota)
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Tipo de Combustivel] DESC, [Numero da Nota] DESC, Ordem DESC")
    If Localizar(55) Then
        If rsEntradaCombustivel.RecordCount > 1 Then
            Do Until rsEntradaCombustivel.EOF
                If xRegIgual = False Then
                    If rsEntradaCombustivel("Data").Value = mData And rsEntradaCombustivel("Tipo de Combustivel").Value = mTipoCombustivel And rsEntradaCombustivel("Numero da Nota").Value = mNumeroNota And rsEntradaCombustivel("Ordem").Value = mOrdem Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsEntradaCombustivel.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Tipo de Combustivel] ASC, [Numero da Nota] ASC, Ordem ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Tipo de Combustivel] ASC, [Numero da Nota] ASC, Ordem ASC")
    If Localizar(55) Then
        If rsEntradaCombustivel.RecordCount > 1 Then
            Do Until rsEntradaCombustivel.EOF
                If xRegIgual = False Then
                    If rsEntradaCombustivel("Data").Value = mData And rsEntradaCombustivel("Tipo de Combustivel").Value = mTipoCombustivel And rsEntradaCombustivel("Numero da Nota").Value = mNumeroNota And rsEntradaCombustivel("Ordem").Value = mOrdem Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsEntradaCombustivel.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Tipo de Combustivel] DESC, [Numero da Nota] DESC, Ordem DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoCombustivel(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoCombustivel = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data < " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, [Tipo de Combustivel] DESC, [Numero da Nota] DESC, Ordem DESC")
    LocalizarUltimoCombustivel = Localizar(1)
    If LocalizarUltimoCombustivel Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function IdentidicarNotasInexistentes(ByVal pEmpresa As Integer, ByVal pDicChavesAcessoNotasDestinadas As Dictionary) As Dictionary
Dim xCondicao As String
Dim xDicNotasInexistentes As New Dictionary
Dim xChaveLocalizada As Boolean
Dim xIndice As Integer
Dim xChaveAcesso

On Error GoTo trata_erro


    Set IdentidicarNotasInexistentes = xDicNotasInexistentes
    xChaveLocalizada = False
    
    For Each xChaveAcesso In pDicChavesAcessoNotasDestinadas
    
        gSQL = ""
        
        xCondicao = " WHERE Empresa = " & pEmpresa
        xCondicao = xCondicao & " AND [Chave de Acesso] = " & preparaTexto(RetiraString(1, CStr(xChaveAcesso))) & " AND Ordem = " & RetiraString(2, xChaveAcesso)
        xCondicao = xCondicao & " AND Ordem = " & RetiraString(2, xChaveAcesso)
        
        gSQL = PreparaSQL(xCondicao, " ORDER BY [Chave de Acesso], Ordem")
    
        xChaveLocalizada = Localizar(1)
        
        If Not xChaveLocalizada Then
           Call xDicNotasInexistentes.Add(RetiraString(1, CStr(xChaveAcesso)) & "|@|" & RetiraString(2, xChaveAcesso) & "|@|" & RetiraString(3, xChaveAcesso) & "|@|", pDicChavesAcessoNotasDestinadas(xChaveAcesso)) 'RetiraString(2, xChaveAcesso))
        End If
       

    Next
    
    Set IdentidicarNotasInexistentes = xDicNotasInexistentes
    
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir(ByVal pGeraIntegracaoNuvem As Boolean) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mOrdem = ProximaOrdem
    gSQL = "INSERT INTO " & mNomeTabela & " ( Empresa, Data, [Tipo de Combustivel], [Numero da Nota], [Valor do Litro], "
    gSQL = gSQL & "Quantidade, [Valor da Entrada], [Codigo do Fornecedor], [Tipo de Transporte], Item, "
    gSQL = gSQL & "Modelo, Serie, [Chave de Acesso], [Numero do Tanque], [Forma de Pagamento], "
    gSQL = gSQL & "[Valor do Frete], Ordem, CST, CFOP, [Codigo do Produto], "
    gSQL = gSQL & "[Valor BC ICMS], [Aliquota ICMS], [Valor ICMS], [Valor BC ICMS ST], [Valor ICMS ST], "
    gSQL = gSQL & "[Valor Nao Tributado Reducao BC ICMS], [Data da Emissao] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlTexto(1, mNumeroNota, ", ")
    Call sqlValor(1, mValorLitro, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorEntrada, ", ")
    Call sqlNumero(1, mCodigoFornecedor, ", ")
    Call sqlNumero(1, mTipoTransporte, ", ")
    Call sqlNumero(1, mItem, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlTexto(1, mChaveAcesso, ", ")
    Call sqlNumero(1, mNumeroTanque, ", ")
    Call sqlTexto(1, mFormaPagamento, ", ")
    Call sqlValor(1, mValorFrete, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlTexto(1, mCST, ", ")
    Call sqlTexto(1, mCFOP, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorBCICMS, ", ")
    Call sqlValor(1, mAliquotaICMS, ", ")
    Call sqlValor(1, mValorICMS, ", ")
    Call sqlValor(1, mValorBCICMSST, ", ")
    Call sqlValor(1, mValorICMSST, ", ")
    Call sqlValor(1, mValorNaoTributadoReducaoBCICMS, ", ")
    Call sqlData(1, mDataEmissao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
        If pGeraIntegracaoNuvem Then
            Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mData, mTipoCombustivel, mNumeroNota, mCodigoFornecedor, mOrdem)
        End If
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pCodigoFornecedor As Integer, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlTexto(2, ", [Numero da Nota] = ", mNumeroNota)
    Call sqlValor(2, ", [Valor do Litro] = ", mValorLitro)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor da Entrada] = ", mValorEntrada)
    Call sqlNumero(2, ", [Codigo do Fornecedor] = ", mCodigoFornecedor)
    Call sqlNumero(2, ", [Tipo de Transporte] = ", mTipoTransporte)
    Call sqlNumero(2, ", Item = ", mItem)
    Call sqlTexto(2, ", Modelo = ", mModelo)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlTexto(2, ", [Chave de Acesso] = ", mChaveAcesso)
    Call sqlNumero(2, ", [Numero do Tanque] = ", mNumeroTanque)
    Call sqlTexto(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor do Frete] = ", mValorFrete)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlTexto(2, ", CST = ", mCST)
    Call sqlTexto(2, ", CFOP = ", mCFOP)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor BC ICMS] = ", mValorBCICMS)
    Call sqlValor(2, ", [Aliquota ICMS] = ", mAliquotaICMS)
    Call sqlValor(2, ", [Valor ICMS] = ", mValorICMS)
    Call sqlValor(2, ", [Valor BC ICMS ST] = ", mValorBCICMSST)
    Call sqlValor(2, ", [Valor ICMS ST] = ", mValorICMSST)
    Call sqlValor(2, ", [Valor Nao Tributado Reducao BC ICMS] = ", mValorNaoTributadoReducaoBCICMS)
    Call sqlData(2, ", [Data da Emissao] = ", mDataEmissao)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND [Numero da Nota] = " & preparaTexto(pNumeroNota)
    gSQL = gSQL & " AND [Codigo do Fornecedor] = " & pCodigoFornecedor
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
        Call IncluiIntegracaoNuvem(" EXCLUIR", pEmpresa, pData, pTipoCombustivel, pNumeroNota, pCodigoFornecedor, pOrdem)
        Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mData, mTipoCombustivel, mNumeroNota, mCodigoFornecedor, mOrdem)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarOld(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pCodigoFornecedor As Integer, ByVal pItem As Integer) As Boolean
'Esta rotina será usada somente para criar campo "Ordem" nesta tabela.
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarOld

    AlterarOld = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlTexto(2, ", [Numero da Nota] = ", mNumeroNota)
    Call sqlValor(2, ", [Valor do Litro] = ", mValorLitro)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor da Entrada] = ", mValorEntrada)
    Call sqlNumero(2, ", [Codigo do Fornecedor] = ", mCodigoFornecedor)
    Call sqlNumero(2, ", [Tipo de Transporte] = ", mTipoTransporte)
    Call sqlNumero(2, ", Item = ", mItem)
    Call sqlTexto(2, ", Modelo = ", mModelo)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlTexto(2, ", [Chave de Acesso] = ", mChaveAcesso)
    Call sqlNumero(2, ", [Numero do Tanque] = ", mNumeroTanque)
    Call sqlTexto(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlValor(2, ", [Valor do Frete] = ", mValorFrete)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND [Numero da Nota] = " & preparaTexto(pNumeroNota)
    gSQL = gSQL & " AND [Codigo do Fornecedor] = " & pCodigoFornecedor
    gSQL = gSQL & " AND Item = " & pItem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarOld = True
    End If
    Exit Function

Err_AlterarOld:
End Function


Public Function LocalizarCodigoOld(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pItem As Integer) As Boolean
'Esta rotina será usada somente para criar campo "Ordem" nesta tabela.
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigoOld = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Numero da Nota] = " & preparaTexto(pNumeroNota)
    xCondicao = xCondicao & " AND Item = " & pItem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigoOld = Localizar(1)
    If LocalizarCodigoOld Then
        AtribuiValor
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pCodigoFornecedor As Integer, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND [Numero da Nota] = " & preparaTexto(pNumeroNota)
    gSQL = gSQL & " AND [Codigo do Fornecedor] = " & pCodigoFornecedor
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
        Call IncluiIntegracaoNuvem("EXCLUIR", pEmpresa, pData, pTipoCombustivel, pNumeroNota, pCodigoFornecedor, pOrdem)
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluirMesAno(ByVal pEmpresa As Integer, ByVal pData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirMesAno = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND YEAR(Data) = " & Format(pData, "yyyy")
    gSQL = gSQL & " AND MONTH(Data) = " & Format(pData, "MM")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirMesAno = True
        Call IncluiIntegracaoNuvem("EXCLUIR", pEmpresa, pData, "**", "0", 0, 0)
    End If
    Exit Function

Err_Excluir:
End Function


'Transfere dados Para o LMC
'Public Function TransfereDados(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTabelaOrigem As String) As Boolean
'Dim xSQL As String
'Dim rsEntradaCombustivelOrigem As New adodb.Recordset
'
'On Error GoTo trata_erro
'
'    TransfereDados = False
'
'    xSQL = " WHERE Empresa = " & pEmpresa
'    xSQL = xSQL & " AND Data >= " & preparaData(pData)
'    gSQL = PreparaSQL(xSQL, "ORDER BY Data ASC, [Tipo de Combustivel] ASC, [Numero da Nota] ASC")
'    xSQL = Replace(gSQL, mNomeTabela, pTabelaOrigem)
'
'    Set rsEntradaCombustivelOrigem = New adodb.Recordset
'    rsEntradaCombustivelOrigem.CursorLocation = adUseClient
'    rsEntradaCombustivelOrigem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
'    If Not rsEntradaCombustivelOrigem.EOF Then
'        rsEntradaCombustivelOrigem.MoveFirst
'        Do Until rsEntradaCombustivelOrigem.EOF
'            If Trim(rsEntradaCombustivelOrigem("Numero da Nota").Value) <> "1" Then
'                Empresa = rsEntradaCombustivelOrigem("Empresa").Value
'                Data = rsEntradaCombustivelOrigem("Data").Value
'                TipoCombustivel = rsEntradaCombustivelOrigem("Tipo de Combustivel").Value
'                NumeroNota = rsEntradaCombustivelOrigem("Numero da Nota").Value
'                ValorLitro = rsEntradaCombustivelOrigem("Valor do Litro").Value
'                Quantidade = rsEntradaCombustivelOrigem("Quantidade").Value
'                ValorEntrada = rsEntradaCombustivelOrigem("Valor da Entrada").Value
'                CodigoFornecedor = rsEntradaCombustivelOrigem("Codigo do Fornecedor").Value
'                TipoTransporte = rsEntradaCombustivelOrigem("Tipo de Transporte").Value
'                If Not Incluir Then
'                    MsgBox "Não foi possível incluir registro!", vbInformation, "Erro de Integridade!"
'                    rsEntradaCombustivelOrigem.Close
'                    Set rsEntradaCombustivelOrigem = Nothing
'                    Exit Function
'                End If
'            End If
'            rsEntradaCombustivelOrigem.MoveNext
'        Loop
'        TransfereDados = True
'    End If
'    rsEntradaCombustivelOrigem.Close
'    Set rsEntradaCombustivelOrigem = Nothing
'    Exit Function
'
'trata_erro:
'    MsgBox Err.Number & " - " & Err.Description
'End Function


Public Function TransfereDadosLMC(ByVal pEmpresa As Integer, ByVal pSomenteVerifica As Boolean) As Boolean
    Dim xData As Date
    Dim rsAuxiliar As New adodb.Recordset

On Error GoTo trata_erro

    TransfereDadosLMC = False
    
    'Busca Ultima Data do Movimento LMC
    xData = UltimaData(pEmpresa)
    If xData = "00:00:00" Then
        Exit Function
    End If
    
    'Busca Movimento de Entrada para Transferencia
    xData = xData + 1
    gSQL = ""
    gSQL = gSQL & "   SELECT Empresa, Data, [Tipo de Combustivel], [Numero da Nota],"
    gSQL = gSQL & "          [Valor do Litro], Quantidade, [Valor da Entrada],"
    gSQL = gSQL & "          [Codigo do Fornecedor], [Tipo de Transporte], Item, Modelo, Serie, "
    gSQL = gSQL & "          [Chave de Acesso], [Numero do Tanque], [Forma de Pagamento] , "
    gSQL = gSQL & "          [Valor do Frete], Ordem, CST, CFOP, [Codigo do Produto], "
    gSQL = gSQL & "          [Valor BC ICMS], [Aliquota ICMS], [Valor ICMS], [Valor BC ICMS ST], "
    gSQL = gSQL & "          [Valor ICMS ST], [Valor Nao Tributado Reducao BC ICMS], [Data da Emissao]"
    gSQL = gSQL & "     FROM Entrada_Combustivel"
    gSQL = gSQL & "    WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "      AND Data >= " & preparaData(xData)
    gSQL = gSQL & " ORDER BY Data, [Numero da Nota], [Tipo de Combustivel]"
    Set rsAuxiliar = New adodb.Recordset
    rsAuxiliar.CursorLocation = adUseClient
    rsAuxiliar.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    'Grava Entradas
    If Not rsAuxiliar.EOF Then
        If pSomenteVerifica Then
            TransfereDadosLMC = True
            Exit Function
        End If
        Do Until rsAuxiliar.EOF
            Empresa = rsAuxiliar("Empresa").Value
            Data = rsAuxiliar("Data").Value
            TipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
            NumeroNota = rsAuxiliar("Numero da Nota").Value
            ValorLitro = rsAuxiliar("Valor do Litro").Value
            Quantidade = rsAuxiliar("Quantidade").Value
            ValorEntrada = rsAuxiliar("Valor da Entrada").Value
            CodigoFornecedor = rsAuxiliar("Codigo do Fornecedor").Value
            TipoTransporte = rsAuxiliar("Tipo de Transporte").Value
            Item = rsAuxiliar("Item").Value
            Modelo = rsAuxiliar("Modelo").Value
            Serie = rsAuxiliar("Serie").Value
            ChaveAcesso = rsAuxiliar("Chave de Acesso").Value
            NumeroTanque = rsAuxiliar("Numero do Tanque").Value
            FormaPagamento = rsAuxiliar("Forma de Pagamento").Value
            ValorFrete = rsAuxiliar("Valor do Frete").Value
            Ordem = rsAuxiliar("Ordem").Value
            CST = rsAuxiliar("CST").Value
            CFOP = rsAuxiliar("CFOP").Value
            CodigoProduto = rsAuxiliar("Codigo do Produto").Value
            ValorBCICMS = rsAuxiliar("Valor BC ICMS").Value
            AliquotaICMS = rsAuxiliar("Aliquota ICMS").Value
            ValorICMS = rsAuxiliar("Valor ICMS").Value
            ValorBCICMSST = rsAuxiliar("Valor BC ICMS ST").Value
            ValorICMSST = rsAuxiliar("Valor ICMS ST").Value
            ValorNaoTributadoReducaoBCICMS = rsAuxiliar("Valor Nao Tributado Reducao BC ICMS").Value
            DataEmissao = rsAuxiliar("Data da Emissao").Value
            If Not Incluir(True) Then
                MsgBox "Registro de Entrada de Combustível não foi gravado!", vbInformation, "Erro de Integridade"
            End If
            rsAuxiliar.MoveNext
        Loop
        TransfereDadosLMC = True
    End If
    
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function UltimaData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro

    UltimaData = "01/01/1900"
    gSQL = "SELECT Data FROM Entrada_Combustivel_LMC WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Data DESC, [Numero da Nota] DESC, [Tipo de Combustivel] DESC"
    If Localizar(1) Then
        UltimaData = rsEntradaCombustivel("Data").Value
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Tipo de Combustivel], [Numero da Nota], [Valor do Litro], Quantidade, [Valor da Entrada], [Codigo do Fornecedor], [Tipo de Transporte], Item, Modelo, Serie, [Chave de Acesso], [Numero do Tanque], [Forma de Pagamento], [Valor do Frete], Ordem, CST, CFOP, [Codigo do Produto], [Valor BC ICMS], [Aliquota ICMS], [Valor ICMS], [Valor BC ICMS ST], [Valor ICMS ST], [Valor Nao Tributado Reducao BC ICMS], [Data da Emissao] FROM " & mNomeTabela & " " & xCondicao & " " & xTipoOrdem)
End Function


Public Function DadosEntradaData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As String
Dim xSQL As String

On Error GoTo trata_erro

    DadosEntradaData = "0|@|0|@|"
    xSQL = ""
    xSQL = xSQL & " SELECT SUM(Quantidade) AS QuantidadeEntrada,"
    xSQL = xSQL & "        SUM([Valor da Entrada]) AS ValorEntrada"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data = " & preparaData(pData)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    Set rsEntradaCombustivel = New adodb.Recordset
    rsEntradaCombustivel.CursorLocation = adUseClient
    rsEntradaCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(rsEntradaCombustivel("QuantidadeEntrada").Value) Then
        DadosEntradaData = rsEntradaCombustivel("QuantidadeEntrada").Value & "|@|" & Format(rsEntradaCombustivel("ValorEntrada").Value / rsEntradaCombustivel("QuantidadeEntrada").Value, "00000000.0000") & "|@|"
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function



Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEntradaCombustivel = New adodb.Recordset
    rsEntradaCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEntradaCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEntradaCombustivel("Empresa").Value
    Data = rsEntradaCombustivel("Data").Value
    TipoCombustivel = rsEntradaCombustivel("Tipo de Combustivel").Value
    NumeroNota = rsEntradaCombustivel("Numero da Nota").Value
    ValorLitro = rsEntradaCombustivel("Valor do Litro").Value
    Quantidade = rsEntradaCombustivel("Quantidade").Value
    ValorEntrada = rsEntradaCombustivel("Valor da Entrada").Value
    CodigoFornecedor = rsEntradaCombustivel("Codigo do Fornecedor").Value
    TipoTransporte = rsEntradaCombustivel("Tipo de Transporte").Value
    Item = rsEntradaCombustivel("Item").Value
    Modelo = rsEntradaCombustivel("Modelo").Value
    Serie = rsEntradaCombustivel("Serie").Value
    ChaveAcesso = rsEntradaCombustivel("Chave de Acesso").Value
    NumeroTanque = rsEntradaCombustivel("Numero do Tanque").Value
    FormaPagamento = rsEntradaCombustivel("Forma de Pagamento").Value
    ValorFrete = rsEntradaCombustivel("Valor do Frete").Value
    Ordem = rsEntradaCombustivel("Ordem").Value
    CST = rsEntradaCombustivel("CST").Value
    CFOP = rsEntradaCombustivel("CFOP").Value
    CodigoProduto = rsEntradaCombustivel("Codigo do Produto").Value
    ValorBCICMS = rsEntradaCombustivel("Valor BC ICMS").Value
    AliquotaICMS = rsEntradaCombustivel("Aliquota ICMS").Value
    ValorICMS = rsEntradaCombustivel("Valor ICMS").Value
    ValorBCICMSST = rsEntradaCombustivel("Valor BC ICMS ST").Value
    ValorICMSST = rsEntradaCombustivel("Valor ICMS ST").Value
    ValorNaoTributadoReducaoBCICMS = rsEntradaCombustivel("Valor Nao Tributado Reducao BC ICMS").Value
    DataEmissao = rsEntradaCombustivel("Data da Emissao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function


'Retorna código do fornecedor da última Nota de Entrada de Combustível
Public Function RetornaUltimoFornecedor(ByVal pEmpresa As Integer) As Integer
Dim xSQL As String

On Error GoTo trata_erro

    RetornaUltimoFornecedor = 1
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 1 [Codigo do Fornecedor]"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data <= " & preparaData(CDate("31/12/2500"))
    xSQL = xSQL & "  ORDER BY Data DESC, [Numero da Nota] DESC"
    
    Set rsEntradaCombustivel = New adodb.Recordset
    rsEntradaCombustivel.CursorLocation = adUseClient
    rsEntradaCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaCombustivel.EOF Then
        RetornaUltimoFornecedor = rsEntradaCombustivel("Codigo do Fornecedor").Value
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalEntradaPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pTanque As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalEntradaPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS Quantidade"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If pTanque > 0 Then
        xSQL = xSQL & "    AND [Numero do Tanque] = " & Str(pTanque)
    End If
    
    Set rsEntradaCombustivel = New adodb.Recordset
    rsEntradaCombustivel.CursorLocation = adUseClient
    rsEntradaCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaCombustivel.EOF Then
        If Not IsNull(rsEntradaCombustivel("Quantidade").Value) Then
            TotalEntradaPeriodo = rsEntradaCombustivel("Quantidade").Value
        End If
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalEntradaPeriodoSN(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalEntradaPeriodoSN = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS Quantidade"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xSQL = xSQL & "    AND RTRIM([Numero da Nota]) = 1"
    
    Set rsEntradaCombustivel = New adodb.Recordset
    rsEntradaCombustivel.CursorLocation = adUseClient
    rsEntradaCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaCombustivel.EOF Then
        If Not IsNull(rsEntradaCombustivel("Quantidade").Value) Then
            TotalEntradaPeriodoSN = rsEntradaCombustivel("Quantidade").Value
        End If
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Sub IncluiIntegracaoNuvem(ByVal pTipoOperacao As String, ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroNota As String, ByVal pCodigoFornecedor As Integer, ByVal pOrdem As Integer)
    If mNomeTabela = "Entrada_Combustivel_LMC" Then
        IntegracaoNuvem.Empresa = mEmpresa
        IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
        IntegracaoNuvem.NomeTabela = mNomeTabela
        IntegracaoNuvem.ChaveAcesso = pEmpresa & "|@|" & pData & "|@|" & pTipoCombustivel & "|@|" & pNumeroNota & "|@|" & pCodigoFornecedor & "|@|" & pOrdem & "|@|"
        IntegracaoNuvem.TipoOperacao = pTipoOperacao
        IntegracaoNuvem.IntegradoEm = CDate("00:00:00")
        If Not IntegracaoNuvem.Incluir Then
            Call CriaLogCadastroDll2("cEntradaCombustivel:IncluiIntegracaoNuvem - Erro não identificado.", Err.Description, "Tabela=" & IntegracaoNuvem.NomeTabela & " - Chv.Acesso=" & IntegracaoNuvem.ChaveAcesso)
        End If
    End If
End Sub
       

Private Function ProximaOrdem() As Integer

On Error GoTo Err_ProximaOrdem

    'Pega Ordem do Proximo Movimento
    ProximaOrdem = 1
    gSQL = ""
    gSQL = gSQL & "SELECT Ordem"
    gSQL = gSQL & "  FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(mData)
    gSQL = gSQL & "   AND [Numero da Nota] = " & preparaTexto(mNumeroNota)
    gSQL = gSQL & " ORDER BY Ordem DESC"
    If Localizar(1) Then
        ProximaOrdem = rsEntradaCombustivel("Ordem").Value + 1
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    Exit Function

Err_ProximaOrdem:
End Function

