VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntradaProdutoCabecalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataEntrada As Date
Private mNumeroDocumento As String
Private mCodigoFornecedor As Integer
Private mTipoEntrada As String
Private mSubEstoque As Integer
Private mTotalProduto As Currency
Private mDesconto As Currency
Private mSubstituicao As Currency
Private mOutros As Currency
Private mTotalNota As Currency
Private mDataDigitacao As Date
Private mCodigoUsuario As Integer
Private mObservacao As String
Private mCFOP As String
Private mModelo As String
Private mSerie As String
Private mChaveAcessoNFe As String
Private mFormaPagamento As String
Private mTipoFrete As String
Private mValorFrete As Currency
Private mValorBCICMS As Currency
Private mAliquotaICMS As Currency
Private mValorICMS As Currency
Private mValorBCICMSST As Currency
Private mValorICMSST As Currency
Private mIncidePisCofins As Boolean

Private rsEntradaProdutoCabecalho As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataEntrada() As Date
    DataEntrada = mDataEntrada
End Property
Public Property Let DataEntrada(ByVal Valor As Date)
    mDataEntrada = Valor
End Property


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get CodigoFornecedor() As Integer
    CodigoFornecedor = mCodigoFornecedor
End Property
Public Property Let CodigoFornecedor(ByVal Valor As Integer)
    mCodigoFornecedor = Valor
End Property


Public Property Get TipoEntrada() As String
    TipoEntrada = mTipoEntrada
End Property
Public Property Let TipoEntrada(ByVal Valor As String)
    mTipoEntrada = Valor
End Property


Public Property Get SubEstoque() As Integer
    SubEstoque = mSubEstoque
End Property
Public Property Let SubEstoque(ByVal Valor As Integer)
    mSubEstoque = Valor
End Property


Public Property Get TotalProduto() As Currency
    TotalProduto = mTotalProduto
End Property
Public Property Let TotalProduto(ByVal Valor As Currency)
    mTotalProduto = Valor
End Property


Public Property Get Desconto() As Currency
    Desconto = mDesconto
End Property
Public Property Let Desconto(ByVal Valor As Currency)
    mDesconto = Valor
End Property


Public Property Get Substituicao() As Currency
    Substituicao = mSubstituicao
End Property
Public Property Let Substituicao(ByVal Valor As Currency)
    mSubstituicao = Valor
End Property


Public Property Get Outros() As Currency
    Outros = mOutros
End Property
Public Property Let Outros(ByVal Valor As Currency)
    mOutros = Valor
End Property


Public Property Get TotalNota() As Currency
    TotalNota = mTotalNota
End Property
Public Property Let TotalNota(ByVal Valor As Currency)
    mTotalNota = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get CFOP() As String
    CFOP = mCFOP
End Property
Public Property Let CFOP(ByVal Valor As String)
    mCFOP = Valor
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


Public Property Get ChaveAcessoNFe() As String
    ChaveAcessoNFe = mChaveAcessoNFe
End Property
Public Property Let ChaveAcessoNFe(ByVal Valor As String)
    mChaveAcessoNFe = Valor
End Property


Public Property Get FormaPagamento() As String
    FormaPagamento = mFormaPagamento
End Property
Public Property Let FormaPagamento(ByVal Valor As String)
    mFormaPagamento = Valor
End Property


Public Property Get TipoFrete() As String
    TipoFrete = mTipoFrete
End Property
Public Property Let TipoFrete(ByVal Valor As String)
    mTipoFrete = Valor
End Property


Public Property Get ValorFrete() As Currency
    ValorFrete = mValorFrete
End Property
Public Property Let ValorFrete(ByVal Valor As Currency)
    mValorFrete = Valor
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


Public Property Get IncidePisCofins() As Boolean
    IncidePisCofins = mIncidePisCofins
End Property
Public Property Let IncidePisCofins(ByVal Valor As Boolean)
    mIncidePisCofins = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pNumeroDocumento As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    xCondicao = xCondicao & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsEntradaProdutoCabecalho.Close
    Set rsEntradaProdutoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data da Entrada] <= " & preparaData(mDataEntrada)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] DESC, [Numero do Documento] DESC")
    If Localizar(30) Then
        If rsEntradaProdutoCabecalho.RecordCount > 1 Then
            Do Until rsEntradaProdutoCabecalho.EOF
                If xRegIgual = False Then
                    If rsEntradaProdutoCabecalho("Data da Entrada").Value = mDataEntrada And rsEntradaProdutoCabecalho("Numero do Documento").Value = mNumeroDocumento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    AtribuiValor
                    Exit Do
                End If
                rsEntradaProdutoCabecalho.MoveNext
            Loop
        End If
    End If
    rsEntradaProdutoCabecalho.Close
    Set rsEntradaProdutoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] ASC, [Numero do Documento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsEntradaProdutoCabecalho.Close
    Set rsEntradaProdutoCabecalho = Nothing
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
    xCondicao = xCondicao & " AND [Data da Entrada] >= " & preparaData(mDataEntrada)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] ASC, [Numero do Documento] ASC")
    If Localizar(30) Then
        If rsEntradaProdutoCabecalho.RecordCount > 1 Then
            Do Until rsEntradaProdutoCabecalho.EOF
                If xRegIgual = False Then
                    If rsEntradaProdutoCabecalho("Data da Entrada").Value = mDataEntrada And rsEntradaProdutoCabecalho("Numero do Documento").Value = mNumeroDocumento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    AtribuiValor
                    Exit Do
                End If
                rsEntradaProdutoCabecalho.MoveNext
            Loop
        End If
    End If
    rsEntradaProdutoCabecalho.Close
    Set rsEntradaProdutoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] DESC, [Numero do Documento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsEntradaProdutoCabecalho.Close
    Set rsEntradaProdutoCabecalho = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO EntradaProdutoCabecalho ( Empresa, [Data da Entrada], [Numero do Documento], [Codigo do Fornecedor], [Tipo da Entrada], "
    gSQL = gSQL & "SubEstoque, [Total do Produto], Desconto, Substituicao, Outros, "
    gSQL = gSQL & "[Total da Nota], [Data da Digitacao], [Codigo do Usuario], Observacao, CFOP, "
    gSQL = gSQL & "Modelo, Serie, [Chave de Acesso NFe], [Forma de Pagamento], [Tipo do Frete], "
    gSQL = gSQL & "[Valor do Frete], [Valor BC ICMS], [Aliquota ICMS], [Valor ICMS], [Valor BC ICMS ST], "
    gSQL = gSQL & "[Valor ICMS ST] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEntrada, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoFornecedor, ", ")
    Call sqlTexto(1, mTipoEntrada, ", ")
    Call sqlNumero(1, mSubEstoque, ", ")
    Call sqlValor(1, mTotalProduto, ", ")
    Call sqlValor(1, mDesconto, ", ")
    Call sqlValor(1, mSubstituicao, ", ")
    Call sqlValor(1, mOutros, ", ")
    Call sqlValor(1, mTotalNota, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlTexto(1, mObservacao, ", ")
    Call sqlTexto(1, mCFOP, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlTexto(1, mChaveAcessoNFe, ", ")
    Call sqlTexto(1, mFormaPagamento, ", ")
    Call sqlTexto(1, mTipoFrete, ", ")
    Call sqlValor(1, mValorFrete, ", ")
    Call sqlValor(1, mValorBCICMS, ", ")
    Call sqlValor(1, mAliquotaICMS, ", ")
    Call sqlValor(1, mValorICMS, ", ")
    Call sqlValor(1, mValorBCICMSST, ", ")
    Call sqlValor(1, mValorICMSST, ", ")
    Call sqlBoolean(1, mIncidePisCofins, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pNumeroDocumento As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE EntradaProdutoCabecalho SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data da Entrada] = ", mDataEntrada)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Fornecedor] = ", mCodigoFornecedor)
    Call sqlTexto(2, ", [Tipo da Entrada] = ", mTipoEntrada)
    Call sqlNumero(2, ", SubEstoque = ", mSubEstoque)
    Call sqlValor(2, ", [Total do Produto] = ", mTotalProduto)
    Call sqlValor(2, ", Desconto = ", mDesconto)
    Call sqlValor(2, ", Substituicao = ", mSubstituicao)
    Call sqlValor(2, ", Outros = ", mOutros)
    Call sqlValor(2, ", [Total da Nota] = ", mTotalNota)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    Call sqlTexto(2, ", CFOP = ", mCFOP)
    Call sqlTexto(2, ", Modelo = ", mModelo)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlTexto(2, ", [Chave de Acesso NFe] = ", mChaveAcessoNFe)
    Call sqlTexto(2, ", [Forma de Pagamento] = ", mFormaPagamento)
    Call sqlTexto(2, ", [Tipo do Frete] = ", mTipoFrete)
    Call sqlValor(2, ", [Valor do Frete] = ", mValorFrete)
    Call sqlValor(2, ", [Valor BC ICMS] = ", mValorBCICMS)
    Call sqlValor(2, ", [Aliquota ICMS] = ", mAliquotaICMS)
    Call sqlValor(2, ", [Valor ICMS] = ", mValorICMS)
    Call sqlValor(2, ", [Valor BC ICMS ST] = ", mValorBCICMSST)
    Call sqlValor(2, ", [Valor ICMS ST] = ", mValorICMSST)
    Call sqlValor(2, ", [Incide Pis Cofins] = ", mIncidePisCofins)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    gSQL = gSQL & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pNumeroDocumento As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM EntradaProdutoCabecalho"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    gSQL = gSQL & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data da Entrada], [Numero do Documento], [Codigo do Fornecedor], [Tipo da Entrada], SubEstoque, [Total do Produto], Desconto, Substituicao, Outros, [Total da Nota], [Data da Digitacao], [Codigo do Usuario], Observacao, CFOP, Modelo, Serie, [Chave de Acesso NFe], [Forma de Pagamento], [Tipo do Frete], [Valor do Frete], [Valor BC ICMS], [Aliquota ICMS], [Valor ICMS], [Valor BC ICMS ST], [Valor ICMS ST], [Incide Pis Cofins] FROM EntradaProdutoCabecalho" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEntradaProdutoCabecalho = New adodb.Recordset
    rsEntradaProdutoCabecalho.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEntradaProdutoCabecalho.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaProdutoCabecalho.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEntradaProdutoCabecalho("Empresa").Value
    DataEntrada = rsEntradaProdutoCabecalho("Data da Entrada").Value
    NumeroDocumento = rsEntradaProdutoCabecalho("Numero do Documento").Value
    CodigoFornecedor = rsEntradaProdutoCabecalho("Codigo do Fornecedor").Value
    TipoEntrada = rsEntradaProdutoCabecalho("Tipo da Entrada").Value
    SubEstoque = rsEntradaProdutoCabecalho("SubEstoque").Value
    TotalProduto = rsEntradaProdutoCabecalho("Total do Produto").Value
    Desconto = rsEntradaProdutoCabecalho("Desconto").Value
    Substituicao = rsEntradaProdutoCabecalho("Substituicao").Value
    Outros = rsEntradaProdutoCabecalho("Outros").Value
    TotalNota = rsEntradaProdutoCabecalho("Total da Nota").Value
    DataDigitacao = rsEntradaProdutoCabecalho("Data da Digitacao").Value
    CodigoUsuario = rsEntradaProdutoCabecalho("Codigo do Usuario").Value
    Observacao = rsEntradaProdutoCabecalho("Observacao").Value
    CFOP = rsEntradaProdutoCabecalho("CFOP").Value
    Modelo = rsEntradaProdutoCabecalho("Modelo").Value
    Serie = rsEntradaProdutoCabecalho("Serie").Value
    ChaveAcessoNFe = rsEntradaProdutoCabecalho("Chave de Acesso NFe").Value
    FormaPagamento = rsEntradaProdutoCabecalho("Forma de Pagamento").Value
    TipoFrete = rsEntradaProdutoCabecalho("Tipo do Frete").Value
    ValorFrete = rsEntradaProdutoCabecalho("Valor do Frete").Value
    ValorBCICMS = rsEntradaProdutoCabecalho("Valor BC ICMS").Value
    AliquotaICMS = rsEntradaProdutoCabecalho("Aliquota ICMS").Value
    ValorICMS = rsEntradaProdutoCabecalho("Valor ICMS").Value
    ValorBCICMSST = rsEntradaProdutoCabecalho("Valor BC ICMS ST").Value
    ValorICMSST = rsEntradaProdutoCabecalho("Valor ICMS ST").Value
    IncidePisCofins = rsEntradaProdutoCabecalho("Incide Pis Cofins").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

