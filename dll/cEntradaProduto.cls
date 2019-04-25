VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntradaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataEntrada As Date
Private mCodigoProduto As Long
Private mNumeroDocumento As String
Private mTipoEntrada As String
Private mPrecoCusto As Currency
Private mQuantidade As Currency
Private mTotalCusto As Currency
Private mCodigoFornecedor As Integer
Private mDataDigitacao As Date
Private mObservacao As String
Private mCodigoUsuario As Integer
Private mCustoUnitarioBruto As Currency
Private mSubEstoque As Integer
Private mCFOP As String
Private mModelo As String
Private mSerie As String
Private mOrdem As Integer
Private mBaseCalculoICMS As Currency
Private mIcms As Currency
Private mAliquotaICMS As Currency

Private rsEntradaProduto As New adodb.Recordset


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


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get TipoEntrada() As String
    TipoEntrada = mTipoEntrada
End Property
Public Property Let TipoEntrada(ByVal Valor As String)
    mTipoEntrada = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get TotalCusto() As Currency
    TotalCusto = mTotalCusto
End Property
Public Property Let TotalCusto(ByVal Valor As Currency)
    mTotalCusto = Valor
End Property


Public Property Get CodigoFornecedor() As Integer
    CodigoFornecedor = mCodigoFornecedor
End Property
Public Property Let CodigoFornecedor(ByVal Valor As Integer)
    mCodigoFornecedor = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property


Public Property Get CustoUnitarioBruto() As Currency
    CustoUnitarioBruto = mCustoUnitarioBruto
End Property
Public Property Let CustoUnitarioBruto(ByVal Valor As Currency)
    mCustoUnitarioBruto = Valor
End Property


Public Property Get SubEstoque() As Integer
    SubEstoque = mSubEstoque
End Property
Public Property Let SubEstoque(ByVal Valor As Integer)
    mSubEstoque = Valor
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


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get BaseCalculoICMS() As Currency
    BaseCalculoICMS = mBaseCalculoICMS
End Property
Public Property Let BaseCalculoICMS(ByVal Valor As Currency)
    mBaseCalculoICMS = Valor
End Property


Public Property Get Icms() As Currency
    Icms = mIcms
End Property
Public Property Let Icms(ByVal Valor As Currency)
    mIcms = Valor
End Property


Public Property Get AliquotaICMS() As Currency
    AliquotaICMS = mAliquotaICMS
End Property
Public Property Let AliquotaICMS(ByVal Valor As Currency)
    mAliquotaICMS = Valor
End Property



'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pCodigoProduto As Long, ByVal pNumeroDocumento As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    xCondicao = xCondicao & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
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
    xCondicao = xCondicao & " AND [Data da Entrada] <= " & preparaData(mDataEntrada)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] DESC, [Numero do Documento] DESC, [Codigo do Produto] DESC")
    If Localizar(200) Then
        If rsEntradaProduto.RecordCount > 1 Then
            Do Until rsEntradaProduto.EOF
                If xRegIgual = False Then
                    If rsEntradaProduto("Data da Entrada").Value = mDataEntrada And rsEntradaProduto("Numero do Documento").Value = mNumeroDocumento And rsEntradaProduto("Codigo do Produto").Value = mCodigoProduto Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsEntradaProduto.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] ASC, [Numero do Documento] ASC, [Codigo do Produto] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] ASC, [Numero do Documento] ASC, [Codigo do Produto] ASC")
    If Localizar(200) Then
        If rsEntradaProduto.RecordCount > 1 Then
            Do Until rsEntradaProduto.EOF
                If xRegIgual = False Then
                    If rsEntradaProduto("Data da Entrada").Value = mDataEntrada And rsEntradaProduto("Numero do Documento").Value = mNumeroDocumento And rsEntradaProduto("Codigo do Produto").Value = mCodigoProduto Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsEntradaProduto.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] DESC, [Numero do Documento] DESC, [Codigo do Produto] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoProduto(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoProduto = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Entrada] DESC")
    If Localizar(1) Then
        LocalizarUltimoProduto = True
        AtribuiValor
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
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
    
        gSQL = "SELECT [Entrada_Produto].[Empresa] ,[Entrada_Produto].[Data da Entrada] ,[Entrada_Produto].[Codigo do Produto] ,"
        gSQL = gSQL & " [Entrada_Produto].[Numero do Documento] ,[Entrada_Produto].[Tipo da Entrada] ,[Entrada_Produto].[Preco de Custo] ,[Entrada_Produto].[Quantidade] ,[Entrada_Produto].[Total do Custo] ,[Entrada_Produto].[Codigo do Fornecedor] ,[Entrada_Produto].[Data da Digitacao] ,[Entrada_Produto].[Observacao] ,[Entrada_Produto].[Codigo do Usuario] ,[Entrada_Produto].[Custo Unitario Bruto] ,[Entrada_Produto].[SubEstoque] ,[Entrada_Produto].[CFOP] ,[Entrada_Produto].[Modelo] ,[Entrada_Produto].[Serie] ,[Entrada_Produto].[Ordem] ,[Entrada_Produto].[CST] ,[Entrada_Produto].[Base de calculo ICMS] ,[Entrada_Produto].[ICMS] ,[Entrada_Produto].[Aliquota ICMS] ,[Entrada_Produto].[Data da Emissao] ,[Entrada_Produto].[ICMS ST] ,[Entrada_Produto].[Aliquota ICMS ST] ,[Entrada_Produto].[Base de calculo ICMS ST],[EntradaProdutoCabecalho].[Chave de Acesso NFe] "
        gSQL = gSQL & " FROM Entrada_Produto, EntradaProdutoCabecalho"
        gSQL = gSQL & " WHERE Entrada_Produto.[Numero do Documento] = EntradaProdutoCabecalho.[Numero do Documento]"
        gSQL = gSQL & " AND Entrada_Produto.[Data da Emissao] = EntradaProdutoCabecalho.[Data da Emissao]"
        gSQL = gSQL & " AND [EntradaProdutoCabecalho].[Chave de Acesso NFe] = " & preparaTexto(RetiraString(1, xChaveAcesso))
        gSQL = gSQL & " AND [Entrada_Produto].Ordem = " & RetiraString(2, xChaveAcesso)
        gSQL = gSQL & " ORDER BY Entrada_Produto.[Numero do Documento], Entrada_Produto.Ordem"
        
        xChaveLocalizada = Localizar(1)
        
        If Not xChaveLocalizada Then
           Call xDicNotasInexistentes.Add(RetiraString(1, CStr(xChaveAcesso)) & "|@|" & RetiraString(2, xChaveAcesso) & "|@|" & RetiraString(3, xChaveAcesso) & "|@|", pDicChavesAcessoNotasDestinadas(xChaveAcesso))  'RetiraString(2, xChaveAcesso))
        End If
        
        rsEntradaProduto.Close

    Next
    
    Set IdentidicarNotasInexistentes = xDicNotasInexistentes
    
    Set rsEntradaProduto = Nothing
    
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Entrada_Produto ( Empresa, [Data da Entrada], [Codigo do Produto], [Numero do Documento], [Tipo da Entrada], "
    gSQL = gSQL & "[Preco de Custo], Quantidade, [Total do Custo], [Codigo do Fornecedor], [Data da Digitacao], "
    gSQL = gSQL & "Observacao, [Codigo do Usuario], [Custo Unitario Bruto], SubEstoque, CFOP, Modelo, Serie, Ordem, [Incide Pis Confins] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEntrada, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlTexto(1, mTipoEntrada, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mTotalCusto, ", ")
    Call sqlNumero(1, mCodigoFornecedor, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlTexto(1, mObservacao, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlValor(1, mCustoUnitarioBruto, ", ")
    Call sqlNumero(1, mSubEstoque, ", ")
    Call sqlTexto(1, mCFOP, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlValor(1, mBaseCalculoICMS, ", ")
    Call sqlValor(1, mIcms, ", ")
    Call sqlValor(1, mAliquotaICMS, ", ")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function

Public Function Alterar(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pCodigoProduto As Long, ByVal pNumeroDocumento As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Entrada_Produto SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data da Entrada] = ", mDataEntrada)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlTexto(2, ", [Tipo da Entrada] = ", mTipoEntrada)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Total do Custo] = ", mTotalCusto)
    Call sqlNumero(2, ", [Codigo do Fornecedor] = ", mCodigoFornecedor)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlValor(2, ", [Custo Unitario Bruto] = ", mCustoUnitarioBruto)
    Call sqlNumero(2, ", SubEstoque = ", mSubEstoque)
    Call sqlTexto(2, ", CFOP = ", mCFOP)
    Call sqlTexto(2, ", Modelo = ", mModelo)
    Call sqlTexto(2, ", Serie = ", mSerie)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlValor(2, ", [Base de Calculo ICMS] = ", mBaseCalculoICMS)
    Call sqlValor(2, ", ICMS = ", mIcms)
    Call sqlValor(2, ", [Aliquota ICMS] = ", mAliquotaICMS)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pDataEntrada As Date, ByVal pCodigoProduto As Long, ByVal pNumeroDocumento As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Entrada_Produto"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Entrada] = " & preparaData(pDataEntrada)
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TotalQtdProdutoDatas(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalQtdProdutoDatas = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum(Quantidade) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM Entrada_Produto"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Codigo do Produto] = " & pCodigoProduto
    xSQL = xSQL & "    AND [Data da Entrada] >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND [Data da Entrada] <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND [Tipo da Entrada] <> " & 3
    
    Set rsEntradaProduto = New adodb.Recordset
    rsEntradaProduto.CursorLocation = adUseClient
    rsEntradaProduto.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaProduto.EOF Then
        If Not IsNull(rsEntradaProduto("QuantidadeVenda").Value) Then
            TotalQtdProdutoDatas = rsEntradaProduto("QuantidadeVenda").Value
        End If
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data da Entrada], [Codigo do Produto], [Numero do Documento], [Tipo da Entrada], [Preco de Custo], Quantidade, [Total do Custo], [Codigo do Fornecedor], [Data da Digitacao], Observacao, [Codigo do Usuario], [Custo Unitario Bruto], SubEstoque, CFOP, Modelo, Serie, Ordem, Base de Calculo ICMS, ICMS, [Aliquota ICMS] FROM Entrada_Produto" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEntradaProduto = New adodb.Recordset
    rsEntradaProduto.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEntradaProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEntradaProduto.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEntradaProduto("Empresa").Value
    DataEntrada = rsEntradaProduto("Data da Entrada").Value
    CodigoProduto = rsEntradaProduto("Codigo do Produto").Value
    NumeroDocumento = rsEntradaProduto("Numero do Documento").Value
    TipoEntrada = rsEntradaProduto("Tipo da Entrada").Value
    PrecoCusto = rsEntradaProduto("Preco de Custo").Value
    Quantidade = rsEntradaProduto("Quantidade").Value
    TotalCusto = rsEntradaProduto("Total do Custo").Value
    CodigoFornecedor = rsEntradaProduto("Codigo do Fornecedor").Value
    DataDigitacao = rsEntradaProduto("Data da Digitacao").Value
    Observacao = rsEntradaProduto("Observacao").Value
    CodigoUsuario = rsEntradaProduto("Codigo do Usuario").Value
    CustoUnitarioBruto = rsEntradaProduto("Custo Unitario Bruto").Value
    SubEstoque = rsEntradaProduto("SubEstoque").Value
    CFOP = rsEntradaProduto("CFOP").Value
    Modelo = rsEntradaProduto("Modelo").Value
    Serie = rsEntradaProduto("Serie").Value
    Ordem = rsEntradaProduto("Ordem").Value
    BaseCalculoICMS = rsEntradaProduto("Base de Calculo ICMS").Value
    Icms = rsEntradaProduto("ICMS").Value
    AliquotaICMS = rsEntradaProduto("Aliquota ICMS").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

