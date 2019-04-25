VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoMapaResumoCorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mNumero As Long
Private mECFNumero As Integer
Private mContagemOperacaoInicial As Long
Private mContagemOperacaoFinal As Long
Private mTotalizadorGeralFinal As Currency
Private mTotalizadorGeralInicial As Currency
Private mCancelamentoItem As Currency
Private mValorContabil As Currency
Private mIsentasNaoTributadas As Currency
Private mSubstituicaoTributaria As Currency
Private mICMS17 As Currency
Private mContadorReducoesZ As Long
Private mObservacao1 As String
Private mObservacao2 As String
Private mValorCombustivel As Currency
Private mValorProduto As Currency

Private rsMapaResumoCorrecao As New adodb.Recordset


'Inicio das Propriedades da Classe
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


Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
End Property


Public Property Get ECFNumero() As Integer
    ECFNumero = mECFNumero
End Property
Public Property Let ECFNumero(ByVal Valor As Integer)
    mECFNumero = Valor
End Property


Public Property Get ContagemOperacaoInicial() As Long
    ContagemOperacaoInicial = mContagemOperacaoInicial
End Property
Public Property Let ContagemOperacaoInicial(ByVal Valor As Long)
    mContagemOperacaoInicial = Valor
End Property


Public Property Get ContagemOperacaoFinal() As Long
    ContagemOperacaoFinal = mContagemOperacaoFinal
End Property
Public Property Let ContagemOperacaoFinal(ByVal Valor As Long)
    mContagemOperacaoFinal = Valor
End Property


Public Property Get TotalizadorGeralFinal() As Currency
    TotalizadorGeralFinal = mTotalizadorGeralFinal
End Property
Public Property Let TotalizadorGeralFinal(ByVal Valor As Currency)
    mTotalizadorGeralFinal = Valor
End Property


Public Property Get TotalizadorGeralInicial() As Currency
    TotalizadorGeralInicial = mTotalizadorGeralInicial
End Property
Public Property Let TotalizadorGeralInicial(ByVal Valor As Currency)
    mTotalizadorGeralInicial = Valor
End Property


Public Property Get CancelamentoItem() As Currency
    CancelamentoItem = mCancelamentoItem
End Property
Public Property Let CancelamentoItem(ByVal Valor As Currency)
    mCancelamentoItem = Valor
End Property


Public Property Get ValorContabil() As Currency
    ValorContabil = mValorContabil
End Property
Public Property Let ValorContabil(ByVal Valor As Currency)
    mValorContabil = Valor
End Property


Public Property Get IsentasNaoTributadas() As Currency
    IsentasNaoTributadas = mIsentasNaoTributadas
End Property
Public Property Let IsentasNaoTributadas(ByVal Valor As Currency)
    mIsentasNaoTributadas = Valor
End Property


Public Property Get SubstituicaoTributaria() As Currency
    SubstituicaoTributaria = mSubstituicaoTributaria
End Property
Public Property Let SubstituicaoTributaria(ByVal Valor As Currency)
    mSubstituicaoTributaria = Valor
End Property


Public Property Get ICMS17() As Currency
    ICMS17 = mICMS17
End Property
Public Property Let ICMS17(ByVal Valor As Currency)
    mICMS17 = Valor
End Property


Public Property Get ContadorReducoesZ() As Long
    ContadorReducoesZ = mContadorReducoesZ
End Property
Public Property Let ContadorReducoesZ(ByVal Valor As Long)
    mContadorReducoesZ = Valor
End Property


Public Property Get Observacao1() As String
    Observacao1 = mObservacao1
End Property
Public Property Let Observacao1(ByVal Valor As String)
    mObservacao1 = Valor
End Property


Public Property Get Observacao2() As String
    Observacao2 = mObservacao2
End Property
Public Property Let Observacao2(ByVal Valor As String)
    mObservacao2 = Valor
End Property


Public Property Get ValorCombustivel() As Currency
    ValorCombustivel = mValorCombustivel
End Property
Public Property Let ValorCombustivel(ByVal Valor As Currency)
    mValorCombustivel = Valor
End Property


Public Property Get ValorProduto() As Currency
    ValorProduto = mValorProduto
End Property
Public Property Let ValorProduto(ByVal Valor As Currency)
    mValorProduto = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pECFNumero As Integer, ByVal pData As Date, ByVal pNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [ECF Numero] = " & pECFNumero
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Numero = " & pNumero
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarData(ByVal pEmpresa As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarData = Localizar(1)
    If LocalizarData Then
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarDataECF(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal pECFNumero As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarDataECF = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [ECF Numero] = " & pECFNumero
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarDataECF = True
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [ECF Numero] DESC, Numero DESC")
    If Localizar(10) Then
        If rsMapaResumoCorrecao.RecordCount > 1 Then
            Do Until rsMapaResumoCorrecao.EOF
                If xRegIgual = False Then
                    If rsMapaResumoCorrecao("Data").Value = mData And rsMapaResumoCorrecao("ECF Numero").Value = mECFNumero And rsMapaResumoCorrecao("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    AtribuiValor
                    Exit Do
                End If
                rsMapaResumoCorrecao.MoveNext
            Loop
        End If
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [ECF Numero] ASC, Numero ASC")
    If Localizar(1) Then
        LocalizarPrimeiro = True
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroData(ByVal pEmpresa As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroData = False
    
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [ECF Numero] ASC, Numero ASC")
    If Localizar(1) Then
        LocalizarPrimeiroData = True
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [ECF Numero] ASC, Numero ASC")
    If Localizar(10) Then
        If rsMapaResumoCorrecao.RecordCount > 1 Then
            Do Until rsMapaResumoCorrecao.EOF
                If xRegIgual = False Then
                    If rsMapaResumoCorrecao("Data").Value = mData And rsMapaResumoCorrecao("ECF Numero").Value = mECFNumero And rsMapaResumoCorrecao("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    AtribuiValor
                    Exit Do
                End If
                rsMapaResumoCorrecao.MoveNext
            Loop
        End If
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [ECF Numero] DESC, Numero DESC")
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MapaResumoCorrecao ( Empresa, Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], "
    gSQL = gSQL & "[Contagem de Operacao Final], [Totalizador Geral Final], [Totalizador Geral Inicial], [Cancelamento de Item], [Valor Contabil], "
    gSQL = gSQL & "[Isentas Nao Tributadas], [Substituicao Tributaria], [ICMS 17], [Contador de Reducoes Z], [Observacao 1], "
    gSQL = gSQL & "[Observacao 2], [Valor Combustivel], [Valor Produto] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlNumero(1, mECFNumero, ", ")
    Call sqlNumero(1, mContagemOperacaoInicial, ", ")
    Call sqlNumero(1, mContagemOperacaoFinal, ", ")
    Call sqlValor(1, mTotalizadorGeralFinal, ", ")
    Call sqlValor(1, mTotalizadorGeralInicial, ", ")
    Call sqlValor(1, mCancelamentoItem, ", ")
    Call sqlValor(1, mValorContabil, ", ")
    Call sqlValor(1, mIsentasNaoTributadas, ", ")
    Call sqlValor(1, mSubstituicaoTributaria, ", ")
    Call sqlValor(1, mICMS17, ", ")
    Call sqlNumero(1, mContadorReducoesZ, ", ")
    Call sqlTexto(1, mObservacao1, ", ")
    Call sqlTexto(1, mObservacao2, ", ")
    Call sqlValor(1, mValorCombustivel, ", ")
    Call sqlValor(1, mValorProduto, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pECFNumero As Integer, ByVal pData As Date, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MapaResumoCorrecao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Numero = ", mNumero)
    Call sqlNumero(2, ", [ECF Numero] = ", mECFNumero)
    Call sqlNumero(2, ", [Contagem de Operacao Inicial] = ", mContagemOperacaoInicial)
    Call sqlNumero(2, ", [Contagem de Operacao Final] = ", mContagemOperacaoFinal)
    Call sqlValor(2, ", [Totalizador Geral Final] = ", mTotalizadorGeralFinal)
    Call sqlValor(2, ", [Totalizador Geral Inicial] = ", mTotalizadorGeralInicial)
    Call sqlValor(2, ", [Cancelamento de Item] = ", mCancelamentoItem)
    Call sqlValor(2, ", [Valor Contabil] = ", mValorContabil)
    Call sqlValor(2, ", [Isentas Nao Tributadas] = ", mIsentasNaoTributadas)
    Call sqlValor(2, ", [Substituicao Tributaria] = ", mSubstituicaoTributaria)
    Call sqlValor(2, ", [ICMS 17] = ", mICMS17)
    Call sqlNumero(2, ", [Contador de Reducoes Z] = ", mContadorReducoesZ)
    Call sqlTexto(2, ", [Observacao 1] = ", mObservacao1)
    Call sqlTexto(2, ", [Observacao 2] = ", mObservacao2)
    Call sqlValor(2, ", [Valor Combustivel] = ", mValorCombustivel)
    Call sqlValor(2, ", [Valor Produto] = ", mValorProduto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [ECF Numero] = " & pECFNumero
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pECFNumero As Integer, ByVal pData As Date, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MapaResumoCorrecao"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [ECF Numero] = " & pECFNumero
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function QuantidadeMapaResumoData(ByVal pEmpresa As Integer, ByVal pData As Date) As Integer
On Error GoTo trata_erro

    QuantidadeMapaResumoData = 0
    gSQL = "SELECT COUNT(1) AS Quantidade"
    gSQL = gSQL & "  FROM MapaResumoCorrecao"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    If Localizar(1) Then
        If Not IsNull(rsMapaResumoCorrecao("Quantidade").Value) Then
            QuantidadeMapaResumoData = rsMapaResumoCorrecao("Quantidade").Value
        End If
    End If
    rsMapaResumoCorrecao.Close
    Set rsMapaResumoCorrecao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], [Contagem de Operacao Final], [Totalizador Geral Final], [Totalizador Geral Inicial], [Cancelamento de Item], [Valor Contabil], [Isentas Nao Tributadas], [Substituicao Tributaria], [ICMS 17], [Contador de Reducoes Z], [Observacao 1], [Observacao 2], [Valor Combustivel], [Valor Produto] FROM MapaResumoCorrecao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMapaResumoCorrecao = New adodb.Recordset
    rsMapaResumoCorrecao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMapaResumoCorrecao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMapaResumoCorrecao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMapaResumoCorrecao("Empresa").Value
    Data = rsMapaResumoCorrecao("Data").Value
    Numero = rsMapaResumoCorrecao("Numero").Value
    ECFNumero = rsMapaResumoCorrecao("ECF Numero").Value
    ContagemOperacaoInicial = rsMapaResumoCorrecao("Contagem de Operacao Inicial").Value
    ContagemOperacaoFinal = rsMapaResumoCorrecao("Contagem de Operacao Final").Value
    TotalizadorGeralFinal = rsMapaResumoCorrecao("Totalizador Geral Final").Value
    TotalizadorGeralInicial = rsMapaResumoCorrecao("Totalizador Geral Inicial").Value
    CancelamentoItem = rsMapaResumoCorrecao("Cancelamento de Item").Value
    ValorContabil = rsMapaResumoCorrecao("Valor Contabil").Value
    IsentasNaoTributadas = rsMapaResumoCorrecao("Isentas Nao Tributadas").Value
    SubstituicaoTributaria = rsMapaResumoCorrecao("Substituicao Tributaria").Value
    ICMS17 = rsMapaResumoCorrecao("ICMS 17").Value
    ContadorReducoesZ = rsMapaResumoCorrecao("Contador de Reducoes Z").Value
    Observacao1 = rsMapaResumoCorrecao("Observacao 1").Value
    Observacao2 = rsMapaResumoCorrecao("Observacao 2").Value
    ValorCombustivel = rsMapaResumoCorrecao("Valor Combustivel").Value
    ValorProduto = rsMapaResumoCorrecao("Valor Produto").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

