VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoMapaResumo"
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
Private mIsentas As Currency
Private mSubstituicaoTributaria As Currency
Private mICMS17 As Currency
Private mContadorReducoesZ As Long
Private mObservacao1 As String
Private mObservacao2 As String
Private mICMS12 As Currency
Private mContagemReinicioOperacao As Long
Private mDesconto As Currency
Private mAcrescimo As Currency
Private mNaoIncidencia As Currency
Private mICMS3 As Currency 'propiedade do campo ICMS 7%
Private mICMS7 As Currency 'propiedade do campo ICMS 7%
Private mICMS25 As Currency 'propiedade do campo ICMS 25%
Private mICMS13 As Currency
Private mICMS19 As Currency

Private rsMapaResumo As New adodb.Recordset


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


Public Property Get Isentas() As Currency
    Isentas = mIsentas
End Property
Public Property Let Isentas(ByVal Valor As Currency)
    mIsentas = Valor
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


Public Property Get ICMS12() As Currency
    ICMS12 = mICMS12
End Property
Public Property Let ICMS12(ByVal Valor As Currency)
    mICMS12 = Valor
End Property


Public Property Get ContagemReinicioOperacao() As Long
    ContagemReinicioOperacao = mContagemReinicioOperacao
End Property
Public Property Let ContagemReinicioOperacao(ByVal Valor As Long)
    mContagemReinicioOperacao = Valor
End Property


Public Property Get Desconto() As Currency
    Desconto = mDesconto
End Property
Public Property Let Desconto(ByVal Valor As Currency)
    mDesconto = Valor
End Property


Public Property Get Acrescimo() As Currency
    Acrescimo = mAcrescimo
End Property
Public Property Let Acrescimo(ByVal Valor As Currency)
    mAcrescimo = Valor
End Property


Public Property Get NaoIncidencia() As Currency
    NaoIncidencia = mNaoIncidencia
End Property
Public Property Let NaoIncidencia(ByVal Valor As Currency)
    mNaoIncidencia = Valor
End Property

Public Property Get ICMS3() As Currency
    ICMS3 = mICMS3
End Property
Public Property Let ICMS3(ByVal Valor As Currency)
    mICMS3 = Valor
End Property

Public Property Get ICMS7() As Currency
    ICMS7 = mICMS7
End Property
Public Property Let ICMS7(ByVal Valor As Currency)
    mICMS7 = Valor
End Property

Public Property Get ICMS25() As Currency
    ICMS25 = mICMS25
End Property
Public Property Let ICMS25(ByVal Valor As Currency)
    mICMS25 = Valor
End Property

Public Property Get ICMS13() As Currency
    ICMS13 = mICMS13
End Property
Public Property Let ICMS13(ByVal Valor As Currency)
    mICMS13 = Valor
End Property

Public Property Get ICMS19() As Currency
    ICMS19 = mICMS19
End Property
Public Property Let ICMS19(ByVal Valor As Currency)
    mICMS19 = Valor
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
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
        If rsMapaResumo.RecordCount > 1 Then
            Do Until rsMapaResumo.EOF
                If xRegIgual = False Then
                    If rsMapaResumo("Data").Value = mData And rsMapaResumo("ECF Numero").Value = mECFNumero And rsMapaResumo("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    AtribuiValor
                    Exit Do
                End If
                rsMapaResumo.MoveNext
            Loop
        End If
    End If
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnteriorDataECF(ByVal pEmpresa As Integer, ByVal pECFNumero As Integer, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnteriorDataECF = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [ECF Numero] = " & pECFNumero
    xCondicao = xCondicao & " AND Data < " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Numero DESC")
    If Localizar(1) Then
        LocalizarAnteriorDataECF = True
        AtribuiValor
    End If
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
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
        If rsMapaResumo.RecordCount > 1 Then
            Do Until rsMapaResumo.EOF
                If xRegIgual = False Then
                    If rsMapaResumo("Data").Value = mData And rsMapaResumo("ECF Numero").Value = mECFNumero And rsMapaResumo("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    AtribuiValor
                    Exit Do
                End If
                rsMapaResumo.MoveNext
            Loop
        End If
    End If
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [ECF Numero] DESC, Numero DESC")
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Mapa_Resumo ( Empresa, Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], "
    gSQL = gSQL & "[Contagem de Operacao Final], [Totalizador Geral Final], [Totalizador Geral Inicial], [Cancelamento de Item], [Valor Contabil], "
    gSQL = gSQL & "[Isentas], [Substituicao Tributaria], [ICMS 17], [Contador de Reducoes Z], [Observacao 1], "
    gSQL = gSQL & "[Observacao 2], [ICMS 12], [Contagem de Reinicio de Operacao], Desconto, Acrescimo, "
    gSQL = gSQL & "[Nao Incidencia], [ICMS 3], [ICMS 7], [ICMS 25], [ICMS 13], [ICMS 19] ) VALUES ( "
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
    Call sqlValor(1, mIsentas, ", ")
    Call sqlValor(1, mSubstituicaoTributaria, ", ")
    Call sqlValor(1, mICMS17, ", ")
    Call sqlNumero(1, mContadorReducoesZ, ", ")
    Call sqlTexto(1, mObservacao1, ", ")
    Call sqlTexto(1, mObservacao2, ", ")
    Call sqlValor(1, mICMS12, ", ")
    Call sqlNumero(1, mContagemReinicioOperacao, ", ")
    Call sqlValor(1, mDesconto, ", ")
    Call sqlValor(1, mAcrescimo, ", ")
    Call sqlValor(1, mNaoIncidencia, ", ")
    Call sqlValor(1, mICMS3, ", ")
    Call sqlValor(1, mICMS7, ", ") 'atribui a propiedade ICMS7 no metodo
    Call sqlValor(1, mICMS25, ", ") 'atribui a propiedade ICMS25 no metodo
    Call sqlValor(1, mICMS13, ", ") 'atribui a propiedade ICMS13 no metodo
    Call sqlValor(1, mICMS19, " )")
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
    gSQL = "UPDATE Mapa_Resumo SET "
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
    Call sqlValor(2, ", [Isentas] = ", mIsentas)
    Call sqlValor(2, ", [Substituicao Tributaria] = ", mSubstituicaoTributaria)
    Call sqlValor(2, ", [ICMS 17] = ", mICMS17)
    Call sqlNumero(2, ", [Contador de Reducoes Z] = ", mContadorReducoesZ)
    Call sqlTexto(2, ", [Observacao 1] = ", mObservacao1)
    Call sqlTexto(2, ", [Observacao 2] = ", mObservacao2)
    Call sqlValor(2, ", [ICMS 12] = ", mICMS12)
    Call sqlNumero(2, ", [Contagem de Reinicio de Operacao] = ", mContagemReinicioOperacao)
    Call sqlValor(2, ", Desconto = ", mDesconto)
    Call sqlValor(2, ", Acrescimo = ", mAcrescimo)
    Call sqlValor(2, ", [Nao Incidencia] = ", mNaoIncidencia)
    Call sqlValor(2, ", [ICMS 3] = ", mICMS3)
    Call sqlValor(2, ", [ICMS 7] = ", mICMS7) 'atribui a propiedade mICMS7 ao metodo
    Call sqlValor(2, ", [ICMS 25] = ", mICMS25) 'atribui a propiedade mICMS25 ao metodo
    Call sqlValor(2, ", [ICMS 13] = ", mICMS13) 'atribui a propiedade mICMS13 ao metodo
    Call sqlValor(2, ", [ICMS 19] = ", mICMS19)
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
    gSQL = "DELETE FROM Mapa_Resumo"
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
    gSQL = gSQL & "  FROM Mapa_Resumo"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    If Localizar(1) Then
        If Not IsNull(rsMapaResumo("Quantidade").Value) Then
            QuantidadeMapaResumoData = rsMapaResumo("Quantidade").Value
        End If
    End If
    rsMapaResumo.Close
    Set rsMapaResumo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], [Contagem de Operacao Final], [Totalizador Geral Final], [Totalizador Geral Inicial], [Cancelamento de Item], [Valor Contabil], [Isentas], [Substituicao Tributaria], [ICMS 17], [Contador de Reducoes Z], [Observacao 1], [Observacao 2], [ICMS 12], [Contagem de Reinicio de Operacao], Desconto, Acrescimo, [Nao Incidencia], [ICMS 3], [ICMS 7], [ICMS 25], [ICMS 13], [ICMS 19] FROM Mapa_Resumo" & " " & xCondicao & " " & xTipoOrdem)
                                                                                                                                                                                                                                                                                                                                                                                                                                    'busca este campo na tabela
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMapaResumo = New adodb.Recordset
    rsMapaResumo.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMapaResumo.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMapaResumo.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMapaResumo("Empresa").Value
    Data = rsMapaResumo("Data").Value
    Numero = rsMapaResumo("Numero").Value
    ECFNumero = rsMapaResumo("ECF Numero").Value
    ContagemOperacaoInicial = rsMapaResumo("Contagem de Operacao Inicial").Value
    ContagemOperacaoFinal = rsMapaResumo("Contagem de Operacao Final").Value
    TotalizadorGeralFinal = rsMapaResumo("Totalizador Geral Final").Value
    TotalizadorGeralInicial = rsMapaResumo("Totalizador Geral Inicial").Value
    CancelamentoItem = rsMapaResumo("Cancelamento de Item").Value
    ValorContabil = rsMapaResumo("Valor Contabil").Value
    Isentas = rsMapaResumo("Isentas").Value
    SubstituicaoTributaria = rsMapaResumo("Substituicao Tributaria").Value
    ICMS17 = rsMapaResumo("ICMS 17").Value
    ContadorReducoesZ = rsMapaResumo("Contador de Reducoes Z").Value
    Observacao1 = rsMapaResumo("Observacao 1").Value
    Observacao2 = rsMapaResumo("Observacao 2").Value
    ICMS12 = rsMapaResumo("ICMS 12").Value
    ContagemReinicioOperacao = rsMapaResumo("Contagem de Reinicio de Operacao").Value
    Desconto = rsMapaResumo("Desconto").Value
    Acrescimo = rsMapaResumo("Acrescimo").Value
    NaoIncidencia = rsMapaResumo("Nao Incidencia").Value
    ICMS3 = rsMapaResumo("ICMS 3").Value
    ICMS7 = rsMapaResumo("ICMS 7").Value
    ICMS25 = rsMapaResumo("ICMS 25").Value
    ICMS13 = rsMapaResumo("ICMS 13").Value
    ICMS19 = rsMapaResumo("ICMS 19").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

