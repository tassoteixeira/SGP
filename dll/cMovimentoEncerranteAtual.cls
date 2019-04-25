VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoEncerranteAtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNomeTabela As String
Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mCodigoBomba As Integer
Private mEncerrante As Currency

Private rsMovBomba As New adodb.Recordset


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


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get CodigoBomba() As Integer
    CodigoBomba = mCodigoBomba
End Property
Public Property Let CodigoBomba(ByVal Valor As Integer)
    mCodigoBomba = Valor
End Property


Public Property Get Encerrante() As Currency
    Encerrante = mEncerrante
End Property
Public Property Let Encerrante(ByVal Valor As Currency)
    mEncerrante = Valor
End Property

'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nSubCaixa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & nCodigoBomba
    'xCondicao = xCondicao & " AND SubCaixa = " & nSubCaixa
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        AtribuiValor
        LocalizarCodigo = True
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "cMovimentoBomba:LocalizarCodigo"
    MsgBox gSQL 'ApenasTeste
End Function

Public Function LocalizarBicoAnteriorData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pCodigoBomba As Integer) As Boolean

Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarBicoAnteriorData = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & pCodigoBomba
    xCondicao = xCondicao & " AND Data <= " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, SubCaixa DESC")
    If Localizar(55) Then
        If rsMovBomba.RecordCount > 1 Then
            Do Until rsMovBomba.EOF
                If xRegIgual = False Then
                    If rsMovBomba("Data").Value = pData And rsMovBomba("Periodo").Value = pPeriodo Then
                        xRegIgual = True
                    Else
                        LocalizarBicoAnteriorData = True
                        Exit Do
                    End If
                Else
                    LocalizarBicoAnteriorData = True
                    Exit Do
                End If
                rsMovBomba.MoveNext
            Loop
        End If
    End If
    If LocalizarBicoAnteriorData = True Then
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function
trata_erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "cMovimentoBomba:LocalizarBicoAnteriorData"
    MsgBox gSQL 'ApenasTeste
End Function
Public Function LocalizarCombAnteriorData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCombAnteriorData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND Data < " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, SubCaixa DESC, [Codigo da Bomba] ASC")
    If Localizar(1) Then
        LocalizarCombAnteriorData = True
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "cMovimentoBomba:LocalizarCombAnteriorData"
    MsgBox gSQL 'ApenasTeste
End Function
Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, [Codigo da Bomba] DESC") 'SubCaixa DESC,
    If Localizar(55) Then
        If rsMovBomba.RecordCount > 1 Then
            Do Until rsMovBomba.EOF
                If xRegIgual = False Then
                    If rsMovBomba("Data").Value = mData And rsMovBomba("Periodo").Value = mPeriodo And rsMovBomba("Codigo da Bomba").Value = mCodigoBomba Then ' And rsMovBomba("SubCaixa").Value = mSubCaixa
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovBomba.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

'Localiza Primeira Data que o tiver período maior que 1
Public Function LocalizarPDPM1(ByVal nEmpresa As Integer) As Date

On Error GoTo trata_erro
    
    LocalizarPDPM1 = "00:00:00"
    gSQL = ""
    gSQL = gSQL & "SELECT Data"
    gSQL = gSQL & "  FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    'gSQL = gSQL & " AND Periodo > 1"
    'gSQL = gSQL & " AND Data > " & preparaData(CDate("31/10/2003"))
    gSQL = gSQL & " ORDER BY Data ASC"
    If Localizar(1) Then
        LocalizarPDPM1 = rsMovBomba("Data").Value
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Codigo da Bomba] ASC") 'SubCaixa ASC,
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
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
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Codigo da Bomba] ASC") 'SubCaixa ASC,
    If Localizar(55) Then
        If rsMovBomba.RecordCount > 1 Then
            Do Until rsMovBomba.EOF
                If xRegIgual = False Then
                    If rsMovBomba("Data").Value = mData And rsMovBomba("Periodo").Value = mPeriodo And rsMovBomba("Codigo da Bomba").Value = mCodigoBomba Then ' And rsMovBomba("SubCaixa").Value = mSubCaixa
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovBomba.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, [Codigo da Bomba] DESC") 'SubCaixa DESC,
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroPeriodoBico(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBomba As Integer, ByVal pPeriodo As Integer) As Boolean
Dim xCondicao As String
Dim xOrdem As String
On Error GoTo trata_erro

    LocalizarPrimeiroPeriodoBico = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & pCodigoBomba
    If pPeriodo > 0 Then
        xCondicao = xCondicao & " AND Periodo = " & pPeriodo
        xOrdem = "ORDER BY SubCaixa ASC"
    Else
        xOrdem = "ORDER BY Periodo ASC, SubCaixa ASC"
    End If
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(1) Then
        LocalizarPrimeiroPeriodoBico = True
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiroBicoComb(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarPrimeiroBicoComb = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xOrdem = "ORDER BY Periodo ASC, SubCaixa ASC, [Codigo da Bomba] ASC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(1) Then
        LocalizarPrimeiroBicoComb = True
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function





Public Function LocalizarUltimoPeriodoBico(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBomba As Integer, ByVal pPeriodo As Integer) As Boolean
Dim xCondicao As String
Dim xOrdem As String
On Error GoTo trata_erro

    LocalizarUltimoPeriodoBico = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & pCodigoBomba
    If pPeriodo > 0 Then
        xCondicao = xCondicao & " AND Periodo = " & pPeriodo
        xOrdem = "ORDER BY SubCaixa DESC"
    Else
        xOrdem = "ORDER BY Periodo DESC, SubCaixa DESC"
    End If
    gSQL = PreparaSQL(xCondicao, xOrdem)
    If Localizar(1) Then
        LocalizarUltimoPeriodoBico = True
        AtribuiValor
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO " & mNomeTabela & " ( Empresa, Data, Periodo, [Codigo da Bomba],  "
    gSQL = gSQL & "Encerrante ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mCodigoBomba, ", ")
    Call sqlValor(1, mEncerrante, ", ")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nSubCaixa As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Codigo da Bomba] = ", mCodigoBomba)
    Call sqlValor(2, ", Encerrante = ", mEncerrante)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nCodigoBomba
    gSQL = gSQL & " AND SubCaixa = " & nSubCaixa
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlteraParaSubCaixa(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNovoSubCaixa As Integer) As Boolean
Dim xRecordsAffected As Long

On Error GoTo TrataErro

    AlteraParaSubCaixa = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlValor(2, "SubCaixa = ", pNovoSubCaixa)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND SubCaixa = " & 999
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        AlteraParaSubCaixa = True
    End If
    Exit Function

TrataErro:
End Function


Public Function AlteraPrecoCusto(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pPrecoCusto As Currency) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_AlteraPrecoCusto

    AlteraPrecoCusto = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlValor(2, "[Preco de Custo] = ", pPrecoCusto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        AlteraPrecoCusto = True
    End If
    Exit Function

Err_AlteraPrecoCusto:
End Function


Public Function AlteraPrecoVendaBomba(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBomba As Integer, ByVal pPrecoVenda As Currency) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_AlteraPrecoVendaBico

    AlteraPrecoVendaBomba = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlValor(2, "[Preco de Venda] = ", pPrecoVenda)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        AlteraPrecoVendaBomba = True
    End If
    Exit Function

Err_AlteraPrecoVendaBico:
End Function


Public Function AlteraQuantidadeDatas(ByVal pEmpresa As Integer, ByVal pOperacao As String, ByVal pDataI As Date, ByVal pDataF As Date, ByVal pBicoBomba As Integer, ByVal pQuantidadeTotal As Currency) As Boolean
Dim lRecordsAffected As Long
Dim xQtdDias As Integer
Dim xQuantidade As Currency
Dim xQtdDiasAlterar As Integer
Dim xMensagem As String

On Error GoTo Err_AlteraQuantidadeDatas

    AlteraQuantidadeDatas = False
    xQtdDias = DateDiff("d", pDataI, pDataF) + 1
    xQuantidade = Format(pQuantidadeTotal / xQtdDias, "00000000.0")
    
    'Prepara Mensagem para Resumo da Operacao
    xQtdDiasAlterar = 0
    gSQL = "SELECT [Quantidade da Saida]"
    gSQL = gSQL & "  FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data >= " & preparaData(pDataI)
    gSQL = gSQL & "   AND Data <= " & preparaData(pDataF)
    gSQL = gSQL & "   AND [Codigo da Bomba] = " & pBicoBomba
    gSQL = gSQL & " ORDER BY Data ASC"
    If Localizar(0) Then
        If rsMovBomba.RecordCount > 1 Then
            Do Until rsMovBomba.EOF
                If pOperacao = "-" Then
                    If rsMovBomba("Quantidade da Saida").Value >= xQuantidade Then
                        xQtdDiasAlterar = xQtdDiasAlterar + 1
                    End If
                Else
                    xQtdDiasAlterar = xQtdDiasAlterar + 1
                End If
                rsMovBomba.MoveNext
            Loop
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    If xQtdDiasAlterar > 0 Then
        xMensagem = ""
        xMensagem = xMensagem & xQtdDias & " - Dias entre o intervalo de datas." & Chr(10)
        xMensagem = xMensagem & Format(pQuantidadeTotal, "###,###,##0.00") & " - Listros a ser "
        If pOperacao = "-" Then
            xMensagem = xMensagem & " DIMINUÍDO "
        Else
            xMensagem = xMensagem & " AUMENTADO "
        End If
        xMensagem = xMensagem & " nas vendas." & Chr(10)
        xMensagem = xMensagem & xQtdDiasAlterar & " - Dias que terao suas vendas alteradas." & Chr(10)
        xMensagem = xMensagem & Format(xQtdDiasAlterar * xQuantidade, "###,###,##0.00") & " - Quantidade que será alterada no período." & Chr(10)
        xMensagem = xMensagem & Chr(10)
        xMensagem = xMensagem & "Deseja realmente efetivar o processamento?"
        If (MsgBox(xMensagem, vbYesNo + vbQuestion + vbDefaultButton2, "Resumo para Processar.")) = vbNo Then
            Exit Function
        End If
    Else
        MsgBox "Não tem movimento para o critério à processar!", vbCritical, "Procedimento não será iniciado!"
        Exit Function
    End If
    
    
    
    gSQL = "UPDATE " & mNomeTabela & " SET "
    gSQL = gSQL & "[Quantidade da Saida] = [Quantidade da Saida]"
    If pOperacao = "+" Then
        gSQL = gSQL & " + "
    Else
        gSQL = gSQL & " - "
    End If
    gSQL = gSQL & preparaValor(xQuantidade)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataI)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataF)
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pBicoBomba
    If pOperacao = "-" Then
        gSQL = gSQL & " AND [Quantidade da Saida] > " & preparaValor(xQuantidade)
    End If
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlteraQuantidadeDatas = True
    End If
    Exit Function

Err_AlteraQuantidadeDatas:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nSubCaixa As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nCodigoBomba
    gSQL = gSQL & " AND SubCaixa = " & nSubCaixa
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluirData(ByVal nEmpresa As Integer, ByVal nData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirData = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirData = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Codigo da Bomba], Encerrante FROM " & mNomeTabela & " " & xCondicao)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovBomba.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovBomba("Empresa").Value
    Data = rsMovBomba("Data").Value
    Periodo = rsMovBomba("Periodo").Value
    CodigoBomba = rsMovBomba("Codigo da Bomba").Value
    Encerrante = rsMovBomba("Encerrante").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function AgrupaPeriodoData(ByVal nEmpresa As Integer, ByVal nData As Date) As Boolean
Dim lRecordsAffected As Long
Dim xSQL As String
Dim i As Integer
Dim xCodigoBomba(0 To 20) As Integer
Dim xAbertura(0 To 20) As Currency
Dim xEncerrante(0 To 20) As Currency
Dim xQuantidade(0 To 20) As Currency
Dim xPrecoCusto(0 To 20) As Currency
Dim xPrecoVenda(0 To 20) As Currency
Dim xTipoCombustivel(0 To 20) As String
Dim xNumeroTanque(0 To 20) As Integer
Dim xNumeroIlha(0 To 20) As Integer
Dim xSubCaixa(0 To 20) As Integer
Dim xContinuaProcessamento As Boolean

On Error GoTo trata_erro

    AgrupaPeriodoData = False
    xContinuaProcessamento = False
    For i = 0 To 20
        xCodigoBomba(i) = 0
        xAbertura(i) = 0
        xEncerrante(i) = 0
        xQuantidade(i) = 0
        xPrecoCusto(i) = 0
        xPrecoVenda(i) = 0
        xTipoCombustivel(i) = ""
        xNumeroTanque(i) = 0
        xNumeroIlha(i) = 0
        xSubCaixa(i) = 0
    Next
    
    'Acumula Dados
    xSQL = ""
    xSQL = xSQL & "   SELECT *"
    xSQL = xSQL & "     FROM " & mNomeTabela
    xSQL = xSQL & "    WHERE Empresa = " & nEmpresa
    xSQL = xSQL & "      AND Data = " & preparaData(nData)
    xSQL = xSQL & " ORDER BY Periodo, [Numero da Ilha], [Codigo da Bomba], [SubCaixa]"
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        Do Until rsMovBomba.EOF
            i = rsMovBomba("Codigo da Bomba").Value
            If xAbertura(i) = 0 Or xEncerrante(i) = 0 Then
                xCodigoBomba(i) = rsMovBomba("Codigo da Bomba").Value
                xAbertura(i) = rsMovBomba("Abertura").Value
                xEncerrante(i) = rsMovBomba("Encerrante").Value
                xQuantidade(i) = rsMovBomba("Quantidade da Saida").Value
                xPrecoCusto(i) = rsMovBomba("Preco de Custo").Value
                xPrecoVenda(i) = rsMovBomba("Preco de Venda").Value
                xTipoCombustivel(i) = rsMovBomba("Tipo de Combustivel").Value
                xNumeroTanque(i) = rsMovBomba("Numero do Tanque").Value
                xNumeroIlha(i) = rsMovBomba("Numero da Ilha").Value
                xSubCaixa(i) = rsMovBomba("SubCaixa").Value
            Else
                xEncerrante(i) = rsMovBomba("Encerrante").Value
                If (xQuantidade(i) + rsMovBomba("Quantidade da Saida").Value) > 0 Then
                    xPrecoCusto(i) = Format(((xPrecoCusto(i) * xQuantidade(i)) + (rsMovBomba("Preco de Custo").Value * rsMovBomba("Quantidade da Saida").Value)) / (xQuantidade(i) + rsMovBomba("Quantidade da Saida").Value), "00000000.0000")
                    xPrecoVenda(i) = Format(((xPrecoVenda(i) * xQuantidade(i)) + (rsMovBomba("Preco de Venda").Value * rsMovBomba("Quantidade da Saida").Value)) / (xQuantidade(i) + rsMovBomba("Quantidade da Saida").Value), "00000000.0000")
                Else
                    xPrecoCusto(i) = rsMovBomba("Preco de Custo").Value
                    xPrecoVenda(i) = rsMovBomba("Preco de Venda").Value
                End If
                xQuantidade(i) = xQuantidade(i) + rsMovBomba("Quantidade da Saida").Value
                xSubCaixa(i) = rsMovBomba("SubCaixa").Value
            End If
            If rsMovBomba("Periodo").Value <> 1 Then
                xContinuaProcessamento = True
            End If
            rsMovBomba.MoveNext
        Loop
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing

    'Testa se continua processamento
    If xContinuaProcessamento = False Then
        AgrupaPeriodoData = True
        Exit Function
    End If

    'Exclui Dados
    xSQL = ""
    xSQL = xSQL & "   DELETE"
    xSQL = xSQL & "     FROM " & mNomeTabela
    xSQL = xSQL & "    WHERE Empresa = " & nEmpresa
    xSQL = xSQL & "      AND Data = " & preparaData(nData)
    gConn.Execute xSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    

    'Grava Dados
    For i = 0 To 20
        If xAbertura(i) <> 0 Or xEncerrante(i) <> 0 Then
            Empresa = nEmpresa
            Data = nData
            Periodo = 1
            CodigoBomba = xCodigoBomba(i)
            'Abertura = xAbertura(i)
            Encerrante = xEncerrante(i)
            'QuantidadeSaida = xQuantidade(i)
            'PrecoCusto = xPrecoCusto(i)
            'PrecoVenda = xPrecoVenda(i)
            'TipoCombustivel = xTipoCombustivel(i)
            'NumeroTanque = xNumeroTanque(i)
            'NumeroIlha = xNumeroIlha(i)
            'SubCaixa = xSubCaixa(i)
            If Not Incluir Then
                Exit Function
            End If
        End If
    Next
    AgrupaPeriodoData = True
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Totaliza a venda em litros por bico entre data e período
Public Function TotalLitrosBicoPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoBico As Integer, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pDataPeriodo As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalLitrosBicoPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM([Quantidade da Saida]) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    xSQL = xSQL & "    AND [Codigo da Bomba] = " & pCodigoBico
    If pDataPeriodo <> "" Then
        xSQL = xSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("QuantidadeVenda").Value) Then
            TotalLitrosBicoPeriodo = rsMovBomba("QuantidadeVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Retorna a Abertura de Uma Bomba em uma Data e Período Específico
Public Function AberturaBicoDataPeriodo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBico As Integer, ByVal pPeriodo As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    AberturaBicoDataPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 1 Abertura "
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data = " & preparaData(pData)
    xSQL = xSQL & "    AND Periodo = " & pPeriodo
    xSQL = xSQL & "    AND [Codigo da Bomba] = " & pCodigoBico
    xSQL = xSQL & "  ORDER BY SubCaixa ASC"
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("Abertura").Value) Then
            AberturaBicoDataPeriodo = rsMovBomba("Abertura").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Retorna a Encerrante de Uma Bomba em uma Data e Período Específico
Public Function EncerranteBicoDataPeriodo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBico As Integer, ByVal pPeriodo As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    EncerranteBicoDataPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 1 Encerrante "
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data = " & preparaData(pData)
    xSQL = xSQL & "    AND Periodo = " & pPeriodo
    xSQL = xSQL & "    AND [Codigo da Bomba] = " & pCodigoBico
    xSQL = xSQL & "  ORDER BY SubCaixa DESC"
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("Encerrante").Value) Then
            EncerranteBicoDataPeriodo = rsMovBomba("Encerrante").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Retorna o ENCERRANTE de um determinado BICO
'Antes de um determinado Período e Data Específica
Public Function EncerranteBicoAnatesDataPeriodo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBico As Integer, ByVal pPeriodo As Integer, ByVal pSubCaixa As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    EncerranteBicoAnatesDataPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 20 [Codigo da Bomba], Data, Periodo, SubCaixa, Encerrante "
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Codigo da Bomba] = " & pCodigoBico
    xSQL = xSQL & "    AND Data <= " & preparaData(pData)
    xSQL = xSQL & "  ORDER BY Data DESC, Periodo DESC, SubCaixa DESC"
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    CriaLogCadastroDll ("EncerranteBicoAntesDataPeriodo SQL:" & gSQL)
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        Do Until rsMovBomba.EOF
            If rsMovBomba("Data").Value = pData Then
                If rsMovBomba("Periodo").Value < pPeriodo Then
                    EncerranteBicoAnatesDataPeriodo = rsMovBomba("Encerrante").Value
                    Exit Do
                ElseIf rsMovBomba("Periodo").Value = pPeriodo And rsMovBomba("SubCaixa").Value < pSubCaixa Then
                    EncerranteBicoAnatesDataPeriodo = rsMovBomba("Encerrante").Value
                    Exit Do
                End If
            Else
                If EncerranteBicoAnatesDataPeriodo = 0 Then
                    EncerranteBicoAnatesDataPeriodo = rsMovBomba("Encerrante").Value
                    Exit Do
                End If
            End If
            rsMovBomba.MoveNext
        Loop
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Totaliza a venda em Valor por bico entre data e período
Public Function TotalValorBicoPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoBico As Integer, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pDataPeriodo As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalValorBicoPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM(" & preparaArredonda("[Quantidade da Saida] * [Preco de Venda]", 2) & ") AS ValorVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    xSQL = xSQL & "    AND [Codigo da Bomba] = " & pCodigoBico
    If pDataPeriodo <> "" Then
        xSQL = xSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("ValorVenda").Value) Then
            TotalValorBicoPeriodo = rsMovBomba("ValorVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Totaliza Litros em uma data e tabela específica
Public Function TotalLitrosData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pNomeTabela As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalLitrosData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM([Quantidade da Saida]) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM " & pNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    If pTipoCombustivel <> "" Then
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("QuantidadeVenda").Value) Then
            TotalLitrosData = rsMovBomba("QuantidadeVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

'Comentei este dia 06/12
''Recalcula Encerrante de Um bico a partir de uma data
'Public Function RecalculaEncerrante(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoBomba As Integer) As Boolean
'Dim xSQL As String
'Dim rsBomba As New adodb.Recordset
'Dim xEncerante As Currency
'
'On Error GoTo trata_erro
'
'    RecalculaEncerrante = False
'
'    xSQL = ""
'    xSQL = xSQL & "SELECT Codigo"
'    xSQL = xSQL & "  FROM Bomba"
'    xSQL = xSQL & " WHERE Empresa = " & pEmpresa
'    If pCodigoBomba > 0 Then
'        xSQL = xSQL & " AND Codigo = " & pCodigoBomba
'    End If
'    xSQL = xSQL & " ORDER BY Codigo ASC"
'    Set rsBomba = New adodb.Recordset
'    rsBomba.CursorLocation = adUseClient
'    rsBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
'    If Not rsBomba.EOF Then
'        rsBomba.MoveFirst
'        Do Until rsBomba.EOF
'            xEncerante = 0
'            xSQL = ""
'            xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
'            xSQL = xSQL & "    AND Data >= " & preparaData(pData)
'            xSQL = xSQL & "    AND [Codigo da Bomba] = " & rsBomba("Codigo").Value
'            gSQL = PreparaSQL(xSQL, "ORDER BY Data, Periodo, SubCaixa")
'            Set rsMovBomba = New adodb.Recordset
'            rsMovBomba.CursorLocation = adUseClient
'            rsMovBomba.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
'
'            If Not rsMovBomba.EOF Then
'                rsMovBomba.MoveFirst
'                Do Until rsMovBomba.EOF
'
'                    If rsMovBomba("Data").Value = pData Then
'                        xEncerante = rsMovBomba("Encerrante").Value
'                    Else
'
'                        'If rsMovBomba("Data").Value = CDate("15/08/2004") Then
'                        '    MsgBox rsMovBomba("Encerrante").Value
'                        'End If
'
'                        AtribuiValor
'                        If xEncerante = 0 Then
'                            xEncerante = rsMovBomba("Abertura").Value
'                        End If
'                        Abertura = xEncerante
'                        Encerrante = Abertura + QuantidadeSaida
'                        'If g_empresa = 14 Or g_empresa = 16 Then
'                        '    If Encerrante > 9999.9 Then
'                        '        Encerrante = Encerrante - 10000
'                        '    End If
'                        'Else
'                        '    If Encerrante > 999999.9 Then
'                        '        Encerrante = Encerrante - 1000000
'                        '    End If
'                        'End If
'                        xEncerante = Encerrante
'                        If Abertura <> rsMovBomba("Abertura").Value Or Encerrante <> rsMovBomba("Encerrante").Value Then
'                            If Not Alterar(pEmpresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, rsMovBomba("Codigo da Bomba").Value, rsMovBomba("SubCaixa").Value) Then
'                                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade!"
'                                rsMovBomba.Close
'                                rsBomba.Close
'                                Set rsBomba = Nothing
'                                Set rsMovBomba = Nothing
'                                Exit Function
'                            End If
'                        End If
'                    End If
'                    rsMovBomba.MoveNext
'                Loop
'            End If
'            rsMovBomba.Close
'            rsBomba.MoveNext
'        Loop
'        RecalculaEncerrante = True
'    End If
'    rsBomba.Close
'    Set rsBomba = Nothing
'    Set rsMovBomba = Nothing
'    Exit Function
'
'trata_erro:
'    MsgBox Err.Number & " - " & Err.Description
'End Function


'Recalcula Encerrante de Um bico a partir de uma data
'Informando o Encerrante inicial
'Public Function RecalculaEncerrante2(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCodigoBomba As Integer, ByVal pEncerranteInicial As Currency) As Boolean
'Dim lRecordsAffected As Long
'Dim xSQL As String
'Dim rsMovBombaOld As New adodb.Recordset
'Dim xEncerrante As Currency
'
'On Error GoTo trata_erro
'
'    RecalculaEncerrante2 = False
'    xEncerrante = pEncerranteInicial
'
'    xSQL = ""
'    xSQL = xSQL & "SELECT Data, Periodo, SubCaixa, [Quantidade da Saida]"
'    xSQL = xSQL & "  FROM " & mNomeTabela
'    xSQL = xSQL & " WHERE Empresa = " & pEmpresa
'    xSQL = xSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
'    xSQL = xSQL & " AND Data >= " & preparaData(pDataInicial)
'    xSQL = xSQL & " AND Data <= " & preparaData(pDataFinal)
'    xSQL = xSQL & " ORDER BY Data ASC, Periodo ASC, SubCaixa ASC"
'    Set rsMovBombaOld = New adodb.Recordset
'    rsMovBombaOld.CursorLocation = adUseClient
'    rsMovBombaOld.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
'    If Not rsMovBombaOld.EOF Then
'        rsMovBombaOld.MoveFirst
'        Do Until rsMovBombaOld.EOF
'            gSQL = "UPDATE " & mNomeTabela & " SET "
'
'            Call sqlValor(2, "Abertura = ", xEncerrante)
'            xEncerrante = xEncerrante + rsMovBombaOld("Quantidade da Saida").Value
'            'If xEncerrante > 999999.9 Then
'            '    xEncerrante = xEncerrante - 1000000
'            'End If
'            Call sqlValor(2, ", Encerrante = ", xEncerrante)
'
'            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
'            gSQL = gSQL & " AND Data = " & preparaData(rsMovBombaOld("Data").Value)
'            gSQL = gSQL & " AND Periodo = " & rsMovBombaOld("Periodo").Value
'            gSQL = gSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
'            gSQL = gSQL & " AND SubCaixa = " & rsMovBombaOld("SubCaixa").Value
'            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
'            If lRecordsAffected = 0 Then
'                Exit Do
'            End If
'            rsMovBombaOld.MoveNext
'        Loop
'        If lRecordsAffected > 0 Then
'            RecalculaEncerrante2 = True
'        End If
'    End If
'    rsMovBombaOld.Close
'    Set rsMovBombaOld = Nothing
'    Exit Function
'
'trata_erro:
'    MsgBox Err.Number & " - " & Err.Description
'End Function


'Transfere dados Para o LMC/Escritorio
Public Function TransfereDados(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTabelaOrigem As String) As Boolean
Dim xSQL As String
Dim rsMovBombaOrigem As New adodb.Recordset

On Error GoTo trata_erro

    TransfereDados = False
    
    xSQL = " WHERE Empresa = " & pEmpresa
    xSQL = xSQL & " AND Data >= " & preparaData(pData)
    gSQL = PreparaSQL(xSQL, "ORDER BY Data ASC, Periodo ASC, [Codigo da Bomba] ASC, SubCaixa ASC")
    xSQL = Replace(gSQL, mNomeTabela, pTabelaOrigem)
    
    Set rsMovBombaOrigem = New adodb.Recordset
    rsMovBombaOrigem.CursorLocation = adUseClient
    rsMovBombaOrigem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBombaOrigem.EOF Then
        rsMovBombaOrigem.MoveFirst
        Do Until rsMovBombaOrigem.EOF
            Empresa = rsMovBombaOrigem("Empresa").Value
            Data = rsMovBombaOrigem("Data").Value
            Periodo = rsMovBombaOrigem("Periodo").Value
            CodigoBomba = rsMovBombaOrigem("Codigo da Bomba").Value
            'Abertura = rsMovBombaOrigem("Abertura").Value 'old
            Encerrante = rsMovBombaOrigem("Encerrante").Value
            'QuantidadeSaida = rsMovBombaOrigem("Quantidade da Saida").Value 'old
            'PrecoCusto = rsMovBombaOrigem("Preco de Custo").Value 'old
            'PrecoVenda = rsMovBombaOrigem("Preco de Venda").Value 'old
            'TipoCombustivel = rsMovBombaOrigem("Tipo de Combustivel").Value 'old
            'NumeroTanque = rsMovBombaOrigem("Numero do Tanque").Value 'old
            'NumeroIlha = rsMovBombaOrigem("Numero da Ilha").Value 'old
            'SubCaixa = rsMovBombaOrigem("SubCaixa").Value 'old
            If Not Incluir Then
                MsgBox "Não foi possível incluir registro!", vbInformation, "Erro de Integridade!"
                rsMovBombaOrigem.Close
                Set rsMovBombaOrigem = Nothing
                Exit Function
            End If
            rsMovBombaOrigem.MoveNext
        Loop
        TransfereDados = True
    End If
    rsMovBombaOrigem.Close
    Set rsMovBombaOrigem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Transfere dados Para o LMC
Public Function TransfereDadosLMC(ByVal pEmpresa As Integer, ByVal pSomenteVerifica As Boolean) As Boolean
    Dim xSQL As String
    Dim rsMovBombaOrigem As New adodb.Recordset
    Dim xData As Date

On Error GoTo trata_erro

    TransfereDadosLMC = False
    
    'Busca Ultima Data do Movimento LMC
    xData = UltimaData(pEmpresa)
    If xData = "00:00:00" Then
        Exit Function
    End If
    
    xData = xData + 1
    xSQL = " WHERE Empresa = " & pEmpresa
    xSQL = xSQL & " AND Data >= " & preparaData(xData)
    gSQL = PreparaSQL(xSQL, "ORDER BY Data ASC, Periodo ASC, [Codigo da Bomba] ASC, SubCaixa ASC")
    xSQL = Replace(gSQL, mNomeTabela, "Movimento_Bomba")
    
    Set rsMovBombaOrigem = New adodb.Recordset
    rsMovBombaOrigem.CursorLocation = adUseClient
    rsMovBombaOrigem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBombaOrigem.EOF Then
        If pSomenteVerifica Then
            TransfereDadosLMC = True
            Exit Function
        End If
        rsMovBombaOrigem.MoveFirst
        Do Until rsMovBombaOrigem.EOF
            Empresa = rsMovBombaOrigem("Empresa").Value
            Data = rsMovBombaOrigem("Data").Value
            Periodo = rsMovBombaOrigem("Periodo").Value
            CodigoBomba = rsMovBombaOrigem("Codigo da Bomba").Value
            'Abertura = rsMovBombaOrigem("Abertura").Value 'old
            Encerrante = rsMovBombaOrigem("Encerrante").Value
            'QuantidadeSaida = rsMovBombaOrigem("Quantidade da Saida").Value 'old
            'PrecoCusto = rsMovBombaOrigem("Preco de Custo").Value 'old
            'PrecoVenda = rsMovBombaOrigem("Preco de Venda").Value 'old
            'TipoCombustivel = rsMovBombaOrigem("Tipo de Combustivel").Value 'old
            'NumeroTanque = rsMovBombaOrigem("Numero do Tanque").Value 'old
            'NumeroIlha = rsMovBombaOrigem("Numero da Ilha").Value 'old
            'SubCaixa = rsMovBombaOrigem("SubCaixa").Value 'old
            If Not Incluir Then
                MsgBox "Não foi possível incluir registro do Movimento de Bomba!", vbInformation, "Erro de Integridade!"
                rsMovBombaOrigem.Close
                Set rsMovBombaOrigem = Nothing
                Exit Function
            End If
            rsMovBombaOrigem.MoveNext
        Loop
        TransfereDadosLMC = True
    End If
    rsMovBombaOrigem.Close
    Set rsMovBombaOrigem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function UltimaData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro

    UltimaData = "01/01/1900"
    gSQL = "SELECT Data FROM Movimento_Bomba_LMC WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Data DESC, Periodo DESC, SubCaixa DESC, [Codigo da Bomba] DESC"
    If Localizar(1) Then
        UltimaData = rsMovBomba("Data").Value
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function QuantidadeVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pTanque As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    QuantidadeVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM([Quantidade da Saida]) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    If pTipoCombustivel <> "" Then
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    If pTanque > 0 Then
        xSQL = xSQL & "    AND [Numero do Tanque] = " & Str(pTanque)
    End If
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("QuantidadeVenda").Value) Then
            QuantidadeVendaData = rsMovBomba("QuantidadeVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalVendaPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalVendaPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM([Quantidade da Saida]) AS QuantidadeVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    If pTipoCombustivel <> "" Then
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("QuantidadeVenda").Value) Then
            TotalVendaPeriodo = rsMovBomba("QuantidadeVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorVendaPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorVendaPeriodo = 0
    xSQL = ""
    'xSQL = xSQL & " SELECT SUM(Format([Quantidade da Saida] * [Preco de Venda], " & Chr(39) & "00000000.00" & Chr(39) & ")) AS ValorVenda"
    xSQL = xSQL & " SELECT SUM([Quantidade da Saida] * [Preco de Venda]) AS ValorVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    If pTipoCombustivel <> "" Then
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("ValorVenda").Value) Then
            ValorVendaPeriodo = Format(rsMovBomba("ValorVenda").Value, "00000000.00")
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorCustoVendaPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pDataPeriodo As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    ValorCustoVendaPeriodo = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM(" & preparaArredonda("[Quantidade da Saida] * [Preco de Custo]", 4) & ") AS TotalCustoVenda"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If pDataPeriodo <> "" Then
        xSQL = xSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    
    Set rsMovBomba = New adodb.Recordset
    rsMovBomba.CursorLocation = adUseClient
    rsMovBomba.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBomba.EOF Then
        If Not IsNull(rsMovBomba("TotalCustoVenda").Value) Then
            ValorCustoVendaPeriodo = rsMovBomba("TotalCustoVenda").Value
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Verifica se existe movimento entre data e período
Public Function ExisteMovimentoPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteMovimentoPeriodo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(pDataInicial)
    xCondicao = xCondicao & " AND Data <= " & preparaData(pDataFinal)
    xCondicao = xCondicao & " AND Periodo >= " & pPeriodoInicial
    xCondicao = xCondicao & " AND Periodo <= " & pPeriodoFinal
    gSQL = PreparaSQL(xCondicao, "")
    ExisteMovimentoPeriodo = Localizar(1)
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Verifica se existe movimento Em Uma Data Periodo e SubCaixa
Public Function ExisteDataPeriodoSubCx(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pSubCaixa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteDataPeriodoSubCx = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND SubCaixa = " & pSubCaixa
    gSQL = PreparaSQL(xCondicao, "")
    ExisteDataPeriodoSubCx = Localizar(1)
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Pega número do próximo SubCaixa de uma data e período
Public Function ProximoSubCaixa(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer) As Integer
Dim xSQL As String
On Error GoTo trata_erro

    ProximoSubCaixa = 1
    gSQL = ""
    gSQL = gSQL & "SELECT SubCaixa"
    gSQL = gSQL & "  FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Periodo = " & pPeriodo
    gSQL = gSQL & "   AND SubCaixa < " & 999
    gSQL = gSQL & " ORDER BY SubCaixa DESC"
    If Localizar(1) Then
        ProximoSubCaixa = rsMovBomba("SubCaixa").Value + 1
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Pega o último bico com movimento entre data e período
Public Function UltimoBicoComMovimento(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Integer
Dim xSQL As String
On Error GoTo trata_erro

    UltimoBicoComMovimento = 0
    gSQL = ""
    gSQL = gSQL & "SELECT [Codigo da Bomba]"
    gSQL = gSQL & "  FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & "   AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & "   AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & "   AND Periodo <= " & pPeriodoFinal
    gSQL = gSQL & " ORDER BY [Codigo da Bomba] DESC"
    If Localizar(1) Then
        UltimoBicoComMovimento = rsMovBomba("Codigo da Bomba").Value
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


