Attribute VB_Name = "ecf_schalter"
Declare Function ecfLeituraX Lib "dll32phi.dll" (ByVal operador As String) As Integer
Declare Function ecfReducaoZ Lib "dll32phi.dll" (ByVal operador As String) As Integer
Declare Function ecfCancVenda Lib "dll32phi.dll" (ByVal operador As String) As Integer
Declare Function ecfFimTrans Lib "dll32phi.dll" (ByVal operador As String) As Integer
Declare Function ecfCancDoc Lib "dll32phi.dll" (ByVal operador As String) As Integer
Declare Function ecfImpLinha Lib "dll32phi.dll" (ByVal szLinha As String) As Integer
Declare Function ecfInicCupomNFiscal Lib "dll32phi.dll" (ByVal byTipo As Integer) As Integer
Declare Function ecfImpCab Lib "dll32phi.dll" (ByVal byTip As Integer) As Integer
Declare Function ecfLineFeed Lib "dll32phi.dll" (ByVal byEst As Integer, ByVal wLin As Integer) As Integer
Declare Function ecfVendaItem3d Lib "dll32phi.dll" (ByVal szCodigo As String, ByVal szDescricao As String, ByVal szQuantidade As String, ByVal szValor As String, ByVal byTaxa As Integer, ByVal szUnidade As String, ByVal szDigito As String) As Integer
Declare Function ecfVendaItem Lib "dll32phi.dll" (ByVal szDescr As String, ByVal szValor As String, ByVal byTaxa As Integer) As Integer
Declare Function ecfVendaItem78 Lib "dll32phi.dll" (ByVal szDescr As String, ByVal szValor As String, ByVal byTaxa As Integer) As Integer
Declare Function ecfStatusVincs Lib "dll32phi.dll" (position As Integer) As String
Declare Function ecfStatusCupom Lib "dll32phi.dll" (ByVal flag_geral As Integer) As String
Declare Function ecfParamStatusImp Lib "dll32phi.dll" (ByVal num1 As String, ByVal num2 As String, ByVal num3 As String) As Integer
Declare Function ecfParamStatusCup Lib "dll32phi.dll" (ByVal pwPDV As String, ByVal pwTipoDoc As String, ByVal pdwCupom As String, ByVal szData As String, ByVal szHora As String, ByVal szSubTotal As String, ByVal szGrandeTotal As String) As Integer
Declare Function ecfAutentica Lib "dll32phi.dll" (ByVal szLinha As String) As Integer
Declare Function ecfPayPatterns Lib "dll32phi.dll" (ByVal szPosTab As String, ByVal szTitulo As String) As Integer
Declare Function ecfPagamento Lib "dll32phi.dll" (ByVal byTipo As Integer, ByVal szPosTable As String, ByVal szValor As String, ByVal byLmens As Integer) As Integer
Declare Function ecfFimTransVinc Lib "dll32phi.dll" (ByVal szOperTerm As String, ByVal szVinculados As String) As Integer
Declare Function ecfLeitMemFisc Lib "dll32phi.dll" (ByVal byTipo As Integer, ByVal szDi As String, ByVal szDf As String, ByVal wRi As Integer, ByVal wRf As Integer, ByVal archive As String) As Integer
Declare Function ecfAcertaData Lib "dll32phi.dll" (ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer, ByVal hor As Integer, ByVal min As Integer, ByVal seg As Integer) As Integer
Function SchalterImprimeCabecalho(x_formato As Integer) As Integer
    'formato = 128
    'formato = 20
    'formato = 0
    SchalterImprimeCabecalho = ecfImpCab(x_formato)
End Function
Function SchalterCancelaCupom(x_operador As String) As Integer
    'nome_do_operador = "caixa_3"
    SchalterCancelaCupom = ecfCancDoc(x_operador)
End Function
Function SchalterFinalizaCupom(x_operador As String) As Integer
    'nome_do_operador = "caixa_3"
    SchalterFinalizaCupom = ecfFimTrans(x_operador)
End Function
Function SchalterParamStatusImp() As Integer
    Dim pbyUsers As String
    Dim pbyStat1 As String
    Dim pbyStat2 As String
    Dim code1 As Integer
    Dim code2 As Integer
    Dim code3 As Integer
    Dim teste As Integer
    Dim nomeErro As String
    pbyUsers = "t"
    pbyStat1 = "t"
    pbyStat2 = "t"
    SchalterParamStatusImp = ecfParamStatusImp(pbyUsers, pbyStat1, pbyStat2)
    If (SchalterParamStatusImp <> 0) Then
        MsgBox "Erro de número: " & SchalterParamStatusImp & Chr(10) & "Verifique se a impressora está ligada.", vbInformation, "Erro de Comunicação! - Schalter"
        Exit Function
    End If
    code1 = Asc(pbyUsers)
    code2 = Asc(pbyStat1)
    code3 = Asc(pbyStat2)
    SchalterParamStatusImp = code2
    If code2 = 0 Then
        nomeErro = "livre"
    ElseIf code2 = 65 Then
        nomeErro = "em venda"
    ElseIf code2 = 90 Then
        nomeErro = "com cupom aberto"
    ElseIf code2 = 99 Then
        nomeErro = "em intervenção técnica"
    ElseIf code2 = 100 Then
        nomeErro = "em período de venda"
    ElseIf code2 = 113 Then
        nomeErro = "esperando fechamento"
    ElseIf code2 = 115 Then
        nomeErro = "com o fechamento do dia feito"
    ElseIf code2 = 122 Then
        nomeErro = "em relatório"
    ElseIf code2 = 123 Then
        nomeErro = "em pagamento"
    ElseIf code2 = 124 Then
        nomeErro = "em linha comercial"
    End If
    If SchalterParamStatusImp = 100 Then
        SchalterParamStatusImp = 0
    End If
    If (SchalterParamStatusImp <> 0) Then
        MsgBox "Erro de número: " & SchalterParamStatusImp & Chr(10) & "Descrição do erro: " & nomeErro, vbInformation, "Erro de Comunicação! - Schalter"
        Exit Function
    End If
    Exit Function
    teste = code3 And 2                    'teste de pouco papel
    If teste <> 0 Then
         MsgBox "Pouco Papel: Sim"
    Else
         MsgBox "Pouco Papel: Não"
    End If
    teste = code3 And 4                    'teste de sensor de autenticação
    If teste <> 0 Then
        MsgBox "Sensor de autenticação: Não"
    Else
        MsgBox "Sensor de autenticação: Sim"
    End If
    'teste = code3 And 8                    'teste do sensor da gaveta
    'If (teste <> 0) Then
    '    MsgBox "Sensor de Gaveta: Não"
    'Else
    '    MsgBox "Sensor de Gaveta: Sim"
    'End If
End Function
Function SchalterParamStatusCup(pbyStat1 As String, pbyStat2 As String, pbyStat3 As String, pbyStat4 As String, pbyStat5 As String, pbyStat6 As String, pbyStat7 As String) As String
    SchalterParamStatusCup = ecfParamStatusCup(pbyStat1, pbyStat2, pbyStat3, pbyStat4, pbyStat5, pbyStat6, pbyStat7)
    If (SchalterParamStatusCup <> 0) Then
        MsgBox "ERRO! Ocorreu o erro de código " & SchalterParamStatusCup
    End If
    Exit Function
    MsgBox "Numero do ECF: " & pbyStat1
    MsgBox "Tipo do Documento: " & pbyStat2
    MsgBox "Numero do Cupom: " & pbyStat3
    MsgBox "Data: " & pbyStat4
    MsgBox "Hora: " & pbyStat5
End Function
Function SchalterVendaDeItem(xCodigoProduto As String, xNomeProduto As String, xQuantidade As String, xValor As String, xTaxa As Integer, xUn As String, xDigitos As String) As Integer
    SchalterVendaDeItem = ecfVendaItem3d(xCodigoProduto, xNomeProduto, xQuantidade, xValor, xTaxa, xUn, xDigitos)
End Function
Function SchalterEfetuaPagamento(xTipo As Integer, xTabela As String, xValor As String, xLinhaMens As Integer) As Integer
    'xTipo = 0
    'xTabela = "01"
    'xValor = "0000004010"
    'xLinhaMens = 0
    SchalterEfetuaPagamento = ecfPagamento(xTipo, xTabela, xValor, xLinhaMens)
End Function
Function SchalterStatusCupom(xFlagGeral As Integer) As String
    xFlagGeral = 0
    SchalterStatusCupom = ecfStatusCupom(xFlagGeral)
End Function

