VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaixaNotaAbastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoCliente As Long
Private mDataPagamento As Date
Private mCodigoProduto2 As Long
Private mNumeroNota As Long
Private mEmpresa As Integer
Private mPeriodo As String
Private mDataAbastecimento As Date
Private mQuantidade As Currency
Private mValorUnitario As Currency
Private mValorTotal As Currency
Private mCodigoConveniado As Long
Private mTipoMovimento As String
Private mPlacaLetra As String
Private mPlacaNumero As String
Private mHistorico As String
Private mValorPago As Currency
Private mValorDescontoUnitario As Currency
Private mNumeroMovimentoCaixa As Long
Private mNumeroMovimentoCaixaBaixa As Long
Private mBaixadoPelaDuplicata As Boolean
Private mNumeroIlha As Integer
Private mOrigem As String
Private mDataBaixa As Date
Private mNumeroCupom As Long
Private mDataConferencia As Date
Private mNumeroDuplicata As Long
Private mOrdem As Integer
Private mKM As Currency

Private rsBaixaNotaAbastecimento As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get DataPagamento() As Date
    DataPagamento = mDataPagamento
End Property
Public Property Let DataPagamento(ByVal Valor As Date)
    mDataPagamento = Valor
End Property


Public Property Get CodigoProduto2() As Long
    CodigoProduto2 = mCodigoProduto2
End Property
Public Property Let CodigoProduto2(ByVal Valor As Long)
    mCodigoProduto2 = Valor
End Property


Public Property Get NumeroNota() As Long
    NumeroNota = mNumeroNota
End Property
Public Property Let NumeroNota(ByVal Valor As Long)
    mNumeroNota = Valor
End Property


Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property


Public Property Get DataAbastecimento() As Date
    DataAbastecimento = mDataAbastecimento
End Property
Public Property Let DataAbastecimento(ByVal Valor As Date)
    mDataAbastecimento = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorUnitario() As Currency
    ValorUnitario = mValorUnitario
End Property
Public Property Let ValorUnitario(ByVal Valor As Currency)
    mValorUnitario = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get CodigoConveniado() As Long
    CodigoConveniado = mCodigoConveniado
End Property
Public Property Let CodigoConveniado(ByVal Valor As Long)
    mCodigoConveniado = Valor
End Property


Public Property Get TipoMovimento() As String
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As String)
    mTipoMovimento = Valor
End Property


Public Property Get PlacaLetra() As String
    PlacaLetra = mPlacaLetra
End Property
Public Property Let PlacaLetra(ByVal Valor As String)
    mPlacaLetra = Valor
End Property


Public Property Get PlacaNumero() As String
    PlacaNumero = mPlacaNumero
End Property
Public Property Let PlacaNumero(ByVal Valor As String)
    mPlacaNumero = Valor
End Property


Public Property Get Historico() As String
    Historico = mHistorico
End Property
Public Property Let Historico(ByVal Valor As String)
    mHistorico = Valor
End Property


Public Property Get ValorPago() As Currency
    ValorPago = mValorPago
End Property
Public Property Let ValorPago(ByVal Valor As Currency)
    mValorPago = Valor
End Property


Public Property Get ValorDescontoUnitario() As Currency
    ValorDescontoUnitario = mValorDescontoUnitario
End Property
Public Property Let ValorDescontoUnitario(ByVal Valor As Currency)
    mValorDescontoUnitario = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get NumeroMovimentoCaixaBaixa() As Long
    NumeroMovimentoCaixaBaixa = mNumeroMovimentoCaixaBaixa
End Property
Public Property Let NumeroMovimentoCaixaBaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixaBaixa = Valor
End Property


Public Property Get BaixadoPelaDuplicata() As Boolean
    BaixadoPelaDuplicata = mBaixadoPelaDuplicata
End Property
Public Property Let BaixadoPelaDuplicata(ByVal Valor As Boolean)
    mBaixadoPelaDuplicata = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get Origem() As String
    Origem = mOrigem
End Property
Public Property Let Origem(ByVal Valor As String)
    mOrigem = Valor
End Property


Public Property Get DataBaixa() As Date
    DataBaixa = mDataBaixa
End Property
Public Property Let DataBaixa(ByVal Valor As Date)
    mDataBaixa = Valor
End Property


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
End Property


Public Property Get DataConferencia() As Date
    DataConferencia = mDataConferencia
End Property
Public Property Let DataConferencia(ByVal Valor As Date)
    mDataConferencia = Valor
End Property


Public Property Get NumeroDuplicata() As Long
    NumeroDuplicata = mNumeroDuplicata
End Property
Public Property Let NumeroDuplicata(ByVal Valor As Long)
    mNumeroDuplicata = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get KM() As Currency
    KM = mKM
End Property
Public Property Let KM(ByVal Valor As Currency)
    mKM = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataAbastecimento As Date, ByVal pNumeroNota As Long, ByVal pOrdem As Integer, ByVal pCodigoProduto2 As Long, ByVal pPeriodo As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Cliente] = " & pCodigoCliente
    xCondicao = xCondicao & " AND [Data do Abastecimento] = " & preparaData(pDataAbastecimento)
    xCondicao = xCondicao & " AND [Numero da Nota] = " & pNumeroNota
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    xCondicao = xCondicao & " AND [Codigo do Produto2] = " & pCodigoProduto2
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(pPeriodo)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior(ByVal nCodigoCliente As Long) As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xSQL = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    xSQL = xSQL & " AND [Data do Abastecimento] <= " & preparaData(mDataAbastecimento)
    gSQL = PreparaSQL(xSQL, "ORDER BY [Data do Abastecimento] DESC, [Codigo do Produto2] DESC, [Numero da Nota] DESC, Ordem DESC, [Data do Pagamento] DESC, Empresa DESC, Periodo DESC")
    If Localizar(55) Then
        If rsBaixaNotaAbastecimento.RecordCount > 1 Then
            Do Until rsBaixaNotaAbastecimento.EOF
                If xRegIgual = False Then
                    If rsBaixaNotaAbastecimento("Data do Abastecimento").Value = mDataAbastecimento And rsBaixaNotaAbastecimento("Codigo do Produto2").Value = mCodigoProduto2 And rsBaixaNotaAbastecimento("Numero da Nota").Value = mNumeroNota And rsBaixaNotaAbastecimento("Ordem").Value = mOrdem And rsBaixaNotaAbastecimento("Data do Pagamento").Value = mDataPagamento And rsBaixaNotaAbastecimento("Empresa").Value = mEmpresa And rsBaixaNotaAbastecimento("Periodo").Value = mPeriodo Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsBaixaNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nCodigoCliente As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] ASC, [Codigo do Produto2] ASC, [Numero da Nota] ASC, Ordem ASC, [Data do Pagamento] ASC, Empresa ASC, Periodo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo(ByVal nCodigoCliente As Long) As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xSQL = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    xSQL = xSQL & " AND [Data do Abastecimento] >= " & preparaData(mDataAbastecimento)
    gSQL = PreparaSQL(xSQL, "ORDER BY [Codigo do Cliente] ASC, [Data do Abastecimento] ASC, [Codigo do Produto2] ASC, [Numero da Nota] ASC, Ordem ASC, [Data do Pagamento] ASC, Empresa ASC, Periodo ASC")
    If Localizar(55) Then
        If rsBaixaNotaAbastecimento.RecordCount > 1 Then
            Do Until rsBaixaNotaAbastecimento.EOF
                If xRegIgual = False Then
                    If rsBaixaNotaAbastecimento("Data do Abastecimento").Value = mDataAbastecimento And rsBaixaNotaAbastecimento("Codigo do Produto2").Value = mCodigoProduto2 And rsBaixaNotaAbastecimento("Numero da Nota").Value = mNumeroNota And rsBaixaNotaAbastecimento("Ordem").Value = mOrdem And rsBaixaNotaAbastecimento("Data do Pagamento").Value = mDataPagamento And rsBaixaNotaAbastecimento("Empresa").Value = mEmpresa And rsBaixaNotaAbastecimento("Periodo").Value = mPeriodo Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsBaixaNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nCodigoCliente As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] DESC, [Codigo do Produto2] DESC, [Numero da Nota] DESC, Ordem DESC, [Data do Pagamento] DESC, Empresa DESC, Periodo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoRegistro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoRegistro = False
    xCondicao = ""
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] DESC, [Data do Abastecimento] DESC, [Codigo do Produto2] DESC, [Numero da Nota] DESC, Ordem DESC, [Data do Pagamento] DESC, Empresa DESC, Periodo DESC")
    If Localizar(1) Then
        LocalizarUltimoRegistro = True
        AtribuiValor
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Baixa_Nota_Abastecimento ( [Codigo do Cliente], [Data do Pagamento], [Codigo do Produto2], [Numero da Nota], Empresa, "
    gSQL = gSQL & "Periodo, [Data do Abastecimento], Quantidade, [Valor Unitario], [Valor Total], "
    gSQL = gSQL & "[Codigo do Conveniado], [Tipo do Movimento], [Placa Letra], [Placa Numero], Historico, "
    gSQL = gSQL & "[Valor Pago], [Valor Desconto Unitario], [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], "
    gSQL = gSQL & "[Baixado Pela Duplicata], [Numero da Ilha], Origem, [Data da Baixa], [Numero do Cupom], [Data da Conferencia], "
    gSQL = gSQL & "[Numero da Duplicata], Ordem, KM ) VALUES ( "
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlData(1, mDataPagamento, ", ")
    Call sqlNumero(1, mCodigoProduto2, ", ")
    Call sqlNumero(1, mNumeroNota, ", ")
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlData(1, mDataAbastecimento, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlNumero(1, mCodigoConveniado, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlTexto(1, mPlacaLetra, ", ")
    Call sqlTexto(1, mPlacaNumero, ", ")
    Call sqlTexto(1, mHistorico, ", ")
    Call sqlValor(1, mValorPago, ", ")
    Call sqlValor(1, mValorDescontoUnitario, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaBaixa, ", ")
    Call sqlBoolean(1, mBaixadoPelaDuplicata, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlTexto(1, mOrigem, ", ")
    Call sqlData(1, mDataBaixa, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlData(1, mDataConferencia, ", ")
    Call sqlNumero(1, mNumeroDuplicata, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlValor(1, mKM, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataAbastecimento As Date, ByVal pNumeroNota As Long, ByVal pOrdem As Integer, ByVal pCodigoProduto2 As Long, ByVal pPeriodo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Baixa_Nota_Abastecimento SET "
    Call sqlNumero(2, "[Codigo do Cliente] = ", mCodigoCliente)
    Call sqlData(2, ", [Data do Pagamento] = ", mDataPagamento)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto2)
    Call sqlNumero(2, ", [Numero da Nota] = ", mNumeroNota)
    Call sqlNumero(2, ", Empresa = ", mEmpresa)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlData(2, ", [Data do Abastecimento] = ", mDataAbastecimento)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlNumero(2, ", [Codigo do Conveniado] = ", mCodigoConveniado)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlTexto(2, ", [Placa Letra] = ", mPlacaLetra)
    Call sqlTexto(2, ", [Placa Numero] = ", mPlacaNumero)
    Call sqlTexto(2, ", Historico = ", mHistorico)
    Call sqlValor(2, ", [Valor Pago] = ", mValorPago)
    Call sqlValor(2, ", [Valor Desconto Unitario] = ", mValorDescontoUnitario)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Baixa] = ", mNumeroMovimentoCaixaBaixa)
    Call sqlBoolean(2, ", [Baixado Pela Duplicata] = ", mBaixadoPelaDuplicata)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlTexto(2, ", Origem = ", mOrigem)
    Call sqlData(2, ", [Data da Baixa] = ", mDataBaixa)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlData(2, ", [Data da Conferencia] = ", mDataConferencia)
    Call sqlNumero(2, ", [Numero da Duplicata] = ", mNumeroDuplicata)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlValor(2, ", KM = ", mKM)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & " AND [Data do Abastecimento] = " & preparaData(pDataAbastecimento)
    gSQL = gSQL & " AND [Numero da Nota] = " & pNumeroNota
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataAbastecimento As Date, ByVal pNumeroNota As Long, ByVal pOrdem As Integer, ByVal pCodigoProduto2 As Long, ByVal pPeriodo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Baixa_Nota_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & " AND [Data do Abastecimento] = " & preparaData(pDataAbastecimento)
    gSQL = gSQL & " AND [Numero da Nota] = " & pNumeroNota
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Cliente], [Data do Pagamento], [Codigo do Produto2], [Numero da Nota], Empresa, Periodo, [Data do Abastecimento], Quantidade, [Valor Unitario], [Valor Total], [Codigo do Conveniado], [Tipo do Movimento], [Placa Letra], [Placa Numero], Historico, [Valor Pago], [Valor Desconto Unitario], [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], [Baixado Pela Duplicata], [Numero da Ilha], Origem, [Data da Baixa], [Numero do Cupom], [Data da Conferencia], [Numero da Duplicata], Ordem, KM FROM Baixa_Nota_Abastecimento" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBaixaNotaAbastecimento = New adodb.Recordset
    rsBaixaNotaAbastecimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBaixaNotaAbastecimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBaixaNotaAbastecimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoCliente = rsBaixaNotaAbastecimento("Codigo do Cliente").Value
    DataPagamento = rsBaixaNotaAbastecimento("Data do Pagamento").Value
    CodigoProduto2 = rsBaixaNotaAbastecimento("Codigo do Produto2").Value
    NumeroNota = rsBaixaNotaAbastecimento("Numero da Nota").Value
    Empresa = rsBaixaNotaAbastecimento("Empresa").Value
    Periodo = rsBaixaNotaAbastecimento("Periodo").Value
    DataAbastecimento = rsBaixaNotaAbastecimento("Data do Abastecimento").Value
    Quantidade = rsBaixaNotaAbastecimento("Quantidade").Value
    ValorUnitario = rsBaixaNotaAbastecimento("Valor Unitario").Value
    ValorTotal = rsBaixaNotaAbastecimento("Valor Total").Value
    CodigoConveniado = rsBaixaNotaAbastecimento("Codigo do Conveniado").Value
    TipoMovimento = rsBaixaNotaAbastecimento("Tipo do Movimento").Value
    If IsNull(rsBaixaNotaAbastecimento("Placa Letra").Value) Then
        PlacaLetra = ""
    Else
        PlacaLetra = rsBaixaNotaAbastecimento("Placa Letra").Value
    End If
    If IsNull(rsBaixaNotaAbastecimento("Placa Numero").Value) Then
        PlacaNumero = ""
    Else
        PlacaNumero = rsBaixaNotaAbastecimento("Placa Numero").Value
    End If
    If IsNull(rsBaixaNotaAbastecimento("Historico").Value) Then
        Historico = ""
    Else
        Historico = rsBaixaNotaAbastecimento("Historico").Value
    End If
    ValorPago = rsBaixaNotaAbastecimento("Valor Pago").Value
    If IsNull(rsBaixaNotaAbastecimento("Valor Desconto Unitario").Value) Then
        ValorDescontoUnitario = 0
    Else
        ValorDescontoUnitario = rsBaixaNotaAbastecimento("Valor Desconto Unitario").Value
    End If
    NumeroMovimentoCaixa = rsBaixaNotaAbastecimento("Numero do Movimento do Caixa").Value
    NumeroMovimentoCaixaBaixa = rsBaixaNotaAbastecimento("Numero do Movimento do Caixa Baixa").Value
    BaixadoPelaDuplicata = rsBaixaNotaAbastecimento("Baixado Pela Duplicata").Value
    NumeroIlha = rsBaixaNotaAbastecimento("Numero da Ilha").Value
    Origem = rsBaixaNotaAbastecimento("Origem").Value
    DataBaixa = rsBaixaNotaAbastecimento("Data da Baixa").Value
    NumeroCupom = rsBaixaNotaAbastecimento("Numero do Cupom").Value
    If IsNull(rsBaixaNotaAbastecimento("Data da Conferencia").Value) Then
        DataConferencia = "00:00:00"
    Else
        DataConferencia = rsBaixaNotaAbastecimento("Data da Conferencia").Value
    End If
    NumeroDuplicata = rsBaixaNotaAbastecimento("Numero da Duplicata").Value
    Ordem = rsBaixaNotaAbastecimento("Ordem").Value
    KM = rsBaixaNotaAbastecimento("KM").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function TotalData(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pConferencia As Boolean) As Currency

On Error GoTo trata_erro
    TotalData = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Baixa_Nota_Abastecimento"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    Else
        gSQL = gSQL & " WHERE Empresa > 0"
    End If
    If pCodigoCliente > 0 Then
        gSQL = gSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    End If
    If pConferencia Then
        gSQL = gSQL & " AND [Data da Conferencia] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data da Conferencia] <= " & preparaData(pDataFinal)
    Else
        gSQL = gSQL & " AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    End If
    If Localizar(0) Then
        If Not IsNull(rsBaixaNotaAbastecimento("Total").Value) Then
            TotalData = rsBaixaNotaAbastecimento("Total").Value
        End If
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalDataLiquido(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pConferencia As Boolean) As Currency
    Dim xValor As Currency

On Error GoTo trata_erro

    xValor = 0
    TotalDataLiquido = 0
    gSQL = "SELECT Quantidade, [Valor Desconto Unitario], [Valor Unitario], [Valor Total] FROM Baixa_Nota_Abastecimento"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    Else
        gSQL = gSQL & " WHERE Empresa > 0"
    End If
    If pCodigoCliente > 0 Then
        gSQL = gSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    End If
    If pConferencia Then
        gSQL = gSQL & " AND [Data da Conferencia] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data da Conferencia] <= " & preparaData(pDataFinal)
    Else
        gSQL = gSQL & " AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
        gSQL = gSQL & " AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    End If
    If Localizar(0) Then
        If rsBaixaNotaAbastecimento.RecordCount > 0 Then
            Do Until rsBaixaNotaAbastecimento.EOF
                If rsBaixaNotaAbastecimento("Valor Desconto Unitario").Value <> 0 Then
                    xValor = xValor + Format((rsBaixaNotaAbastecimento("Valor Unitario").Value - rsBaixaNotaAbastecimento("Valor Desconto Unitario").Value) * rsBaixaNotaAbastecimento("Quantidade").Value, "0000000000.00")
                Else
                    xValor = xValor + rsBaixaNotaAbastecimento("Valor Total").Value
                End If
                rsBaixaNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    TotalDataLiquido = xValor

    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

