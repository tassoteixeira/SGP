VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoNotaAbastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoCliente As Long
Private mDataAbastecimento As Date
Private mNumeroNota As Long
Private mCodigoProduto2 As Long
Private mPeriodo As String
Private mQuantidade As Currency
Private mValorUnitario As Currency
Private mValorTotal As Currency
Private mCodigoConveniado As Long
Private mTipoMovimento As String
Private mPlacaLetra As String
Private mPlacaNumero As String
Private mHistorico As String
Private mNumeroCupom As Long
Private mValorDescontoUnitario As Currency
Private mNumeroMovimentoCaixa As Long
Private mBaixadoPelaDuplicata As Boolean
Private mNumeroIlha As Integer
Private mOrigem As String
Private mDataConferencia As Date
Private mOrdem As Integer
Private mKM As Currency

Private rsMovNotaAbastecimento As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get DataAbastecimento() As Date
    DataAbastecimento = mDataAbastecimento
End Property
Public Property Let DataAbastecimento(ByVal Valor As Date)
    mDataAbastecimento = Valor
End Property


Public Property Get NumeroNota() As Long
    NumeroNota = mNumeroNota
End Property
Public Property Let NumeroNota(ByVal Valor As Long)
    mNumeroNota = Valor
End Property


Public Property Get CodigoProduto2() As Long
    CodigoProduto2 = mCodigoProduto2
End Property
Public Property Let CodigoProduto2(ByVal Valor As Long)
    mCodigoProduto2 = Valor
End Property


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
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


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
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


Public Property Get DataConferencia() As Date
    DataConferencia = mDataConferencia
End Property
Public Property Let DataConferencia(ByVal Valor As Date)
    mDataConferencia = Valor
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
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xRegIgual As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data do Abastecimento] <= " & preparaData(mDataAbastecimento)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] DESC, Periodo DESC, [Numero da Nota] DESC, Ordem DESC, [Codigo do Cliente] DESC, [Codigo do Produto2] DESC")
    If Localizar(55) Then
        If rsMovNotaAbastecimento.RecordCount > 1 Then
            Do Until rsMovNotaAbastecimento.EOF
                If xRegIgual = False Then
                    If rsMovNotaAbastecimento("Data do Abastecimento").Value = mDataAbastecimento And rsMovNotaAbastecimento("Periodo").Value = mPeriodo And rsMovNotaAbastecimento("Numero da Nota").Value = mNumeroNota And rsMovNotaAbastecimento("Ordem").Value = mOrdem And rsMovNotaAbastecimento("Codigo do Cliente").Value = mCodigoCliente And rsMovNotaAbastecimento("Codigo do Produto2").Value = mCodigoProduto2 Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] ASC, Periodo ASC, [Numero da Nota] ASC, Ordem ASC, [Codigo do Cliente] ASC, [Codigo do Produto2] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
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
    xCondicao = xCondicao & " AND [Data do Abastecimento] >= " & preparaData(mDataAbastecimento)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] ASC, Periodo ASC, [Numero da Nota] ASC, Ordem ASC, [Codigo do Cliente] ASC, [Codigo do Produto2] ASC")
    If Localizar(55) Then
        If rsMovNotaAbastecimento.RecordCount > 1 Then
            Do Until rsMovNotaAbastecimento.EOF
                If xRegIgual = False Then
                    If rsMovNotaAbastecimento("Data do Abastecimento").Value = mDataAbastecimento And rsMovNotaAbastecimento("Periodo").Value = mPeriodo And rsMovNotaAbastecimento("Numero da Nota").Value = mNumeroNota And rsMovNotaAbastecimento("Ordem").Value = mOrdem And rsMovNotaAbastecimento("Codigo do Cliente").Value = mCodigoCliente And rsMovNotaAbastecimento("Codigo do Produto2").Value = mCodigoProduto2 Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Abastecimento] DESC, Periodo DESC, [Numero da Nota] DESC, Ordem DESC, [Codigo do Cliente] DESC, [Codigo do Produto2] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Nota_Abastecimento ( Empresa, [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], [Codigo do Produto2], "
    gSQL = gSQL & "Periodo, Quantidade, [Valor Unitario], [Valor Total], [Codigo do Conveniado], "
    gSQL = gSQL & "[Tipo do Movimento], [Placa Letra], [Placa Numero], Historico, [Numero do Cupom], "
    gSQL = gSQL & "[Valor Desconto Unitario], [Numero do Movimento do Caixa], [Baixado Pela Duplicata], "
    gSQL = gSQL & "[Numero da Ilha], Origem, [Data da Conferencia], Ordem, KM ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlData(1, mDataAbastecimento, ", ")
    Call sqlNumero(1, mNumeroNota, ", ")
    Call sqlNumero(1, mCodigoProduto2, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlNumero(1, mCodigoConveniado, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlTexto(1, mPlacaLetra, ", ")
    Call sqlTexto(1, mPlacaNumero, ", ")
    Call sqlTexto(1, mHistorico, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlValor(1, mValorDescontoUnitario, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlBoolean(1, mBaixadoPelaDuplicata, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlTexto(1, mOrigem, ", ")
    Call sqlData(1, mDataConferencia, ", ")
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
    gSQL = "UPDATE Movimento_Nota_Abastecimento SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    Call sqlData(2, ", [Data do Abastecimento] = ", mDataAbastecimento)
    Call sqlNumero(2, ", [Numero da Nota] = ", mNumeroNota)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto2)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlNumero(2, ", [Codigo do Conveniado] = ", mCodigoConveniado)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlTexto(2, ", [Placa Letra] = ", mPlacaLetra)
    Call sqlTexto(2, ", [Placa Numero] = ", mPlacaNumero)
    Call sqlTexto(2, ", Historico = ", mHistorico)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlValor(2, ", [Valor Desconto Unitario] = ", mValorDescontoUnitario)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlBoolean(2, ", [Baixado Pela Duplicata] = ", mBaixadoPelaDuplicata)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlTexto(2, ", Origem = ", mOrigem)
    Call sqlData(2, ", [Data da Conferencia] = ", mDataConferencia)
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


Public Function AlterarPlacaKM(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataAbastecimento As Date, ByVal pNumeroNota As Long, ByVal pPeriodo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarPlacaKM = False
    gSQL = "UPDATE Movimento_Nota_Abastecimento SET "
    Call sqlTexto(2, "[Placa Letra] = ", mPlacaLetra)
    Call sqlTexto(2, ", [Placa Numero] = ", mPlacaNumero)
    Call sqlValor(2, ", KM = ", mKM)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & " AND [Data do Abastecimento] = " & preparaData(pDataAbastecimento)
    gSQL = gSQL & " AND [Numero da Nota] = " & pNumeroNota
    gSQL = gSQL & " AND Periodo = " & preparaTexto(pPeriodo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarPlacaKM = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataAbastecimento As Date, ByVal pNumeroNota As Long, ByVal pOrdem As Integer, ByVal pCodigoProduto2 As Long, ByVal pPeriodo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Nota_Abastecimento"
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
    PreparaSQL = Trim("SELECT Empresa, [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], [Codigo do Produto2], Periodo, Quantidade, [Valor Unitario], [Valor Total], [Codigo do Conveniado], [Tipo do Movimento], [Placa Letra], [Placa Numero], Historico, [Numero do Cupom], [Valor Desconto Unitario], [Numero do Movimento do Caixa], [Baixado Pela Duplicata], [Numero da Ilha], Origem, [Data da Conferencia], Ordem, KM FROM Movimento_Nota_Abastecimento" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovNotaAbastecimento = New adodb.Recordset
    rsMovNotaAbastecimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovNotaAbastecimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovNotaAbastecimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovNotaAbastecimento("Empresa").Value
    CodigoCliente = rsMovNotaAbastecimento("Codigo do Cliente").Value
    DataAbastecimento = rsMovNotaAbastecimento("Data do Abastecimento").Value
    NumeroNota = rsMovNotaAbastecimento("Numero da Nota").Value
    CodigoProduto2 = rsMovNotaAbastecimento("Codigo do Produto2").Value
    Periodo = rsMovNotaAbastecimento("Periodo").Value
    Quantidade = rsMovNotaAbastecimento("Quantidade").Value
    ValorUnitario = rsMovNotaAbastecimento("Valor Unitario").Value
    ValorTotal = rsMovNotaAbastecimento("Valor Total").Value
    CodigoConveniado = rsMovNotaAbastecimento("Codigo do Conveniado").Value
    TipoMovimento = rsMovNotaAbastecimento("Tipo do Movimento").Value
    If IsNull(rsMovNotaAbastecimento("Placa Letra").Value) Then
        PlacaLetra = ""
    Else
        PlacaLetra = rsMovNotaAbastecimento("Placa Letra").Value
    End If
    If IsNull(rsMovNotaAbastecimento("Placa Numero").Value) Then
        PlacaNumero = ""
    Else
        PlacaNumero = rsMovNotaAbastecimento("Placa Numero").Value
    End If
    If IsNull(rsMovNotaAbastecimento("Historico").Value) Then
        Historico = ""
    Else
        Historico = rsMovNotaAbastecimento("Historico").Value
    End If
    If IsNull(rsMovNotaAbastecimento("Numero do Cupom").Value) Then
        NumeroCupom = 0
    Else
        NumeroCupom = rsMovNotaAbastecimento("Numero do Cupom").Value
    End If
    If IsNull(rsMovNotaAbastecimento("Valor Desconto Unitario").Value) Then
        ValorDescontoUnitario = 0
    Else
        ValorDescontoUnitario = rsMovNotaAbastecimento("Valor Desconto Unitario").Value
    End If
    NumeroMovimentoCaixa = rsMovNotaAbastecimento("Numero do Movimento do Caixa").Value
    BaixadoPelaDuplicata = rsMovNotaAbastecimento("Baixado Pela Duplicata").Value
    NumeroIlha = rsMovNotaAbastecimento("Numero da Ilha").Value
    Origem = rsMovNotaAbastecimento("Origem").Value
    If IsNull(rsMovNotaAbastecimento("Data da Conferencia").Value) Then
        DataConferencia = "00:00:00"
    Else
        DataConferencia = rsMovNotaAbastecimento("Data da Conferencia").Value
    End If
    Ordem = rsMovNotaAbastecimento("Ordem").Value
    KM = rsMovNotaAbastecimento("KM").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function LocalizaPrimeiraData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro
    LocalizaPrimeiraData = Date
    gSQL = "SELECT TOP 1 [Data do Abastecimento] FROM Movimento_Nota_Abastecimento"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    End If
    gSQL = gSQL & " ORDER BY [Data do Abastecimento] ASC"
    If Localizar(0) Then
        If Not IsNull(rsMovNotaAbastecimento("Data do Abastecimento").Value) Then
            LocalizaPrimeiraData = rsMovNotaAbastecimento("Data do Abastecimento").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizaPrimeiraDataCliente(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long) As Date
On Error GoTo trata_erro
    
    LocalizaPrimeiraDataCliente = Date
    gSQL = "SELECT TOP 1 [Data do Abastecimento] FROM Movimento_Nota_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & " ORDER BY [Data do Abastecimento] ASC"
    If Localizar(0) Then
        If Not IsNull(rsMovNotaAbastecimento("Data do Abastecimento").Value) Then
            LocalizaPrimeiraDataCliente = rsMovNotaAbastecimento("Data do Abastecimento").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizaUltimaData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro
    LocalizaUltimaData = Date
    gSQL = "SELECT TOP 1 [Data do Abastecimento] FROM Movimento_Nota_Abastecimento"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    End If
    gSQL = gSQL & " ORDER BY [Data do Abastecimento] DESC"
    If Localizar(0) Then
        If Not IsNull(rsMovNotaAbastecimento("Data do Abastecimento").Value) Then
            LocalizaUltimaData = rsMovNotaAbastecimento("Data do Abastecimento").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ProximoNumeroNota(ByVal nEmpresa As Integer, ByVal nDataAbastecimento As Date) As Long

On Error GoTo trata_erro
    
    ProximoNumeroNota = 1
    gSQL = "SELECT [Numero da Nota] FROM Movimento_Nota_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "   AND [Data do Abastecimento] = " & preparaData(nDataAbastecimento)
    gSQL = gSQL & " ORDER BY [Numero da Nota] DESC"
    If Localizar(1) Then
        ProximoNumeroNota = rsMovNotaAbastecimento("Numero da Nota").Value + 1
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal pData As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Currency
'TotalPeriodo(g_empresa, CDate(msk_data.Text), cbo_periodo.Text, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
On Error GoTo trata_erro
    
    TotalPeriodo = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Nota_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data do Abastecimento] = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        If Not IsNull(rsMovNotaAbastecimento("Total").Value) Then
            TotalPeriodo = rsMovNotaAbastecimento("Total").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function TotalCliente(ByVal nEmpresa As Integer, ByVal nCliente As Integer) As Currency

On Error GoTo trata_erro
    
    TotalCliente = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Nota_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Cliente] = " & preparaTexto(nCliente)
'    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
'    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        If Not IsNull(rsMovNotaAbastecimento("Total").Value) Then
            TotalCliente = rsMovNotaAbastecimento("Total").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalData(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pConferencia As Boolean) As Currency

On Error GoTo trata_erro
    
    TotalData = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Nota_Abastecimento"
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
        If Not IsNull(rsMovNotaAbastecimento("Total").Value) Then
            TotalData = rsMovNotaAbastecimento("Total").Value
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function TotalDataLiquido(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pConferencia As Boolean, Optional ByVal pCondiderarBaixadas As Boolean = False) As Currency
    Dim xValor As Currency

On Error GoTo trata_erro

    xValor = 0
    TotalDataLiquido = 0
    gSQL = "SELECT Quantidade, [Valor Desconto Unitario], [Valor Unitario], [Valor Total] FROM Movimento_Nota_Abastecimento"
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
    
    If pCondiderarBaixadas Then
    
        gSQL = gSQL & " UNION "
        gSQL = gSQL & " SELECT Quantidade, [Valor Desconto Unitario], [Valor Unitario], [Valor Total] FROM Baixa_Nota_Abastecimento"
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
    
    End If
    
    If Localizar(0) Then
        If rsMovNotaAbastecimento.RecordCount > 0 Then
            Do Until rsMovNotaAbastecimento.EOF
                'Desconto
                If rsMovNotaAbastecimento("Valor Desconto Unitario").Value > 0 Then
                    xValor = xValor + rsMovNotaAbastecimento("Valor Total").Value
                    xValor = xValor - Format(rsMovNotaAbastecimento("Valor Desconto Unitario").Value * rsMovNotaAbastecimento("Quantidade").Value, "0000000000.00")
                'Acrescimo
                ElseIf rsMovNotaAbastecimento("Valor Desconto Unitario").Value < 0 Then
                    xValor = xValor + rsMovNotaAbastecimento("Valor Total").Value
                    xValor = xValor + Format(rsMovNotaAbastecimento("Valor Desconto Unitario").Value * -1 * rsMovNotaAbastecimento("Quantidade").Value, "0000000000.00")
                Else
                    xValor = xValor + rsMovNotaAbastecimento("Valor Total").Value
                End If
                rsMovNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    rsMovNotaAbastecimento.Close
    Set rsMovNotaAbastecimento = Nothing
    TotalDataLiquido = xValor

    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


