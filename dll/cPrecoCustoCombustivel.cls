VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrecoCustoCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoCombustivel As String
Private mData As Date
Private mOrdem As Integer
Private mQuantidadeInicial As Currency
Private mValorInicial As Currency
Private mQuantidadeEntrada As Currency
Private mValorEntrada As Currency
Private mQuantidadeVenda As Currency
Private mValorVenda As Currency
Private mQuantidadePrecoAnterior As Currency

Private rsPrecoCustoCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property


Public Property Get QuantidadeInicial() As Currency
    QuantidadeInicial = mQuantidadeInicial
End Property
Public Property Let QuantidadeInicial(ByVal Valor As Currency)
    mQuantidadeInicial = Valor
End Property


Public Property Get ValorInicial() As Currency
    ValorInicial = mValorInicial
End Property
Public Property Let ValorInicial(ByVal Valor As Currency)
    mValorInicial = Valor
End Property


Public Property Get QuantidadeEntrada() As Currency
    QuantidadeEntrada = mQuantidadeEntrada
End Property
Public Property Let QuantidadeEntrada(ByVal Valor As Currency)
    mQuantidadeEntrada = Valor
End Property


Public Property Get ValorEntrada() As Currency
    ValorEntrada = mValorEntrada
End Property
Public Property Let ValorEntrada(ByVal Valor As Currency)
    mValorEntrada = Valor
End Property


Public Property Get QuantidadeVenda() As Currency
    QuantidadeVenda = mQuantidadeVenda
End Property
Public Property Let QuantidadeVenda(ByVal Valor As Currency)
    mQuantidadeVenda = Valor
End Property


Public Property Get ValorVenda() As Currency
    ValorVenda = mValorVenda
End Property
Public Property Let ValorVenda(ByVal Valor As Currency)
    mValorVenda = Valor
End Property


Public Property Get QuantidadePrecoAnterior() As Currency
    QuantidadePrecoAnterior = mQuantidadePrecoAnterior
End Property
Public Property Let QuantidadePrecoAnterior(ByVal Valor As Currency)
    mQuantidadePrecoAnterior = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND Ordem = " & nOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] < " & preparaTexto(mTipoCombustivel)
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem < " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Tipo de Combustivel] DESC, Data DESC, Ordem DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] > " & preparaTexto(mTipoCombustivel)
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem > " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Tipo de Combustivel] ASC, Data ASC, Ordem ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] > " & preparaTexto(mTipoCombustivel)
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem > " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Tipo de Combustivel] ASC, Data ASC, Ordem ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
    xCondicao = xCondicao & " AND Data = " & preparaData(mData)
    xCondicao = xCondicao & " AND Ordem = " & mOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Tipo de Combustivel] DESC, Data DESC, Ordem DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir(ByVal xTipoCusto As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Preco_Custo_Combustivel ( Empresa, [Tipo de Combustivel], Data, Ordem, [Quantidade Inicial], "
    gSQL = gSQL & "[Valor Inicial], [Quantidade de Entrada], [Valor de Entrada], [Quantidade da Venda], [Valor da Venda], "
    gSQL = gSQL & "[Quantidade no Preco Anterior] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlValor(1, mQuantidadeInicial, ", ")
    If xTipoCusto = "CustoReal" Then
        mValorInicial = CalculaValorInicialCustoReal
    End If
    Call sqlValor(1, mValorInicial, ", ")
    Call sqlValor(1, mQuantidadeEntrada, ", ")
    Call sqlValor(1, mValorEntrada, ", ")
    Call sqlValor(1, mQuantidadeVenda, ", ")
    If xTipoCusto = "CustoReal" Then
       mValorVenda = CalculaValorVendaCustoReal
    End If
    Call sqlValor(1, mValorVenda, ", ")
    Call sqlValor(1, mQuantidadePrecoAnterior, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Preco_Custo_Combustivel SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlValor(2, ", [Quantidade Inicial] = ", mQuantidadeInicial)
    Call sqlValor(2, ", [Valor Inicial] = ", mValorInicial)
    Call sqlValor(2, ", [Quantidade de Entrada] = ", mQuantidadeEntrada)
    Call sqlValor(2, ", [Valor de Entrada] = ", mValorEntrada)
    Call sqlValor(2, ", [Quantidade da Venda] = ", mQuantidadeVenda)
    Call sqlValor(2, ", [Valor da Venda] = ", mValorVenda)
    Call sqlValor(2, ", [Quantidade no Preco Anterior] = ", mQuantidadePrecoAnterior)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Ordem = " & nOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Private Function CalculaValorInicialCustoReal() As Currency
Dim xSQL As String

On Error GoTo trata_erro

    CalculaValorInicialCustoReal = mValorInicial
    xSQL = ""
    xSQL = xSQL & " SELECT TOP 1 [Valor de Entrada], [Quantidade de Entrada]"
    xSQL = xSQL & "   FROM Preco_Custo_Combustivel"
    xSQL = xSQL & "  WHERE Empresa = " & mEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
    xSQL = xSQL & "    AND Data < " & preparaData(mData)
    xSQL = xSQL & "  ORDER BY Empresa DESC, [Tipo de Combustivel] DESC, Data DESC"
    Set rsPrecoCustoCombustivel = New adodb.Recordset
    rsPrecoCustoCombustivel.CursorLocation = adUseClient
    rsPrecoCustoCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrecoCustoCombustivel.EOF Then
        If rsPrecoCustoCombustivel("Quantidade de Entrada").Value > 0 Then
            CalculaValorInicialCustoReal = rsPrecoCustoCombustivel("Valor de Entrada").Value
        End If
    End If
    rsPrecoCustoCombustivel.Close
    Set rsPrecoCustoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function CalculaValorVendaCustoReal() As Currency
Dim xSQL As String

On Error GoTo trata_erro

    CalculaValorVendaCustoReal = mValorInicial
    If mQuantidadeEntrada > 0 Then
        'xSQL = ""
        'xSQL = xSQL & " SELECT TOP 1 [Valor Inicial]"
        'xSQL = xSQL & "   FROM Preco_Custo_Combustivel"
        'xSQL = xSQL & "  WHERE Empresa = " & mEmpresa
        'xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
        'xSQL = xSQL & "    AND Data < " & preparaData(mData)
        'xSQL = xSQL & "  ORDER BY Empresa DESC, [Tipo de Combustivel] DESC, Data DESC"
        'Set rsPrecoCustoCombustivel = New adodb.Recordset
        'rsPrecoCustoCombustivel.CursorLocation = adUseClient
        'rsPrecoCustoCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
        'If Not rsPrecoCustoCombustivel.EOF Then
            If mQuantidadeInicial <= 0 Then
                If mValorInicial > mValorEntrada Then
                    CalculaValorVendaCustoReal = mValorInicial
                Else
                    CalculaValorVendaCustoReal = mValorEntrada
                End If
            ElseIf mQuantidadeVenda > mQuantidadeInicial Then
                CalculaValorVendaCustoReal = ((mQuantidadeInicial * mValorInicial) + ((mQuantidadeVenda - mQuantidadeInicial) * mValorEntrada)) / mQuantidadeVenda
            Else
                CalculaValorVendaCustoReal = mValorInicial
            End If
        'End If
        'rsPrecoCustoCombustivel.Close
        'Set rsPrecoCustoCombustivel = Nothing
    End If
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Preco_Custo_Combustivel"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Ordem = " & nOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Tipo de Combustivel], Data, Ordem, [Quantidade Inicial], [Valor Inicial], [Quantidade de Entrada], [Valor de Entrada], [Quantidade da Venda], [Valor da Venda], [Quantidade no Preco Anterior] FROM Preco_Custo_Combustivel" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPrecoCustoCombustivel = New adodb.Recordset
    rsPrecoCustoCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPrecoCustoCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrecoCustoCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsPrecoCustoCombustivel("Empresa").Value
    TipoCombustivel = rsPrecoCustoCombustivel("Tipo de Combustivel").Value
    Data = rsPrecoCustoCombustivel("Data").Value
    Ordem = rsPrecoCustoCombustivel("Ordem").Value
    QuantidadeInicial = rsPrecoCustoCombustivel("Quantidade Inicial").Value
    ValorInicial = rsPrecoCustoCombustivel("Valor Inicial").Value
    QuantidadeEntrada = rsPrecoCustoCombustivel("Quantidade de Entrada").Value
    ValorEntrada = rsPrecoCustoCombustivel("Valor de Entrada").Value
    QuantidadeVenda = rsPrecoCustoCombustivel("Quantidade da Venda").Value
    ValorVenda = rsPrecoCustoCombustivel("Valor da Venda").Value
    QuantidadePrecoAnterior = rsPrecoCustoCombustivel("Quantidade no Preco Anterior").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

