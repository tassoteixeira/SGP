VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCupomFiscalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroCupom As Long
Private mData As Date
Private mOrdem As Integer
Private mCodigoProduto As Long
Private mValorUnitario As Currency
Private mQuantidade As Currency
Private mValorTotal As Currency
Private mItemCancelado As Boolean
Private mValorDesconto As Currency
Private mValorAcrescimo As Currency
Private mDescontoEmbutido As Boolean
Private mPeriodo As Integer
Private mTipoCombustivel As String
Private mCodigoECF As Integer
Private mCodigoAliquota As Integer
Private mCodigoGrupo As Integer

Private rsMovCupomFiscalItem As New adodb.Recordset
Private IntegracaoNuvem As New cIntegracaoNuvem


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
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


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get ValorUnitario() As Currency
    ValorUnitario = mValorUnitario
End Property
Public Property Let ValorUnitario(ByVal Valor As Currency)
    mValorUnitario = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get ItemCancelado() As Boolean
    ItemCancelado = mItemCancelado
End Property
Public Property Let ItemCancelado(ByVal Valor As Boolean)
    mItemCancelado = Valor
End Property


Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property


Public Property Get ValorAcrescimo() As Currency
    ValorAcrescimo = mValorAcrescimo
End Property
Public Property Let ValorAcrescimo(ByVal Valor As Currency)
    mValorAcrescimo = Valor
End Property


Public Property Get DescontoEmbutido() As Boolean
    DescontoEmbutido = mDescontoEmbutido
End Property
Public Property Let DescontoEmbutido(ByVal Valor As Boolean)
    mDescontoEmbutido = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property

Public Property Get CodigoECF() As Integer
    CodigoECF = mCodigoECF
End Property
Public Property Let CodigoECF(ByVal Valor As Integer)
    mCodigoECF = Valor
End Property


Public Property Get CodigoAliquota() As Integer
    CodigoAliquota = mCodigoAliquota
End Property
Public Property Let CodigoAliquota(ByVal Valor As Integer)
    mCodigoAliquota = Valor
End Property


Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Ecf] = " & pCodigoECF
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Numero do Cupom] = " & pNumeroCupom
    xCondicao = xCondicao & " AND Ordem = " & pOrdem
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Cupom_Fiscal_Item ( Empresa, [Numero do Cupom], Data, Ordem, [Codigo do Produto], "
    gSQL = gSQL & "[Valor Unitario], Quantidade, [Valor Total], [Item Cancelado], [Valor do Desconto], "
    gSQL = gSQL & "[Valor do Acrescimo], [Desconto Embutido], Periodo, [Tipo de Combustivel], "
    gSQL = gSQL & "[Codigo da Ecf], [Codigo da Aliquota], [Codigo do Grupo] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlBoolean(1, mItemCancelado, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlValor(1, mValorAcrescimo, ", ")
    Call sqlBoolean(1, mDescontoEmbutido, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlNumero(1, mCodigoECF, ", ")
    Call sqlNumero(1, mCodigoAliquota, ", ")
    Call sqlNumero(1, mCodigoGrupo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
        'Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mCodigoECF, mNumeroCupom, mData, mOrdem)
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlBoolean(2, ", [Item Cancelado] = ", mItemCancelado)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlValor(2, ", [Valor do Acrescimo] = ", mValorAcrescimo)
    Call sqlBoolean(2, ", [Desconto Embutido] = ", mDescontoEmbutido)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlNumero(2, ", [Codigo da Ecf] = ", mCodigoECF)
    Call sqlNumero(2, ", [Codigo da Aliquota] = ", mCodigoAliquota)
    Call sqlNumero(2, ", [Codigo do Grupo] = ", mCodigoGrupo)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
        'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarDesconto(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pTotalCupom As Currency, ByVal pValorDesconto As Currency) As Boolean
Dim lRecordsAffected As Long
Dim xSomaDesconto As Currency
Dim xPercentual As Currency
Dim xValor As Currency
Dim xOrdem As Integer

On Error GoTo Err_Alterar

    AlterarDesconto = False
    xSomaDesconto = 0
    xPercentual = Format(pValorDesconto * 100 / pTotalCupom, "0000000000.0000")
    gSQL = "SELECT [Valor do Desconto], [Valor Total], Ordem"
    gSQL = gSQL & "  FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & "   AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    gSQL = gSQL & " ORDER BY Ordem"
    If Localizar(55) Then
        If rsMovCupomFiscalItem.RecordCount > 1 Then
            Do Until rsMovCupomFiscalItem.EOF
                xOrdem = rsMovCupomFiscalItem("Ordem").Value
                xValor = Format(rsMovCupomFiscalItem("Valor Total").Value * xPercentual / 100, "0000000000.00")
                xSomaDesconto = xSomaDesconto + xValor
                gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
                Call sqlValor(2, "[Valor do Desconto] = ", xValor)
                gSQL = gSQL & " WHERE Empresa = " & pEmpresa
                gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
                gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
                gSQL = gSQL & " AND Data = " & preparaData(pData)
                gSQL = gSQL & " AND Ordem = " & rsMovCupomFiscalItem("Ordem").Value
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                    'Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, rsMovCupomFiscalItem("Ordem").Value)
                End If
                rsMovCupomFiscalItem.MoveNext
            Loop
            If pValorDesconto <> xSomaDesconto Then
                If pValorDesconto < xSomaDesconto Then
                    xValor = xValor - (xSomaDesconto - pValorDesconto)
                Else
                    xValor = xValor + (pValorDesconto - xSomaDesconto)
                End If
                gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
                Call sqlValor(2, "[Valor do Desconto] = ", xValor)
                gSQL = gSQL & " WHERE Empresa = " & pEmpresa
                gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
                gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
                gSQL = gSQL & " AND Data = " & preparaData(pData)
                gSQL = gSQL & " AND Ordem = " & xOrdem
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                End If
            End If
        Else
            gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
            Call sqlValor(2, "[Valor do Desconto] = ", pValorDesconto)
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
            gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
            gSQL = gSQL & " AND Data = " & preparaData(pData)
            gSQL = gSQL & " AND Ordem = " & rsMovCupomFiscalItem("Ordem").Value
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected > 0 Then
                AlterarDesconto = True
            End If
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
        'Call IncluiIntegracaoNuvem("EXCLUIR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_Excluir:
End Function
Public Function ExcluirNumeroData(ByVal pEmpresa As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirNumeroData = False
    gSQL = "DELETE FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirNumeroData = True
    End If
    Exit Function

Err_Excluir:
End Function



Public Function CancelaCupom(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_CancelaCupom

    CancelaCupom = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
    Call sqlBoolean(2, "[Item Cancelado] = ", "" & "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaCupom = True
        'Call IncluiIntegracaoNuvem("CANCELAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, 0)
    End If
    Exit Function

Err_CancelaCupom:
End Function


Public Function CancelaItem(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long, ByVal pOrdem As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_CancelaItem

    CancelaItem = False
    gSQL = "UPDATE Movimento_Cupom_Fiscal_Item SET "
    Call sqlBoolean(2, "[Item Cancelado] = ", "" & "True")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Ecf] = " & pCodigoECF
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gSQL = gSQL & " AND Ordem = " & pOrdem
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CancelaItem = True
        'Call IncluiIntegracaoNuvem("CANCELAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function

Err_CancelaItem:
End Function


Public Function TotalAcrescimo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pDescontoEmbutido As Boolean) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalAcrescimo = 0
    gSQL = "SELECT SUM([Valor do Acrescimo]) AS Total FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    If pDescontoEmbutido = True Then
        gSQL = gSQL & " AND [Desconto Embutido] = " & preparaBooleano(True)
    Else
        gSQL = gSQL & " AND [Desconto Embutido] = " & preparaBooleano(False)
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCupomFiscalItem("Total").Value) Then
            TotalAcrescimo = rsMovCupomFiscalItem("Total").Value
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalDesconto(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoProduto As Long, ByVal pDescontoEmbutido As Boolean) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalDesconto = 0
    gSQL = "SELECT SUM([Valor do Desconto]) AS Total FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    If pDescontoEmbutido = True Then
        gSQL = gSQL & " AND [Desconto Embutido] = " & preparaBooleano(True)
    Else
        gSQL = gSQL & " AND [Desconto Embutido] = " & preparaBooleano(False)
    End If
    If pCodigoProduto > 0 Then
        gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCupomFiscalItem("Total").Value) Then
            TotalDesconto = rsMovCupomFiscalItem("Total").Value
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalCancelamento(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoECF As Integer) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    
    TotalCancelamento = 0
    gSQL = "SELECT SUM([Valor Total]) AS Total FROM Movimento_Cupom_Fiscal_Item"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Item Cancelado] = " & preparaBooleano(True)
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    If pCodigoProduto > 0 Then
        gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    End If
    If pCodigoECF > 0 Then
        gSQL = gSQL & "    AND [Codigo da ECF] = " & pCodigoECF
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCupomFiscalItem("Total").Value) Then
            TotalCancelamento = rsMovCupomFiscalItem("Total").Value
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotValorProdutosSubstVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoECF As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotValorProdutosSubstVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal_Item, Aliquota"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Desconto Embutido] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto("  ")
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal_Item.[Codigo da Aliquota] = Aliquota.Codigo"
    xSQL = xSQL & "    AND Aliquota.[Codigo Fiscal] = " & preparaTexto("FF")
    If pCodigoECF > 0 Then
        xSQL = xSQL & "    AND Movimento_Cupom_Fiscal_Item.[Codigo da ECF] = " & pCodigoECF
    End If
    
    Set rsMovCupomFiscalItem = New adodb.Recordset
    rsMovCupomFiscalItem.CursorLocation = adUseClient
    rsMovCupomFiscalItem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscalItem.EOF Then
        If Not IsNull(rsMovCupomFiscalItem("TotalVenda").Value) Then
            TotValorProdutosSubstVendaData = rsMovCupomFiscalItem("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscalItem("TotalDesconto").Value) Then
            TotValorProdutosSubstVendaData = TotValorProdutosSubstVendaData - rsMovCupomFiscalItem("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotValorProdutosTribVendaData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pCodigoECF As Integer, ByVal pAliquota As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotValorProdutosTribVendaData = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Valor Total]) AS TotalVenda, Sum([Valor do Desconto]) AS TotalDesconto"
    xSQL = xSQL & "   FROM Movimento_Cupom_Fiscal_Item, Aliquota"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    xSQL = xSQL & "    AND Periodo >= " & pPeriodoInicial
    xSQL = xSQL & "    AND Periodo <= " & pPeriodoFinal
    xSQL = xSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Desconto Embutido] = " & preparaBooleano(False)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto("  ")
    xSQL = xSQL & "    AND Movimento_Cupom_Fiscal_Item.[Codigo da Aliquota] = Aliquota.Codigo"
    xSQL = xSQL & "    AND Aliquota.[Aliquota do Imposto] = " & preparaValor(pAliquota)
    If pCodigoECF > 0 Then
        xSQL = xSQL & "    AND Movimento_Cupom_Fiscal_Item.[Codigo da ECF] = " & pCodigoECF
    End If
    
    Set rsMovCupomFiscalItem = New adodb.Recordset
    rsMovCupomFiscalItem.CursorLocation = adUseClient
    rsMovCupomFiscalItem.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscalItem.EOF Then
        If Not IsNull(rsMovCupomFiscalItem("TotalVenda").Value) Then
            TotValorProdutosTribVendaData = rsMovCupomFiscalItem("TotalVenda").Value
        End If
        If Not IsNull(rsMovCupomFiscalItem("TotalDesconto").Value) Then
            TotValorProdutosTribVendaData = TotValorProdutosTribVendaData - rsMovCupomFiscalItem("TotalDesconto").Value
        End If
    End If
    rsMovCupomFiscalItem.Close
    Set rsMovCupomFiscalItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Cupom], Data, Ordem, [Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Item Cancelado], [Valor do Desconto], [Valor do Acrescimo], [Desconto Embutido], Periodo, [Tipo de Combustivel], [Codigo da Ecf], [Codigo da Aliquota], [Codigo do Grupo] FROM Movimento_Cupom_Fiscal_Item" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCupomFiscalItem = New adodb.Recordset
    rsMovCupomFiscalItem.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCupomFiscalItem.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCupomFiscalItem.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCupomFiscalItem("Empresa").Value
    NumeroCupom = rsMovCupomFiscalItem("Numero do Cupom").Value
    Data = rsMovCupomFiscalItem("Data").Value
    Ordem = rsMovCupomFiscalItem("Ordem").Value
    CodigoProduto = rsMovCupomFiscalItem("Codigo do Produto").Value
    ValorUnitario = rsMovCupomFiscalItem("Valor Unitario").Value
    Quantidade = rsMovCupomFiscalItem("Quantidade").Value
    ValorTotal = rsMovCupomFiscalItem("Valor Total").Value
    ItemCancelado = rsMovCupomFiscalItem("Item Cancelado").Value
    ValorDesconto = rsMovCupomFiscalItem("Valor do Desconto").Value
    ValorAcrescimo = rsMovCupomFiscalItem("Valor do Acrescimo").Value
    DescontoEmbutido = rsMovCupomFiscalItem("Desconto Embutido").Value
    Periodo = rsMovCupomFiscalItem("Periodo").Value
    TipoCombustivel = rsMovCupomFiscalItem("Tipo de Combustivel").Value
    CodigoECF = rsMovCupomFiscalItem("Codigo da Ecf").Value
    CodigoAliquota = rsMovCupomFiscalItem("Codigo da Aliquota").Value
    CodigoGrupo = rsMovCupomFiscalItem("Codigo do Grupo").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Private Sub IncluiIntegracaoNuvem(ByVal pTipoOperacao As String, ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer)
    IntegracaoNuvem.Empresa = mEmpresa
    IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
    IntegracaoNuvem.NomeTabela = "MovimentoCupomFiscalItem"
    IntegracaoNuvem.ChaveAcesso = pEmpresa & "|@|" & pCodigoECF & "|@|" & pNumeroCupom & "|@|" & pData & "|@|" & pOrdem & "|@|"
    IntegracaoNuvem.TipoOperacao = pTipoOperacao
    IntegracaoNuvem.IntegradoEm = CDate("00:00:00")
    If Not IntegracaoNuvem.Incluir Then
        Call CriaLogCadastroDll2("cMovimentoCupomFiscalItem:IncluiIntegracaoNuvem - Erro não identificado.", Err.Description, "Tabela=" & IntegracaoNuvem.NomeTabela & " - Chv.Acesso=" & IntegracaoNuvem.ChaveAcesso)
    End If
End Sub

