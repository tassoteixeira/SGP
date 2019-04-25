VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstoque2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mGrupoProduto As Integer
Private mCodigoProduto2 As Long
Private mQuantidade As Currency
Private mPrecoCusto As Currency
Private mPrecoVenda As Currency
Private mPrecoCustoMedio As Currency

Private rsEstoque2 As New adodb.Recordset


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


Public Property Get GrupoProduto() As Integer
    GrupoProduto = mGrupoProduto
End Property
Public Property Let GrupoProduto(ByVal Valor As Integer)
    mGrupoProduto = Valor
End Property


Public Property Get CodigoProduto2() As Long
    CodigoProduto2 = mCodigoProduto2
End Property
Public Property Let CodigoProduto2(ByVal Valor As Long)
    mCodigoProduto2 = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get PrecoCustoMedio() As Currency
    PrecoCustoMedio = mPrecoCustoMedio
End Property
Public Property Let PrecoCustoMedio(ByVal Valor As Currency)
    mPrecoCustoMedio = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoProduto2 As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsEstoque2.Close
    Set rsEstoque2 = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ExisteEstoqueData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoGrupo As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteEstoqueData = False
    gSQL = ""
    gSQL = gSQL & "SELECT [Codigo do Produto2]"
    gSQL = gSQL & "  FROM Estoque2"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    If pCodigoGrupo > 0 Then
        gSQL = gSQL & " AND [Grupo do Produto] = " & pCodigoGrupo
    End If
    If pCodigoProduto > 0 Then
        gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    End If
    If Localizar(1) = True Then
        ExisteEstoqueData = True
    End If
    rsEstoque2.Close
    Set rsEstoque2 = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Estoque2 ( Empresa, Data, [Grupo do Produto], [Codigo do Produto2], Quantidade, "
    gSQL = gSQL & "[Preco de Custo], [Preco de Venda], [Preco de Custo Medio] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mGrupoProduto, ", ")
    Call sqlNumero(1, mCodigoProduto2, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mPrecoCustoMedio, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoProduto2 As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Estoque2 SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Grupo do Produto] = ", mGrupoProduto)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto2)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlValor(2, ", [Preco de Custo Medio] = ", mPrecoCustoMedio)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoProduto2 As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Estoque2"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluirData(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoGrupo As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirData = False
    gSQL = "DELETE FROM Estoque2"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    If pCodigoGrupo > 0 Then
        gSQL = gSQL & " AND [Grupo do Produto] = " & pCodigoGrupo
    End If
    If pCodigoProduto > 0 Then
        gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    End If
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirData = True
    End If
    Exit Function

Err_Excluir:
End Function

'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Grupo do Produto], [Codigo do Produto2], Quantidade, [Preco de Custo], [Preco de Venda], [Preco de Custo Medio] FROM Estoque2" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEstoque2 = New adodb.Recordset
    rsEstoque2.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEstoque2.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEstoque2.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEstoque2("Empresa").Value
    Data = rsEstoque2("Data").Value
    GrupoProduto = rsEstoque2("Grupo do Produto").Value
    CodigoProduto2 = rsEstoque2("Codigo do Produto2").Value
    Quantidade = rsEstoque2("Quantidade").Value
    PrecoCusto = rsEstoque2("Preco de Custo").Value
    PrecoVenda = rsEstoque2("Preco de Venda").Value
    PrecoCustoMedio = rsEstoque2("Preco de Custo Medio").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

