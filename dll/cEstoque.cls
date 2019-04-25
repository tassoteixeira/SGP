VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mGrupoProduto As Integer
Private mCodigoProduto2 As Long
Private mQuantidade As Currency
Private mPrecoVenda As Currency
Private mQuantidadeTemporaria As Currency

Private rsEstoque As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
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


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get QuantidadeTemporaria() As Currency
    QuantidadeTemporaria = mQuantidadeTemporaria
End Property
Public Property Let QuantidadeTemporaria(ByVal Valor As Currency)
    mQuantidadeTemporaria = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigoProduto2 As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto2] = " & nCodigoProduto2
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsEstoque.Close
    Set rsEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Estoque ( Empresa, [Grupo do Produto], [Codigo do Produto2], Quantidade, [Preco de Venda], [Quantidade Temporaria] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mGrupoProduto, ", ")
    Call sqlNumero(1, mCodigoProduto2, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mQuantidadeTemporaria, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Adicionar(ByVal nEmpresa As Integer, ByVal nCodigoProduto2 As Long, ByVal nQuantidade As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Adicionar = False
    gSQL = "UPDATE Estoque SET "
    Call sqlValor(2, "Quantidade = Quantidade + ", nQuantidade)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & nCodigoProduto2
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Adicionar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Estoque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Grupo do Produto] = ", mGrupoProduto)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto2)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlValor(2, ", [Quantidade Temporaria] = ", mQuantidadeTemporaria)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Call CriaLogCadastroDll2("cEstoque:Alterar - ** Investigação Especial - Registro alterado p/:", "Investigação Especial", gSQL)
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarQuantidade(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency, ByVal pAdiciona As Boolean) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarQuantidade

    AlterarQuantidade = False
    gSQL = "UPDATE Estoque SET "
    If pAdiciona Then
        Call sqlValor(2, "Quantidade = Quantidade + ", pQuantidade)
    Else
        Call sqlValor(2, "Quantidade = Quantidade - ", pQuantidade)
    End If
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarQuantidade = True
    End If
    Exit Function

Err_AlterarQuantidade:
End Function
Public Function AlterarQuantidadeTemporaria(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pQuantidadeTemporaria As Currency, ByVal pAdiciona As Boolean) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarQuantidadeTemporaria

    AlterarQuantidadeTemporaria = False
    gSQL = "UPDATE Estoque SET "
    If pAdiciona Then
        Call sqlValor(2, "[Quantidade Temporaria] = [Quantidade Temporaria] + ", pQuantidadeTemporaria)
    Else
        Call sqlValor(2, "[Quantidade Temporaria] = [Quantidade Temporaria] - ", pQuantidadeTemporaria)
    End If
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarQuantidadeTemporaria = True
    End If
    Exit Function

Err_AlterarQuantidadeTemporaria:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigoProduto2 As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Estoque"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & nCodigoProduto2
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Grupo do Produto], [Codigo do Produto2], Quantidade, [Preco de Venda], [Quantidade Temporaria] FROM Estoque" & " " & xCondicao & " " & xTipoOrdem)
End Function


Public Function Subtrair(ByVal nEmpresa As Integer, ByVal nCodigoProduto2 As Long, ByVal nQuantidade As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Subtrair = False
    gSQL = "UPDATE Estoque SET "
    Call sqlValor(2, "Quantidade = Quantidade - ", nQuantidade)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Produto2] = " & nCodigoProduto2
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Subtrair = True
    End If
    Exit Function

Err_Alterar:
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEstoque = New adodb.Recordset
    rsEstoque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEstoque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEstoque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEstoque("Empresa").Value
    GrupoProduto = rsEstoque("Grupo do Produto").Value
    CodigoProduto2 = rsEstoque("Codigo do Produto2").Value
    Quantidade = rsEstoque("Quantidade").Value
    PrecoVenda = rsEstoque("Preco de Venda").Value
    QuantidadeTemporaria = rsEstoque("Quantidade Temporaria").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function TotalGrupoProduto(ByVal pEmpresa As Integer, ByVal pCodigoGrupo As Integer, ByVal pCodigoProduto As Long) As Currency

On Error GoTo trata_erro
    
    TotalGrupoProduto = 0
    gSQL = "SELECT Sum(Quantidade * [Preco de Venda]) As Total"
    gSQL = gSQL & "  FROM Estoque"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    Else
        gSQL = gSQL & " WHERE Empresa > 0"
    End If
    If pCodigoGrupo > 0 Then
        gSQL = gSQL & " AND [Grupo do Produto] = " & pCodigoGrupo
    End If
    If pCodigoProduto > 0 Then
        gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto
    End If
    If Localizar(0) Then
        If Not IsNull(rsEstoque("Total").Value) Then
            TotalGrupoProduto = rsEstoque("Total").Value
        End If
    End If
    rsEstoque.Close
    Set rsEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

