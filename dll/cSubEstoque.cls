VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoProduto As Long
Private mCodigoTipoSubEstoque As Integer
Private mQuantidade As Currency

Private rsSubEstoque As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get CodigoTipoSubEstoque() As Integer
    CodigoTipoSubEstoque = mCodigoTipoSubEstoque
End Property
Public Property Let CodigoTipoSubEstoque(ByVal Valor As Integer)
    mCodigoTipoSubEstoque = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] < " & mCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] < " & mCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Produto] DESC, [Codigo do Tipo do SubEstoque] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] > " & mCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] > " & mCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Codigo do Produto] ASC, [Codigo do Tipo do SubEstoque] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] > " & mCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] > " & mCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Codigo do Produto] ASC, [Codigo do Tipo do SubEstoque] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Produto] DESC, [Codigo do Tipo do SubEstoque] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsSubEstoque.Close
    Set rsSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO SubEstoque ( Empresa, [Codigo do Produto], [Codigo do Tipo do SubEstoque], Quantidade ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlNumero(1, mCodigoTipoSubEstoque, ", ")
    Call sqlValor(1, mQuantidade, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE SubEstoque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlNumero(2, ", [Codigo do Tipo do SubEstoque] = ", mCodigoTipoSubEstoque)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarQuantidade(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer, ByVal pQuantidade As Currency, ByVal pAdiciona As Boolean) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarQuantidade

    AlterarQuantidade = False
    gSQL = "UPDATE SubEstoque SET "
    If pAdiciona Then
        Call sqlValor(2, "Quantidade = Quantidade + ", pQuantidade)
    Else
        Call sqlValor(2, "Quantidade = Quantidade - ", pQuantidade)
    End If
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarQuantidade = True
    End If
    Exit Function

Err_AlterarQuantidade:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM SubEstoque"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Tipo do SubEstoque] = " & pCodigoTipoSubEstoque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Codigo do Produto], [Codigo do Tipo do SubEstoque], Quantidade FROM SubEstoque" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsSubEstoque = New adodb.Recordset
    rsSubEstoque.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSubEstoque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSubEstoque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsSubEstoque("Empresa").Value
    CodigoProduto = rsSubEstoque("Codigo do Produto").Value
    CodigoTipoSubEstoque = rsSubEstoque("Codigo do Tipo do SubEstoque").Value
    Quantidade = rsSubEstoque("Quantidade").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

