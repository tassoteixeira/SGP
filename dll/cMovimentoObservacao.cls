VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoObservacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoObservacao As Integer
Private mCodigoProduto As Long
Private mObservacao As String

Private rsMovObservacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoObservacao() As Integer
    TipoObservacao = mTipoObservacao
End Property
Public Property Let TipoObservacao(ByVal Valor As Integer)
    mTipoObservacao = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pTipoObservacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Observacao] = " & pTipoObservacao
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCodigo = True
        AtribuiValor
    End If
    rsMovObservacao.Close
    Set rsMovObservacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Observacao] = " & mTipoObservacao
    xCondicao = xCondicao & " AND [Codigo do Produto] > " & mCodigoProduto
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto] DESC")
    If Localizar(1) Then
        LocalizarAnterior = True
        AtribuiValor
    End If
    rsMovObservacao.Close
    Set rsMovObservacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Observacao] = " & mTipoObservacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto] ASC")
    If Localizar(1) Then
        LocalizarPrimeiro = True
        AtribuiValor
    End If
    rsMovObservacao.Close
    Set rsMovObservacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Tipo de Observacao] = " & mTipoObservacao
    xCondicao = xCondicao & " AND [Codigo do Produto] < " & mCodigoProduto
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto] ASC")
    If Localizar(1) Then
        LocalizarProximo = True
        AtribuiValor
    End If
    rsMovObservacao.Close
    Set rsMovObservacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pTipoObservacao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Observacao] = " & pTipoObservacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto] DESC")
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsMovObservacao.Close
    Set rsMovObservacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoObservacao ( Empresa, [Tipo de Observacao], [Codigo do Produto], Observacao ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mTipoObservacao, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlTexto(1, mObservacao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pTipoObservacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoObservacao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Tipo de Observacao] = ", mTipoObservacao)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Tipo de Observacao] = " & pTipoObservacao
    gSQL = gSQL & "   AND [Codigo do Produto] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pTipoObservacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoObservacao"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Tipo de Observacao] = " & pTipoObservacao
    gSQL = gSQL & "   AND [Codigo do Produto] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Tipo de Observacao], [Codigo do Produto], Observacao FROM MovimentoObservacao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovObservacao = New adodb.Recordset
    rsMovObservacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovObservacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovObservacao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovObservacao("Empresa").Value
    TipoObservacao = rsMovObservacao("Tipo de Observacao").Value
    CodigoProduto = rsMovObservacao("Codigo do Produto").Value
    Observacao = rsMovObservacao("Observacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

