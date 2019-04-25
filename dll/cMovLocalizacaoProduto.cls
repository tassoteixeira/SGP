VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovLocalizacaoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoLocalizacao As Integer
Private mCodigoProduto As Long

Private rsMovLocalizacaoProduto As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoLocalizacao() As Integer
    CodigoLocalizacao = mCodigoLocalizacao
End Property
Public Property Let CodigoLocalizacao(ByVal Valor As Integer)
    mCodigoLocalizacao = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoLocalizacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Localizacao] = " & pCodigoLocalizacao
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovLocalizacaoProduto.Close
    Set rsMovLocalizacaoProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo da Localizacao] < " & mCodigoLocalizacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Localizacao] DESC, [Codigo do Produto] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovLocalizacaoProduto.Close
    Set rsMovLocalizacaoProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Localizacao] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsMovLocalizacaoProduto.Close
    Set rsMovLocalizacaoProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo da Localizacao] > " & mCodigoLocalizacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Localizacao] ASC, [Codigo do Produto] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsMovLocalizacaoProduto.Close
    Set rsMovLocalizacaoProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Localizacao] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsMovLocalizacaoProduto.Close
    Set rsMovLocalizacaoProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoLocalizacaoProduto ( Empresa, [Codigo da Localizacao], [Codigo do Produto] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoLocalizacao, ", ")
    Call sqlNumero(1, mCodigoProduto, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoLocalizacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoLocalizacaoProduto SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo da Localizacao] = ", mCodigoLocalizacao)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Codigo da Localizacao] = " & pCodigoLocalizacao
    gSQL = gSQL & "   AND [Codigo do Produto] = " & pCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoLocalizacao As Integer, ByVal pCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoLocalizacaoProduto"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Codigo da Localizacao] = " & pCodigoLocalizacao
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
    PreparaSQL = Trim("SELECT Empresa, [Codigo da Localizacao], [Codigo do Produto] FROM MovimentoLocalizacaoProduto" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovLocalizacaoProduto = New adodb.Recordset
    rsMovLocalizacaoProduto.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovLocalizacaoProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovLocalizacaoProduto.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovLocalizacaoProduto("Empresa").Value
    CodigoLocalizacao = rsMovLocalizacaoProduto("Codigo da Localizacao").Value
    CodigoProduto = rsMovLocalizacaoProduto("Codigo do Produto").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

