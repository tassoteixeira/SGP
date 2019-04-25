VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrupoTipoMovimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mTipoMovimento As Integer
Private mCodigoGrupo As Integer

Private rsGrupoTipoMovCaixa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pTipoMovimento As Integer, ByVal pCodigoGrupo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Grupo] = " & pCodigoGrupo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsGrupoTipoMovCaixa.Close
    Set rsGrupoTipoMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarGrupo(ByVal pCodigoGrupo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarGrupo = False
    xCondicao = " WHERE [Codigo do Grupo] = " & pCodigoGrupo
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarGrupo = True
        AtribuiValor
    End If
    rsGrupoTipoMovCaixa.Close
    Set rsGrupoTipoMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY [Tipo do Movimento] ASC, [Codigo do Grupo] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsGrupoTipoMovCaixa.Close
    Set rsGrupoTipoMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY [Tipo do Movimento] DESC, [Codigo do Grupo] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsGrupoTipoMovCaixa.Close
    Set rsGrupoTipoMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO GrupoTipoMovimentoCaixa ( [Tipo do Movimento], [Codigo do Grupo] ) VALUES ( "
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoGrupo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pTipoMovimento As Integer, ByVal pCodigoGrupo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE GrupoTipoMovimentoCaixa SET "
    Call sqlNumero(2, "[Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Grupo] = ", mCodigoGrupo)
    gSQL = gSQL & " WHERE [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Grupo] = " & pCodigoGrupo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pTipoMovimento As Integer, ByVal pCodigoGrupo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM GrupoTipoMovimentoCaixa"
    gSQL = gSQL & " WHERE [Tipo do Movimento] = " & pTipoMovimento
    gSQL = gSQL & " AND [Codigo do Grupo] = " & pCodigoGrupo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Tipo do Movimento], [Codigo do Grupo] FROM GrupoTipoMovimentoCaixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsGrupoTipoMovCaixa = New adodb.Recordset
    rsGrupoTipoMovCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsGrupoTipoMovCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsGrupoTipoMovCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    TipoMovimento = rsGrupoTipoMovCaixa("Tipo do Movimento").Value
    CodigoGrupo = rsGrupoTipoMovCaixa("Codigo do Grupo").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

