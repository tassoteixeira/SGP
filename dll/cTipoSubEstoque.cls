VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTipoSubEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String

Private rsTipoSubEstoque As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property
'Fim das Propriedades da Classe


'Inicio M�todos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsTipoSubEstoque.Close
    Set rsTipoSubEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO TipoSubEstoque ( Codigo, Nome ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE TipoSubEstoque SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM TipoSubEstoque"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Fun��es / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome FROM TipoSubEstoque" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTipoSubEstoque = New adodb.Recordset
    rsTipoSubEstoque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTipoSubEstoque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTipoSubEstoque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsTipoSubEstoque("Codigo").Value
    Nome = rsTipoSubEstoque("Nome").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

