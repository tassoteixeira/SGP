VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBicoAutomacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoBico As Integer
Private mCodigoAutomacao As String

Private rsBicoAutomacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoBico() As Integer
    CodigoBico = mCodigoBico
End Property
Public Property Let CodigoBico(ByVal Valor As Integer)
    mCodigoBico = Valor
End Property


Public Property Get CodigoAutomacao() As String
    CodigoAutomacao = mCodigoAutomacao
End Property
Public Property Let CodigoAutomacao(ByVal Valor As String)
    mCodigoAutomacao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigoBico(ByVal pCodigoBico As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigoBico = False
    xCondicao = " WHERE [Codigo do Bico] = " & pCodigoBico
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCodigoBico = True
        AtribuiValor
    End If
    rsBicoAutomacao.Close
    Set rsBicoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigoAutomacao(ByVal pCodigoAutomacao As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarCodigoAutomacao = False
    gSQL = PreparaSQL(" WHERE [Codigo da Automacao] = " & preparaTexto(pCodigoAutomacao), "")
    If Localizar(1) Then
        LocalizarCodigoAutomacao = True
        AtribuiValor
    End If
    rsBicoAutomacao.Close
    Set rsBicoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO BicoAutomacao ( [Codigo do Bico], [Codigo da Automacao] ) VALUES ( "
    Call sqlNumero(1, mCodigoBico, ", ")
    Call sqlTexto(1, mCodigoAutomacao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigoBico As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE BicoAutomacao SET "
    Call sqlNumero(2, "[Codigo do Bico] = ", mCodigoBico)
    Call sqlTexto(2, ", [Codigo da Automacao] = ", mCodigoAutomacao)
    gSQL = gSQL & " WHERE [Codigo do Bico] = " & pCodigoBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function ConverteAutomacaoBico(ByVal pCodigoAutomacao As String) As Integer
Dim xSQL As String
On Error GoTo trata_erro

    ConverteAutomacaoBico = 0
    gSQL = ""
    gSQL = gSQL & "SELECT [Codigo do Bico]"
    gSQL = gSQL & "  FROM BicoAutomacao"
    gSQL = gSQL & " WHERE [Codigo da Automacao] = " & preparaTexto(pCodigoAutomacao)
    If Localizar(1) Then
        ConverteAutomacaoBico = rsBicoAutomacao("Codigo do Bico").Value
    End If
    rsBicoAutomacao.Close
    Set rsBicoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ConverteBicoAutomacao(ByVal pCodigoBico As Integer) As String
Dim xSQL As String
On Error GoTo trata_erro

    ConverteBicoAutomacao = "00"
    gSQL = ""
    gSQL = gSQL & "SELECT [Codigo da Automacao]"
    gSQL = gSQL & "  FROM BicoAutomacao"
    gSQL = gSQL & " WHERE [Codigo do Bico] = " & pCodigoBico
    If Localizar(1) Then
        ConverteBicoAutomacao = rsBicoAutomacao("Codigo da Automacao").Value
    End If
    rsBicoAutomacao.Close
    Set rsBicoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Excluir(ByVal pCodigoBico As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM BicoAutomacao"
    gSQL = gSQL & " WHERE [Codigo do Bico] = " & pCodigoBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Bico], [Codigo da Automacao] FROM BicoAutomacao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBicoAutomacao = New adodb.Recordset
    rsBicoAutomacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBicoAutomacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBicoAutomacao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoBico = rsBicoAutomacao("Codigo do Bico").Value
    CodigoAutomacao = rsBicoAutomacao("Codigo da Automacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

