VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClienteConveniado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoConvenio As Integer
Private mCodigoConveniado As Long
Private mNome As String

Private rsClienteConveniado As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoConvenio() As Integer
    CodigoConvenio = mCodigoConvenio
End Property
Public Property Let CodigoConvenio(ByVal Valor As Integer)
    mCodigoConvenio = Valor
End Property


Public Property Get CodigoConveniado() As Long
    CodigoConveniado = mCodigoConveniado
End Property
Public Property Let CodigoConveniado(ByVal Valor As Long)
    mCodigoConveniado = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigoConvenio As Integer, ByVal nCodigoConveniado As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Codigo do Convenio] = " & nCodigoConvenio
    xCondicao = xCondicao & " AND [Codigo do Conveniado] = " & nCodigoConveniado
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        AtribuiValor
        LocalizarCodigo = True
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
    Dim xCondicao As String
    Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE [Codigo do Convenio] <= " & mCodigoConvenio
    xCondicao = xCondicao & " AND [Codigo do Conveniado] <= " & mCodigoConveniado
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Convenio] DESC, [Codigo do Conveniado] DESC")
    If Localizar(0) Then
        If rsClienteConveniado.RecordCount > 1 Then
            Do Until rsClienteConveniado.EOF
                If xRegIgual = False Then
                    If rsClienteConveniado("Codigo do Convenio").Value = mCodigoConvenio And rsClienteConveniado("Codigo do Conveniado").Value = mCodigoConveniado Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsClienteConveniado.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nCodigoConvenio As Integer, ByVal nNome As String) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE [Codigo do Convenio] = " & nCodigoConvenio
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(nNome)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        AtribuiValor
        LocalizarNome = True
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo do Convenio] >= 0"
    xCondicao = xCondicao & " AND [Codigo do Conveniado] >= 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Convenio] ASC, [Codigo do Conveniado] ASC")
    If Localizar(1) Then
        AtribuiValor
        LocalizarPrimeiro = True
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
    Dim xCondicao As String
    Dim xRegIgual As Boolean
On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE [Codigo do Convenio] >= " & mCodigoConvenio
    xCondicao = xCondicao & " AND [Codigo do Conveniado] >= " & mCodigoConveniado
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Convenio] ASC, [Codigo do Conveniado] ASC")
    If Localizar(0) Then
        If rsClienteConveniado.RecordCount > 1 Then
            Do Until rsClienteConveniado.EOF
                If xRegIgual = False Then
                    If rsClienteConveniado("Codigo do Convenio").Value = mCodigoConvenio And rsClienteConveniado("Codigo do Conveniado").Value = mCodigoConveniado Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsClienteConveniado.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximoCodigo(ByVal nCodigoConvenio As Integer) As Long
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximoCodigo = 1
    xCondicao = " WHERE [Codigo do Convenio] = " & nCodigoConvenio
    xCondicao = xCondicao & " AND [Codigo do Conveniado] < " & 99999999
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Convenio] DESC, [Codigo do Conveniado] DESC")
    If Localizar(1) Then
        LocalizarProximoCodigo = rsClienteConveniado("Codigo do Conveniado").Value + 1
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Codigo do Convenio] < " & 9999
    xCondicao = xCondicao & " AND [Codigo do Conveniado] < " & 99999999
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Convenio] DESC, [Codigo do Conveniado] DESC")
    If Localizar(1) Then
        AtribuiValor
        LocalizarUltimo = True
    End If
    rsClienteConveniado.Close
    Set rsClienteConveniado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Cliente_Conveniado ( [Codigo do Convenio], [Codigo do Conveniado], Nome ) VALUES ( "
    Call sqlNumero(1, mCodigoConvenio, ", ")
    Call sqlNumero(1, mCodigoConveniado, ", ")
    Call sqlTexto(1, mNome, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigoConvenio As Integer, ByVal nCodigoConveniado As Long) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Cliente_Conveniado SET "
    Call sqlNumero(2, "[Codigo do Convenio] = ", mCodigoConvenio)
    Call sqlNumero(2, ", [Codigo do Conveniado] = ", mCodigoConveniado)
    Call sqlTexto(2, ", Nome = ", mNome)
    gSQL = gSQL & " WHERE [Codigo do Convenio] = " & nCodigoConvenio
    gSQL = gSQL & " AND [Codigo do Conveniado] = " & nCodigoConveniado
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigoConvenio As Integer, ByVal nCodigoConveniado As Long) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Cliente_Conveniado"
    gSQL = gSQL & " WHERE [Codigo do Convenio] = " & nCodigoConvenio
    gSQL = gSQL & " AND [Codigo do Conveniado] = " & nCodigoConveniado
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Convenio], [Codigo do Conveniado], Nome FROM Cliente_Conveniado" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsClienteConveniado = New adodb.Recordset
    rsClienteConveniado.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsClienteConveniado.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsClienteConveniado.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoConvenio = rsClienteConveniado("Codigo do Convenio").Value
    CodigoConveniado = rsClienteConveniado("Codigo do Conveniado").Value
    Nome = rsClienteConveniado("Nome").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

