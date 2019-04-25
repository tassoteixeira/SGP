VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovDescontoPersonalizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoCliente As Long
Private mCodigoProduto As Long
Private mPercentualaDescontar As Currency
Private mValoraDescontar As Currency
Private mDesconto As Boolean
Private mPrecoFixo As Currency
Private mPrecoParaECF As Currency

Private rsMovDescontoPersonalizado As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get PercentualaDescontar() As Currency
    PercentualaDescontar = mPercentualaDescontar
End Property
Public Property Let PercentualaDescontar(ByVal Valor As Currency)
    mPercentualaDescontar = Valor
End Property


Public Property Get ValoraDescontar() As Currency
    ValoraDescontar = mValoraDescontar
End Property
Public Property Let ValoraDescontar(ByVal Valor As Currency)
    mValoraDescontar = Valor
End Property


Public Property Get Desconto() As Boolean
    Desconto = mDesconto
End Property
Public Property Let Desconto(ByVal Valor As Boolean)
    mDesconto = Valor
End Property


Public Property Get PrecoFixo() As Currency
    PrecoFixo = mPrecoFixo
End Property
Public Property Let PrecoFixo(ByVal Valor As Currency)
    mPrecoFixo = Valor
End Property


Public Property Get PrecoParaECF() As Currency
    PrecoParaECF = mPrecoParaECF
End Property
Public Property Let PrecoParaECF(ByVal Valor As Currency)
    mPrecoParaECF = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigoCliente As Long, ByVal nCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & nCodigoProduto
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCliente(ByVal nCodigoCliente As Long, ByVal nCodigoProduto As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCliente = False
    xCondicao = " WHERE [Codigo do Cliente] = " & nCodigoCliente
    xCondicao = xCondicao & " AND [Codigo do Produto] > " & nCodigoProduto
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Produto] ASC")
    LocalizarCliente = Localizar(1)
    If LocalizarCliente Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
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
    xCondicao = " WHERE [Codigo do Cliente] <= " & mCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] DESC, [Codigo do Produto] DESC")
    If Localizar(20) Then
        If rsMovDescontoPersonalizado.RecordCount > 1 Then
            Do Until rsMovDescontoPersonalizado.EOF
                If xRegIgual = False Then
                    If rsMovDescontoPersonalizado("Codigo do Cliente").Value = mCodigoCliente And rsMovDescontoPersonalizado("Codigo do Produto").Value = mCodigoProduto Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovDescontoPersonalizado.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo do Cliente] > " & 0
    xCondicao = xCondicao & " AND [Codigo do Produto] > " & 0
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] ASC, [Codigo do Produto] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
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
    xCondicao = " WHERE [Codigo do Cliente] >= " & mCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] ASC, [Codigo do Produto] ASC")
    If Localizar(20) Then
        If rsMovDescontoPersonalizado.RecordCount > 1 Then
            Do Until rsMovDescontoPersonalizado.EOF
                If xRegIgual = False Then
                    If rsMovDescontoPersonalizado("Codigo do Cliente").Value = mCodigoCliente And rsMovDescontoPersonalizado("Codigo do Produto").Value = mCodigoProduto Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovDescontoPersonalizado.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Codigo do Cliente] < " & 999999999
    xCondicao = xCondicao & " AND [Codigo do Produto] < " & 999999999
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] DESC, [Codigo do Produto] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovDescontoPersonalizado.Close
    Set rsMovDescontoPersonalizado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Desconto_Personalizado ( [Codigo do Cliente], [Codigo do Produto], "
    gSQL = gSQL & " [Percentual a Descontar], [Valor a Descontar], Desconto, [Preco Fixo], "
    gSQL = gSQL & " [Preco para ECF] ) VALUES ( "
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mPercentualaDescontar, ", ")
    Call sqlValor(1, mValoraDescontar, ", ")
    Call sqlBoolean(1, mDesconto, ", ")
    Call sqlValor(1, mPrecoFixo, ", ")
    Call sqlValor(1, mPrecoParaECF, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigoCliente As Long, ByVal nCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Desconto_Personalizado SET "
    Call sqlNumero(2, "[Codigo do Cliente] = ", mCodigoCliente)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Percentual a Descontar] = ", mPercentualaDescontar)
    Call sqlValor(2, ", [Valor a Descontar] = ", mValoraDescontar)
    Call sqlBoolean(2, ", Desconto = ", mDesconto)
    Call sqlValor(2, ", [Preco Fixo] = ", mPrecoFixo)
    Call sqlValor(2, ", [Preco para ECF] = ", mPrecoParaECF)
    gSQL = gSQL & " WHERE [Codigo do Cliente] = " & nCodigoCliente
    gSQL = gSQL & " AND [Codigo do Produto] = " & nCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigoCliente As Long, ByVal nCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Desconto_Personalizado"
    gSQL = gSQL & " WHERE [Codigo do Cliente] = " & nCodigoCliente
    gSQL = gSQL & " AND [Codigo do Produto] = " & nCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Cliente], [Codigo do Produto], [Percentual a Descontar], [Valor a Descontar], Desconto, [Preco Fixo], [Preco para ECF] FROM Movimento_Desconto_Personalizado" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovDescontoPersonalizado = New adodb.Recordset
    rsMovDescontoPersonalizado.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovDescontoPersonalizado.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovDescontoPersonalizado.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoCliente = rsMovDescontoPersonalizado("Codigo do Cliente").Value
    CodigoProduto = rsMovDescontoPersonalizado("Codigo do Produto").Value
    PercentualaDescontar = rsMovDescontoPersonalizado("Percentual a Descontar").Value
    ValoraDescontar = rsMovDescontoPersonalizado("Valor a Descontar").Value
    Desconto = rsMovDescontoPersonalizado("Desconto").Value
    PrecoFixo = rsMovDescontoPersonalizado("Preco Fixo").Value
    PrecoParaECF = rsMovDescontoPersonalizado("Preco para ECF").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

