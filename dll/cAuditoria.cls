VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoUsuario As Integer
Private mData As Date
Private mHora As Date
Private mSequencia As Integer
Private mNomeInternoPrograma As String
Private mOperacao As Integer
Private mComputador As String
Private mObservacao As String

Private rsAuditoria As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Hora() As Date
    Hora = mHora
End Property
Public Property Let Hora(ByVal Valor As Date)
    mHora = Valor
End Property


Public Property Get Sequencia() As Integer
    Sequencia = mSequencia
End Property
Public Property Let Sequencia(ByVal Valor As Integer)
    mSequencia = Valor
End Property


Public Property Get NomeInternoPrograma() As String
    NomeInternoPrograma = mNomeInternoPrograma
End Property
Public Property Let NomeInternoPrograma(ByVal Valor As String)
    mNomeInternoPrograma = Valor
End Property


Public Property Get Operacao() As Integer
    Operacao = mOperacao
End Property
Public Property Let Operacao(ByVal Valor As Integer)
    mOperacao = Valor
End Property


Public Property Get Computador() As String
    Computador = mComputador
End Property
Public Property Let Computador(ByVal Valor As String)
    mComputador = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigoUsuario As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pSequencia As Integer, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Codigo do Usuario] = " & pCodigoUsuario
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Hora = " & preparaData(pHora)
    xCondicao = xCondicao & " AND Sequencia = " & pSequencia
    xCondicao = xCondicao & " AND [Nome Interno do Programa] = " & preparaTexto(pNomeInternoPrograma)
    xCondicao = xCondicao & " AND Operacao = " & pOperacao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE [Codigo do Usuario] < " & mCodigoUsuario
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    xCondicao = xCondicao & " AND Hora < " & preparaData(mHora)
    xCondicao = xCondicao & " AND Sequencia < " & mSequencia
    xCondicao = xCondicao & " AND [Nome Interno do Programa] < " & preparaTexto(mNomeInternoPrograma)
    xCondicao = xCondicao & " AND Operacao < " & mOperacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Usuario] DESC, Data DESC, Hora DESC, Sequencia DESC, [Nome Interno do Programa] DESC, Operacao DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo do Usuario] > " & mCodigoUsuario
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Hora > " & preparaData(mHora)
    xCondicao = xCondicao & " AND Sequencia > " & mSequencia
    xCondicao = xCondicao & " AND [Nome Interno do Programa] > " & preparaTexto(mNomeInternoPrograma)
    xCondicao = xCondicao & " AND Operacao > " & mOperacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Usuario] ASC, Data ASC, Hora ASC, Sequencia ASC, [Nome Interno do Programa] ASC, Operacao ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE [Codigo do Usuario] > " & mCodigoUsuario
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Hora > " & preparaData(mHora)
    xCondicao = xCondicao & " AND Sequencia > " & mSequencia
    xCondicao = xCondicao & " AND [Nome Interno do Programa] > " & preparaTexto(mNomeInternoPrograma)
    xCondicao = xCondicao & " AND Operacao > " & mOperacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Usuario] ASC, Data ASC, Hora ASC, Sequencia ASC, [Nome Interno do Programa] ASC, Operacao ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pCodigoUsuario As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pSequencia As Integer, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Codigo do Usuario] = " & pCodigoUsuario
    xCondicao = xCondicao & " AND Data = " & preparaData(mData)
    xCondicao = xCondicao & " AND Hora = " & preparaData(mHora)
    xCondicao = xCondicao & " AND Sequencia = " & pSequencia
    xCondicao = xCondicao & " AND [Nome Interno do Programa] = " & preparaTexto(mNomeInternoPrograma)
    xCondicao = xCondicao & " AND Operacao = " & pOperacao
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Usuario] DESC, Data DESC, Hora DESC, Sequencia DESC, [Nome Interno do Programa] DESC, Operacao DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mSequencia = ProximaSequencia(mCodigoUsuario, mData, mHora, mNomeInternoPrograma, mOperacao)
    gSQL = "INSERT INTO Auditoria ( [Codigo do Usuario], Data, Hora, Sequencia, [Nome Interno do Programa], "
    gSQL = gSQL & "Operacao, Computador, Observacao ) VALUES ( "
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlNumero(1, mSequencia, ", ")
    Call sqlTexto(1, mNomeInternoPrograma, ", ")
    Call sqlNumero(1, mOperacao, ", ")
    Call sqlTexto(1, mComputador, ", ")
    Call sqlTexto(1, Mid(mObservacao, 1, 80), " )")
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigoUsuario As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pSequencia As Integer, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Auditoria SET "
    Call sqlNumero(2, "[Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlData(2, ", Data = ", mData)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlNumero(2, ", Sequencia = ", mSequencia)
    Call sqlTexto(2, ", [Nome Interno do Programa] = ", mNomeInternoPrograma)
    Call sqlNumero(2, ", Operacao = ", mOperacao)
    Call sqlTexto(2, ", Computador = ", mComputador)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    gSQL = gSQL & " WHERE [Codigo do Usuario] = " & pCodigoUsuario
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHora)
    gSQL = gSQL & " AND Sequencia = " & pSequencia
    gSQL = gSQL & " AND [Nome Interno do Programa] = " & pNomeInternoPrograma
    gSQL = gSQL & " AND Operacao = " & pOperacao
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigoUsuario As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pSequencia As Integer, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Auditoria"
    gSQL = gSQL & " WHERE [Codigo do Usuario] = " & pCodigoUsuario
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHora)
    gSQL = gSQL & " AND Sequencia = " & pSequencia
    gSQL = gSQL & " AND [Nome Interno do Programa] = " & pNomeInternoPrograma
    gSQL = gSQL & " AND Operacao = " & pOperacao
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Usuario], Data, Hora, Sequencia, [Nome Interno do Programa], Operacao, Computador, Observacao FROM Auditoria" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function ProximaSequencia(ByVal pCodigoUsuario As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer) As Integer
    ProximaSequencia = 1
    gSQL = "SELECT Sequencia FROM Auditoria"
    gSQL = gSQL & " WHERE [Codigo do Usuario] = " & pCodigoUsuario
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHora)
    gSQL = gSQL & " AND [Nome Interno do Programa] = " & preparaTexto(pNomeInternoPrograma)
    gSQL = gSQL & " AND Operacao = " & pOperacao
    gSQL = gSQL & " ORDER BY Sequencia DESC"
    If Localizar(1) Then
        ProximaSequencia = rsAuditoria("Sequencia").Value + 1
    End If
    rsAuditoria.Close
    Set rsAuditoria = Nothing
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAuditoria = New adodb.Recordset
    rsAuditoria.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAuditoria.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAuditoria.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoUsuario = rsAuditoria("Codigo do Usuario").Value
    Data = rsAuditoria("Data").Value
    Hora = rsAuditoria("Hora").Value
    Sequencia = rsAuditoria("Sequencia").Value
    NomeInternoPrograma = rsAuditoria("Nome Interno do Programa").Value
    Operacao = rsAuditoria("Operacao").Value
    Computador = rsAuditoria("Computador").Value
    Observacao = rsAuditoria("Observacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

