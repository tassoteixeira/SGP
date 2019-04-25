VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoJustificativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNumero As Long
Private mData As Date
Private mHora As Date
Private mOperacao As String
Private mCodigoUsuario As Integer
Private mNomeInternoPrograma As String
Private mComputador As String
Private mJustificativa As String
Private mCodigoFuncionario As Integer
Private mDadosInterno As String

Private rsMovJustificativa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
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


Public Property Get Operacao() As String
    Operacao = mOperacao
End Property
Public Property Let Operacao(ByVal Valor As String)
    mOperacao = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property


Public Property Get NomeInternoPrograma() As String
    NomeInternoPrograma = mNomeInternoPrograma
End Property
Public Property Let NomeInternoPrograma(ByVal Valor As String)
    mNomeInternoPrograma = Valor
End Property


Public Property Get Computador() As String
    Computador = mComputador
End Property
Public Property Let Computador(ByVal Valor As String)
    mComputador = Valor
End Property


Public Property Get Justificativa() As String
    Justificativa = mJustificativa
End Property
Public Property Let Justificativa(ByVal Valor As String)
    mJustificativa = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get DadosInterno() As String
    DadosInterno = mDadosInterno
End Property
Public Property Let DadosInterno(ByVal Valor As String)
    mDadosInterno = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Numero = " & pNumero
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovJustificativa.Close
    Set rsMovJustificativa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mNumero = ProximoNumero
    gSQL = "INSERT INTO MovimentoJustificativa ( Numero, Data, Hora, Operacao, [Codigo do Usuario], "
    gSQL = gSQL & "[Nome Interno do Programa], Computador, Justificativa, [Codigo do Funcionario], "
    gSQL = gSQL & "[Dados Interno] ) VALUES ( "
    Call sqlNumero(1, mNumero, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlTexto(1, mOperacao, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlTexto(1, mNomeInternoPrograma, ", ")
    Call sqlTexto(1, mComputador, ", ")
    Call sqlTexto(1, mJustificativa, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlTexto(1, mDadosInterno, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoJustificativa SET "
    Call sqlNumero(2, "Numero = ", mNumero)
    Call sqlData(2, ", Data = ", mData)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlTexto(2, ", Operacao = ", mOperacao)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlTexto(2, ", [Nome Interno do Programa] = ", mNomeInternoPrograma)
    Call sqlTexto(2, ", Computador = ", mComputador)
    Call sqlTexto(2, ", Justificativa = ", mJustificativa)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlTexto(2, ", [Dados Interno] = ", mDadosInterno)
    gSQL = gSQL & " WHERE Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoJustificativa"
    gSQL = gSQL & " WHERE Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Numero, Data, Hora, Operacao, [Codigo do Usuario], [Nome Interno do Programa], Computador, Justificativa, [Codigo do Funcionario], [Dados Interno] FROM MovimentoJustificativa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function ProximoNumero() As Integer
    ProximoNumero = 1
    gSQL = "SELECT Numero FROM MovimentoJustificativa"
    gSQL = gSQL & " ORDER BY Numero DESC"
    If Localizar(1) Then
        ProximoNumero = rsMovJustificativa("Numero").Value + 1
    End If
    rsMovJustificativa.Close
    Set rsMovJustificativa = Nothing
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovJustificativa = New adodb.Recordset
    rsMovJustificativa.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovJustificativa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovJustificativa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Numero = rsMovJustificativa("Numero").Value
    Data = rsMovJustificativa("Data").Value
    Hora = rsMovJustificativa("Hora").Value
    Operacao = rsMovJustificativa("Operacao").Value
    CodigoUsuario = rsMovJustificativa("Codigo do Usuario").Value
    NomeInternoPrograma = rsMovJustificativa("Nome Interno do Programa").Value
    Computador = rsMovJustificativa("Computador").Value
    Justificativa = rsMovJustificativa("Justificativa").Value
    CodigoFuncionario = rsMovJustificativa("Codigo do Funcionario").Value
    DadosInterno = rsMovJustificativa("Dados Interno").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

