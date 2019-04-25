VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConvercaoMedicaoCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cConexao As adodb.Connection
Public mConversao As Integer
Public mMedicaoTanque10 As Long
Public mMedicaoTanque15 As Long
Public mMedicaoTanque30 As Long

Private rsConvercaoMedicaoCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Conversao() As Integer
    Conversao = mConversao
End Property
Public Property Let Conversao(ByVal Valor As Integer)
    mConversao = Valor
End Property


Public Property Get MedicaoTanque10() As Long
    MedicaoTanque10 = mMedicaoTanque10
End Property
Public Property Let MedicaoTanque10(ByVal Valor As Long)
    mMedicaoTanque10 = Valor
End Property


Public Property Get MedicaoTanque15() As Long
    MedicaoTanque15 = mMedicaoTanque15
End Property
Public Property Let MedicaoTanque15(ByVal Valor As Long)
    mMedicaoTanque15 = Valor
End Property


Public Property Get MedicaoTanque30() As Long
    MedicaoTanque30 = mMedicaoTanque30
End Property
Public Property Let MedicaoTanque30(ByVal Valor As Long)
    mMedicaoTanque30 = Valor
End Property


Public Property Set Conexao(Valor As adodb.Connection)
    Set cConexao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nConversao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " Where Conversao = " & nConversao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " Where Conversao < " & mConversao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Conversao DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" Where Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " Where Conversao > " & mConversao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Conversao ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " Where Conversao > " & mConversao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Conversao ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " Where Conversao = " & mConversao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Conversao DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsConvercaoMedicaoCombustivel.Close
    Set rsConvercaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Convercao_Medicao_Combustivel ( Conversao, [Medicao Tanque 10], [Medicao Tanque 15], [Medicao Tanque 30] ) VALUES ( "
    Call sqlNumero(1, mConversao, ", ")
    Call sqlNumero(1, mMedicaoTanque10, ", ")
    Call sqlNumero(1, mMedicaoTanque15, ", ")
    Call sqlNumero(1, mMedicaoTanque30, " )")
    cConexao.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nConversao As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Convercao_Medicao_Combustivel SET "
    Call sqlNumero(2, "Conversao = ", mConversao)
    Call sqlNumero(2, ", [Medicao Tanque 10] = ", mMedicaoTanque10)
    Call sqlNumero(2, ", [Medicao Tanque 15] = ", mMedicaoTanque15)
    Call sqlNumero(2, ", [Medicao Tanque 30] = ", mMedicaoTanque30)
    gSQL = gSQL & " WHERE Conversao = " & nConversao
    cConexao.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nConversao As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Convercao_Medicao_Combustivel"
    gSQL = gSQL & " WHERE Conversao = " & nConversao
    cConexao.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Conversao, [Medicao Tanque 10], [Medicao Tanque 15], [Medicao Tanque 30] FROM Convercao_Medicao_Combustivel" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConvercaoMedicaoCombustivel = New adodb.Recordset
    rsConvercaoMedicaoCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConvercaoMedicaoCombustivel.Open gSQL, cConexao, adOpenForwardOnly, adLockReadOnly
    If Not rsConvercaoMedicaoCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Conversao = rsConvercaoMedicaoCombustivel("Conversao").Value
    MedicaoTanque10 = rsConvercaoMedicaoCombustivel("Medicao Tanque 10").Value
    MedicaoTanque15 = rsConvercaoMedicaoCombustivel("Medicao Tanque 15").Value
    MedicaoTanque30 = rsConvercaoMedicaoCombustivel("Medicao Tanque 30").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

