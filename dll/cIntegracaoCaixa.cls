VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIntegracaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNomeIntegracao As String
Private mContaDebito As String
Private mContaCredito As String
Private mHistoricoPadrao As Integer

Private rsIntegracaoCaixa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NomeIntegracao() As String
    NomeIntegracao = mNomeIntegracao
End Property
Public Property Let NomeIntegracao(ByVal Valor As String)
    mNomeIntegracao = Valor
End Property


Public Property Get ContaDebito() As String
    ContaDebito = mContaDebito
End Property
Public Property Let ContaDebito(ByVal Valor As String)
    mContaDebito = Valor
End Property


Public Property Get ContaCredito() As String
    ContaCredito = mContaCredito
End Property
Public Property Let ContaCredito(ByVal Valor As String)
    mContaCredito = Valor
End Property


Public Property Get HistoricoPadrao() As Integer
    HistoricoPadrao = mHistoricoPadrao
End Property
Public Property Let HistoricoPadrao(ByVal Valor As Integer)
    mHistoricoPadrao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarNome(ByVal nEmpresa As Integer, ByVal nNomeIntegracao As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNome = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Nome da Integracao] = " & preparaTexto(nNomeIntegracao)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsIntegracaoCaixa.Close
    Set rsIntegracaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Nome da Integracao] < " & preparaTexto(mNomeIntegracao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Nome da Integracao] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsIntegracaoCaixa.Close
    Set rsIntegracaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Nome da Integracao] > " & preparaTexto(mNomeIntegracao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Nome da Integracao] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsIntegracaoCaixa.Close
    Set rsIntegracaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Nome da Integracao] > " & preparaTexto(mNomeIntegracao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Nome da Integracao] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsIntegracaoCaixa.Close
    Set rsIntegracaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer, ByVal nNomeIntegracao As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Nome da Integracao] = " & preparaTexto(mNomeIntegracao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Nome da Integracao] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsIntegracaoCaixa.Close
    Set rsIntegracaoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Integracao_Caixa ( Empresa, [Nome da Integracao], [Conta Debito], "
    gSQL = gSQL & "[Conta Credito], [Historico Padrao] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mNomeIntegracao, ", ")
    Call sqlTexto(1, mContaDebito, ", ")
    Call sqlTexto(1, mContaCredito, ", ")
    Call sqlNumero(1, mHistoricoPadrao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nNomeIntegracao As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Integracao_Caixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Nome da Integracao] = ", mNomeIntegracao)
    Call sqlTexto(2, ", [Conta Debito] = ", mContaDebito)
    Call sqlTexto(2, ", [Conta Credito] = ", mContaCredito)
    Call sqlNumero(2, ", [Historico Padrao] = ", mHistoricoPadrao)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Nome da Integracao] = " & preparaTexto(nNomeIntegracao)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nNomeIntegracao As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Integracao_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Nome da Integracao] = " & preparaTexto(nNomeIntegracao)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Nome da Integracao], [Conta Debito], [Conta Credito], [Historico Padrao] FROM Integracao_Caixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsIntegracaoCaixa = New adodb.Recordset
    rsIntegracaoCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsIntegracaoCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsIntegracaoCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsIntegracaoCaixa("Empresa").Value
    NomeIntegracao = rsIntegracaoCaixa("Nome da Integracao").Value
    ContaDebito = rsIntegracaoCaixa("Conta Debito").Value
    ContaCredito = rsIntegracaoCaixa("Conta Credito").Value
    HistoricoPadrao = rsIntegracaoCaixa("Historico Padrao").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

