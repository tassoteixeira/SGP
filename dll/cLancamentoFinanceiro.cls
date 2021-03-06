VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLancamentoFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mContaDebito As String
Private mContaCredito As String
Private mCodigoPrograma As Integer
Private mInformaNumeroDocumento As Boolean
Private mInformaComplemento As Boolean
Private mNomePlanoDebito As Boolean
Private mNomePlanoCredito As Boolean
Private mHistoricoPadrao As Integer

Private rsLancamentoFinanceiro As New adodb.Recordset


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


Public Property Get CodigoPrograma() As Integer
    CodigoPrograma = mCodigoPrograma
End Property
Public Property Let CodigoPrograma(ByVal Valor As Integer)
    mCodigoPrograma = Valor
End Property


Public Property Get InformaNumeroDocumento() As Boolean
    InformaNumeroDocumento = mInformaNumeroDocumento
End Property
Public Property Let InformaNumeroDocumento(ByVal Valor As Boolean)
    mInformaNumeroDocumento = Valor
End Property


Public Property Get InformaComplemento() As Boolean
    InformaComplemento = mInformaComplemento
End Property
Public Property Let InformaComplemento(ByVal Valor As Boolean)
    mInformaComplemento = Valor
End Property


Public Property Get NomePlanoDebito() As Boolean
    NomePlanoDebito = mNomePlanoDebito
End Property
Public Property Let NomePlanoDebito(ByVal Valor As Boolean)
    mNomePlanoDebito = Valor
End Property


Public Property Get NomePlanoCredito() As Boolean
    NomePlanoCredito = mNomePlanoCredito
End Property
Public Property Let NomePlanoCredito(ByVal Valor As Boolean)
    mNomePlanoCredito = Valor
End Property


Public Property Get HistoricoPadrao() As Integer
    HistoricoPadrao = mHistoricoPadrao
End Property
Public Property Let HistoricoPadrao(ByVal Valor As Integer)
    mHistoricoPadrao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio M�todos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
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
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE Nome = " & preparaTexto(pNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
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
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsLancamentoFinanceiro.Close
    Set rsLancamentoFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO LancamentoFinanceiro ( Codigo, Nome, [Conta de Debito], [Conta de Credito], "
    gSQL = gSQL & "[Codigo do Programa], [Informa Numero do Documento], [Informa Complemento], "
    gSQL = gSQL & "[Nome do Plano Debito], [Nome do Plano Credito], [Historico Padrao] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mContaDebito, ", ")
    Call sqlTexto(1, mContaCredito, ", ")
    Call sqlNumero(1, mCodigoPrograma, ", ")
    Call sqlBoolean(1, mInformaNumeroDocumento, ", ")
    Call sqlBoolean(1, mInformaComplemento, ", ")
    Call sqlBoolean(1, mNomePlanoDebito, ", ")
    Call sqlBoolean(1, mNomePlanoCredito, ", ")
    Call sqlNumero(1, mHistoricoPadrao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE LancamentoFinanceiro SET "
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", [Conta de Debito] = ", mContaDebito)
    Call sqlTexto(2, ", [Conta de Credito] = ", mContaCredito)
    Call sqlNumero(2, ", [Codigo do Programa] = ", mCodigoPrograma)
    Call sqlBoolean(2, ", [Informa Numero do Documento] = ", mInformaNumeroDocumento)
    Call sqlBoolean(2, ", [Informa Complemento] = ", mInformaComplemento)
    Call sqlBoolean(2, ", [Nome do Plano Debito] = ", mNomePlanoDebito)
    Call sqlBoolean(2, ", [Nome do Plano Credito] = ", mNomePlanoCredito)
    Call sqlNumero(2, "[Historico Padrao] = ", mHistoricoPadrao)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM LancamentoFinanceiro"
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Fun��es / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, [Conta de Debito], [Conta de Credito], [Codigo do Programa], [Informa Numero do Documento], [Informa Complemento], [Nome do Plano Debito], [Nome do Plano Credito], [Historico Padrao] FROM LancamentoFinanceiro" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsLancamentoFinanceiro = New adodb.Recordset
    rsLancamentoFinanceiro.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsLancamentoFinanceiro.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsLancamentoFinanceiro.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsLancamentoFinanceiro("Codigo").Value
    Nome = rsLancamentoFinanceiro("Nome").Value
    ContaDebito = rsLancamentoFinanceiro("Conta de Debito").Value
    ContaCredito = rsLancamentoFinanceiro("Conta de Credito").Value
    CodigoPrograma = rsLancamentoFinanceiro("Codigo do Programa").Value
    InformaNumeroDocumento = rsLancamentoFinanceiro("Informa Numero do Documento").Value
    InformaComplemento = rsLancamentoFinanceiro("Informa Complemento").Value
    NomePlanoDebito = rsLancamentoFinanceiro("Nome do Plano Debito").Value
    NomePlanoCredito = rsLancamentoFinanceiro("Nome do Plano Credito").Value
    HistoricoPadrao = rsLancamentoFinanceiro("Historico Padrao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

