VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTanqueCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroTanque As Integer
Private mTipoCombustivel As String
Private mCapacidadeArmazenamento As Long

Private rsTanqueCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroTanque() As Integer
    NumeroTanque = mNumeroTanque
End Property
Public Property Let NumeroTanque(ByVal Valor As Integer)
    mNumeroTanque = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get CapacidadeArmazenamento() As Long
    CapacidadeArmazenamento = mCapacidadeArmazenamento
End Property
Public Property Let CapacidadeArmazenamento(ByVal Valor As Long)
    mCapacidadeArmazenamento = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nNumeroTanque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Numero do Tanque] = " & nNumeroTanque
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Tanque] < " & mNumeroTanque
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Tanque] > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Localiza primeiro tanque de um determinado tipo de combustivel
Public Function LocalizarPrimeiroTanqueComb(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroTanqueComb = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque] ASC")
    LocalizarPrimeiroTanqueComb = Localizar(1)
    If LocalizarPrimeiroTanqueComb = True Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Tanque] > " & mNumeroTanque
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Numero do Tanque] < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Tanque_Combustivel ( Empresa, [Numero do Tanque], [Tipo de Combustivel], [Capacidade de Armazenamento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroTanque, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlNumero(1, mCapacidadeArmazenamento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nNumeroTanque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Tanque_Combustivel SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Tanque] = ", mNumeroTanque)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlNumero(2, ", [Capacidade de Armazenamento] = ", mCapacidadeArmazenamento)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Numero do Tanque] = " & nNumeroTanque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nNumeroTanque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Tanque_Combustivel"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Numero do Tanque] = " & nNumeroTanque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Totaliza a Capacidade de Armazenamento de um determinado Combustivel
Public Function TotalCapacidadeArmazenamento(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalCapacidadeArmazenamento = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM([Capacidade de Armazenamento]) AS Total"
    xSQL = xSQL & "   FROM Tanque_Combustivel"
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    
    Set rsTanqueCombustivel = New adodb.Recordset
    rsTanqueCombustivel.CursorLocation = adUseClient
    rsTanqueCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTanqueCombustivel.EOF Then
        If Not IsNull(rsTanqueCombustivel("Total").Value) Then
            TotalCapacidadeArmazenamento = rsTanqueCombustivel("Total").Value
        End If
    End If
    rsTanqueCombustivel.Close
    Set rsTanqueCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Tanque], [Tipo de Combustivel], [Capacidade de Armazenamento] FROM Tanque_Combustivel" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTanqueCombustivel = New adodb.Recordset
    rsTanqueCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTanqueCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTanqueCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsTanqueCombustivel("Empresa").Value
    NumeroTanque = rsTanqueCombustivel("Numero do Tanque").Value
    TipoCombustivel = rsTanqueCombustivel("Tipo de Combustivel").Value
    CapacidadeArmazenamento = rsTanqueCombustivel("Capacidade de Armazenamento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

