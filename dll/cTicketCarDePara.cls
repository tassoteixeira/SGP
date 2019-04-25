VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTicketCarDePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoAC As Long
Private mCodigoTCS As Long

Private rsTicketCarDePara As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoAC() As Long
    CodigoAC = mCodigoAC
End Property
Public Property Let CodigoAC(ByVal Valor As Long)
    mCodigoAC = Valor
End Property


Public Property Get CodigoTCS() As Long
    CodigoTCS = mCodigoTCS
End Property
Public Property Let CodigoTCS(ByVal Valor As Long)
    mCodigoTCS = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigoAC As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Codigo AC] = " & pCodigoAC
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsTicketCarDePara.Close
    Set rsTicketCarDePara = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO TicketCarDePara ( [Codigo AC], [Codigo TCS] ) VALUES ( "
    Call sqlNumero(1, mCodigoAC, ", ")
    Call sqlNumero(1, mCodigoTCS, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigoAC As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE TicketCarDePara SET "
    Call sqlNumero(2, "[Codigo AC] = ", mCodigoAC)
    Call sqlNumero(2, ", [Codigo TCS] = ", mCodigoTCS)
    gSQL = gSQL & " WHERE [Codigo AC] = " & pCodigoAC
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigoAC As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM TicketCarDePara"
    gSQL = gSQL & " WHERE [Codigo AC] = " & pCodigoAC
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo AC], [Codigo TCS] FROM TicketCarDePara" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTicketCarDePara = New adodb.Recordset
    rsTicketCarDePara.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTicketCarDePara.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTicketCarDePara.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoAC = rsTicketCarDePara("Codigo AC").Value
    CodigoTCS = rsTicketCarDePara("Codigo TCS").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

