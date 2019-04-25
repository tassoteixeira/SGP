VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mDia1Inicial As Integer
Private mDia1Final As Integer
Private mDia1Vencimento As Integer
Private mDia2Inicial As Integer
Private mDia2Final As Integer
Private mDia2Vencimento As Integer
Private mDia3Inicial As Integer
Private mDia3Final As Integer
Private mDia3Vencimento As Integer

Private rsVencimento As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Dia1Inicial() As Integer
    Dia1Inicial = mDia1Inicial
End Property
Public Property Let Dia1Inicial(ByVal Valor As Integer)
    mDia1Inicial = Valor
End Property


Public Property Get Dia1Final() As Integer
    Dia1Final = mDia1Final
End Property
Public Property Let Dia1Final(ByVal Valor As Integer)
    mDia1Final = Valor
End Property


Public Property Get Dia1Vencimento() As Integer
    Dia1Vencimento = mDia1Vencimento
End Property
Public Property Let Dia1Vencimento(ByVal Valor As Integer)
    mDia1Vencimento = Valor
End Property


Public Property Get Dia2Inicial() As Integer
    Dia2Inicial = mDia2Inicial
End Property
Public Property Let Dia2Inicial(ByVal Valor As Integer)
    mDia2Inicial = Valor
End Property


Public Property Get Dia2Final() As Integer
    Dia2Final = mDia2Final
End Property
Public Property Let Dia2Final(ByVal Valor As Integer)
    mDia2Final = Valor
End Property


Public Property Get Dia2Vencimento() As Integer
    Dia2Vencimento = mDia2Vencimento
End Property
Public Property Let Dia2Vencimento(ByVal Valor As Integer)
    mDia2Vencimento = Valor
End Property


Public Property Get Dia3Inicial() As Integer
    Dia3Inicial = mDia3Inicial
End Property
Public Property Let Dia3Inicial(ByVal Valor As Integer)
    mDia3Inicial = Valor
End Property


Public Property Get Dia3Final() As Integer
    Dia3Final = mDia3Final
End Property
Public Property Let Dia3Final(ByVal Valor As Integer)
    mDia3Final = Valor
End Property


Public Property Get Dia3Vencimento() As Integer
    Dia3Vencimento = mDia3Vencimento
End Property
Public Property Let Dia3Vencimento(ByVal Valor As Integer)
    mDia3Vencimento = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
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
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
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
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Vencimento ( Codigo, [Dia 1 Inicial], [Dia 1 Final], [Dia 1 Vencimento], [Dia 2 Inicial], "
    gSQL = gSQL & "[Dia 2 Final], [Dia 2 Vencimento], [Dia 3 Inicial], [Dia 3 Final], [Dia 3 Vencimento] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlNumero(1, mDia1Inicial, ", ")
    Call sqlNumero(1, mDia1Final, ", ")
    Call sqlNumero(1, mDia1Vencimento, ", ")
    Call sqlNumero(1, mDia2Inicial, ", ")
    Call sqlNumero(1, mDia2Final, ", ")
    Call sqlNumero(1, mDia2Vencimento, ", ")
    Call sqlNumero(1, mDia3Inicial, ", ")
    Call sqlNumero(1, mDia3Final, ", ")
    Call sqlNumero(1, mDia3Vencimento, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Vencimento SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlNumero(2, ", [Dia 1 Inicial] = ", mDia1Inicial)
    Call sqlNumero(2, ", [Dia 1 Final] = ", mDia1Final)
    Call sqlNumero(2, ", [Dia 1 Vencimento] = ", mDia1Vencimento)
    Call sqlNumero(2, ", [Dia 2 Inicial] = ", mDia2Inicial)
    Call sqlNumero(2, ", [Dia 2 Final] = ", mDia2Final)
    Call sqlNumero(2, ", [Dia 2 Vencimento] = ", mDia2Vencimento)
    Call sqlNumero(2, ", [Dia 3 Inicial] = ", mDia3Inicial)
    Call sqlNumero(2, ", [Dia 3 Final] = ", mDia3Final)
    Call sqlNumero(2, ", [Dia 3 Vencimento] = ", mDia3Vencimento)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Vencimento"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo,  [Dia 1 Inicial],  [Dia 1 Final],  [Dia 1 Vencimento],  [Dia 2 Inicial],  [Dia 2 Final],  [Dia 2 Vencimento],  [Dia 3 Inicial],  [Dia 3 Final],  [Dia 3 Vencimento] FROM Vencimento" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsVencimento = New adodb.Recordset
    rsVencimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsVencimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsVencimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsVencimento("Codigo").Value
    Dia1Inicial = rsVencimento("Dia 1 Inicial").Value
    Dia1Final = rsVencimento("Dia 1 Final").Value
    Dia1Vencimento = rsVencimento("Dia 1 Vencimento").Value
    Dia2Inicial = rsVencimento("Dia 2 Inicial").Value
    Dia2Final = rsVencimento("Dia 2 Final").Value
    Dia2Vencimento = rsVencimento("Dia 2 Vencimento").Value
    Dia3Inicial = rsVencimento("Dia 3 Inicial").Value
    Dia3Final = rsVencimento("Dia 3 Final").Value
    Dia3Vencimento = rsVencimento("Dia 3 Vencimento").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsVencimento.Close
    Set rsVencimento = Nothing
End Function

