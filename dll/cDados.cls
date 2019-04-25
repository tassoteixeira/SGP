VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mEmpresa As Integer
Private mContaBancaria As String
Private mEmpresa2 As Integer

Private rsDados As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get ContaBancaria() As String
    ContaBancaria = mContaBancaria
End Property
Public Property Let ContaBancaria(ByVal Valor As String)
    mContaBancaria = Valor
End Property


Public Property Get Empresa2() As Integer
    Empresa2 = mEmpresa2
End Property
Public Property Let Empresa2(ByVal Valor As Integer)
    mEmpresa2 = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsDados.Close
    Set rsDados = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Dados ( Codigo, Empresa, [Conta Bancaria], [Empresa 2] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mContaBancaria, ", ")
    Call sqlNumero(1, mEmpresa2, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Dados SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlNumero(2, ", Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Conta Bancaria] = ", mContaBancaria)
    Call sqlNumero(2, ", [Empresa 2] = ", mEmpresa2)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Empresa, [Conta Bancaria], [Empresa 2] FROM Dados" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsDados = New adodb.Recordset
    rsDados.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsDados.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsDados.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsDados("Codigo").Value
    Empresa = rsDados("Empresa").Value
    ContaBancaria = rsDados("Conta Bancaria").Value
    Empresa2 = rsDados("Empresa 2").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

