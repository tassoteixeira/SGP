VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mCodigoFinContaTesouraria As Integer
Private mEmpresa As Integer

Private rsContas As New adodb.Recordset


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


Public Property Get CodigoFinContaTesouraria() As Integer
    CodigoFinContaTesouraria = mCodigoFinContaTesouraria
End Property
Public Property Let CodigoFinContaTesouraria(ByVal Valor As Integer)
    mCodigoFinContaTesouraria = Valor
End Property

Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer, ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo & " AND Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Codigo < " & mCodigo & " AND Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String, ByVal pEmpresa As Integer) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome) & " AND Empresa = " & pEmpresa, "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > 0 AND Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Codigo > " & mCodigo & " AND Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < 9999 AND Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Contas ( Codigo, Nome, CodigoFinContaTesouraria, Empresa ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlNumero(1, mCodigoFinContaTesouraria, ", ")
    Call sqlNumero(1, mEmpresa, " )")

    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer, ByVal pEmpresa As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Contas SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlNumero(2, ", CodigoFinContaTesouraria = ", mCodigoFinContaTesouraria)
    Call sqlNumero(2, ", Empresa = ", mEmpresa)

    gSQL = gSQL & " WHERE Codigo = " & nCodigo & " AND Empresa = " & pEmpresa
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer, ByVal pEmpresa As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Contas"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo & " AND Empresa = " & pEmpresa
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, CodigoFinContaTesouraria, Empresa FROM Contas" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsContas = New adodb.Recordset
    rsContas.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsContas.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsContas.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsContas("Codigo").Value
    Nome = rsContas("Nome").Value
    CodigoFinContaTesouraria = rsContas("CodigoFinContaTesouraria").Value
    Empresa = rsContas("Empresa").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsContas.Close
    Set rsContas = Nothing
End Function

