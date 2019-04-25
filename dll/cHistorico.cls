VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As Integer
Private mNatureza As String
Private mNome As String

Private rsHistorico As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Natureza() As String
    Natureza = mNatureza
End Property
Public Property Let Natureza(ByVal Valor As String)
    mNatureza = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo = " & nCodigo
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
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nEmpresa As Integer, ByVal nNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(nNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Historico ( Empresa, Codigo, Natureza, Nome ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNatureza, ", ")
    Call sqlTexto(1, mNome, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Historico SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", Natureza = ", mNatureza)
    Call sqlTexto(2, ", Nome = ", mNome)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Historico"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, Natureza, Nome FROM Historico" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsHistorico = New adodb.Recordset
    rsHistorico.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsHistorico.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsHistorico.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsHistorico("Empresa").Value
    Codigo = rsHistorico("Codigo").Value
    Natureza = rsHistorico("Natureza").Value
    Nome = rsHistorico("Nome").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsHistorico.Close
    Set rsHistorico = Nothing
End Function

