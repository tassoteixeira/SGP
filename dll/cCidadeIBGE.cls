VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCidadeIBGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As String
Private mNome As String
Private mUF As String

Private rsCidadeIBGE As New adodb.Recordset

'Inicio das Propriedades da Classe

Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As String)
    mCodigo = Valor
End Property

Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property

Public Property Get UF() As String
    UF = mUF
End Property
Public Property Let UF(ByVal Valor As String)
    mUF = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarNome(ByVal pUF As String, ByVal pNomeCidade As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE UF = " & preparaTexto(pUF)
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNomeCidade)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsCidadeIBGE.Close
    Set rsCidadeIBGE = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, UF FROM CidadeIBGE" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsCidadeIBGE = New adodb.Recordset
    rsCidadeIBGE.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCidadeIBGE.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCidadeIBGE.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsCidadeIBGE("Codigo").Value
    Nome = rsCidadeIBGE("Nome").Value
    UF = rsCidadeIBGE("UF").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
'FIM Funções / Procedures internas

