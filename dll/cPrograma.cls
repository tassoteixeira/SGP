VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mTipo As String
Private mNomeInterno As String
Private mNomeparaMenu As String
Private mNomenoDisco As String
Private mConfiguravel As String
Private mObservacao As String
Private mCodigo As Integer

Private rsPrograma As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Tipo() As String
    Tipo = mTipo
End Property
Public Property Let Tipo(ByVal Valor As String)
    mTipo = Valor
End Property


Public Property Get NomeInterno() As String
    NomeInterno = mNomeInterno
End Property
Public Property Let NomeInterno(ByVal Valor As String)
    mNomeInterno = Valor
End Property


Public Property Get NomeparaMenu() As String
    NomeparaMenu = mNomeparaMenu
End Property
Public Property Let NomeparaMenu(ByVal Valor As String)
    mNomeparaMenu = Valor
End Property


Public Property Get NomenoDisco() As String
    NomenoDisco = mNomenoDisco
End Property
Public Property Let NomenoDisco(ByVal Valor As String)
    mNomenoDisco = Valor
End Property


Public Property Get Configuravel() As String
    Configuravel = mConfiguravel
End Property
Public Property Let Configuravel(ByVal Valor As String)
    mConfiguravel = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property


Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarNomeDisco(ByVal nNomeDisco As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNomeDisco = False
    xCondicao = " WHERE [Nome no Disco] = " & preparaTexto(nNomeDisco)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNomeDisco = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNomeInterno(ByVal nNomeInterno As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNomeInterno = False
    xCondicao = " WHERE [Nome Interno] = " & preparaTexto(nNomeInterno)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNomeInterno = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE [Nome no Disco] < " & preparaTexto(mNomenoDisco)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Nome no Disco] DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNomeMenu(ByVal pTipo As String, ByVal pNomeMenu As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNomeMenu = False
    xCondicao = ""
    xCondicao = xCondicao & " WHERE Tipo = " & preparaTexto(pTipo)
    xCondicao = xCondicao & " AND [Nome para Menu] = " & preparaTexto(pNomeMenu)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNomeMenu = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Nome no Disco] > " & preparaTexto(" ")
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Nome no Disco] ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE [Nome no Disco] > " & preparaTexto(mNomenoDisco)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Nome no Disco] ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarTipoNome(ByVal nTipo As String, ByVal nNomeMenu As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarTipoNome = False
    xCondicao = ""
    xCondicao = xCondicao & " WHERE Tipo = " & preparaTexto(nTipo)
    xCondicao = xCondicao & " AND [Nome para Menu] = " & preparaTexto(nNomeMenu)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarTipoNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Nome no Disco] < " & preparaTexto("ZZZZZZZZZZZZZZZZZZZZ")
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Nome no Disco] DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Programa ( Tipo, [Nome Interno], [Nome para Menu], [Nome no Disco], Configuravel, "
    gSQL = gSQL & "Observacao, Codigo ) VALUES ( "
    Call sqlTexto(1, mTipo, ", ")
    Call sqlTexto(1, mNomeInterno, ", ")
    Call sqlTexto(1, mNomeparaMenu, ", ")
    Call sqlTexto(1, mNomenoDisco, ", ")
    Call sqlTexto(1, mConfiguravel, ", ")
    Call sqlTexto(1, mObservacao, ", ")
    Call sqlNumero(1, mCodigo, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nNomeNoDisco As String) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Programa SET "
    Call sqlTexto(2, "Tipo = ", mTipo)
    Call sqlTexto(2, ", [Nome Interno] = ", mNomeInterno)
    Call sqlTexto(2, ", [Nome para Menu] = ", mNomeparaMenu)
    Call sqlTexto(2, ", [Nome no Disco] = ", mNomenoDisco)
    Call sqlTexto(2, ", Configuravel = ", mConfiguravel)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    gSQL = gSQL & " WHERE [Nome no Disco] = " & preparaTexto(nNomeNoDisco)
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nNomeNoDisco As String) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Programa"
    gSQL = gSQL & " WHERE [Nome no Disco] = " & preparaTexto(nNomeNoDisco)
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


Public Function ProximoCodigo() As Integer
Dim xCondicao As String
On Error GoTo trata_erro

    ProximoCodigo = 1
    xCondicao = " WHERE Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    If Localizar Then
        ProximoCodigo = Codigo + 1
    End If
    Exit Function

trata_erro:
    ProximoCodigo = 0
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Tipo, [Nome Interno], [Nome para Menu], [Nome no Disco], Configuravel, Observacao, Codigo FROM Programa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPrograma = New adodb.Recordset
    rsPrograma.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPrograma.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrograma.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Tipo = rsPrograma("Tipo").Value
    NomeInterno = rsPrograma("Nome Interno").Value
    NomeparaMenu = rsPrograma("Nome para Menu").Value
    NomenoDisco = rsPrograma("Nome no Disco").Value
    Configuravel = rsPrograma("Configuravel").Value
    Observacao = rsPrograma("Observacao").Value
    Codigo = rsPrograma("Codigo").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsPrograma.Close
    Set rsPrograma = Nothing
End Function

