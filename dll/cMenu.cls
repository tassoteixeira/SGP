VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mUsuario As Integer
Private mTipo As String
Private mMenu As String
Private mDisco As String
Private mInterno As String

Private rsMenu As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Usuario() As Integer
    Usuario = mUsuario
End Property
Public Property Let Usuario(ByVal Valor As Integer)
    mUsuario = Valor
End Property


Public Property Get Tipo() As String
    Tipo = mTipo
End Property
Public Property Let Tipo(ByVal Valor As String)
    mTipo = Valor
End Property


Public Property Get Menu() As String
    Menu = mMenu
End Property
Public Property Let Menu(ByVal Valor As String)
    mMenu = Valor
End Property


Public Property Get Disco() As String
    Disco = mDisco
End Property
Public Property Let Disco(ByVal Valor As String)
    mDisco = Valor
End Property


Public Property Get Interno() As String
    Interno = mInterno
End Property
Public Property Let Interno(ByVal Valor As String)
    mInterno = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function ExisteProgramaUsuario(ByVal nUsuario As Integer, ByVal nTipo As String, ByVal nMenu As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteProgramaUsuario = False
    xCondicao = " WHERE Usuario = " & nUsuario
    xCondicao = xCondicao & " AND Tipo = " & preparaTexto(nTipo)
    xCondicao = xCondicao & " AND Menu = " & preparaTexto(nMenu)
    gSQL = PreparaSQL(xCondicao, "")
    ExisteProgramaUsuario = Localizar(1)
    rsMenu.Close
    Set rsMenu = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Menu ( Usuario, Tipo, Menu, Disco, Interno ) VALUES ( "
    Call sqlNumero(1, mUsuario, ", ")
    Call sqlTexto(1, mTipo, ", ")
    Call sqlTexto(1, mMenu, ", ")
    Call sqlTexto(1, mDisco, ", ")
    Call sqlTexto(1, mInterno, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Excluir(ByVal nUsuario As Integer, ByVal nTipo As String, ByVal nMenu As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Menu"
    gSQL = gSQL & " WHERE Usuario = " & nUsuario
    gSQL = gSQL & " AND Tipo = " & preparaTexto(nTipo)
    gSQL = gSQL & " AND Menu = " & preparaTexto(nMenu)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Usuario, Tipo, Menu, Disco, Interno FROM Menu" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMenu = New adodb.Recordset
    rsMenu.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMenu.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMenu.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Usuario = rsMenu("Usuario").Value
    Tipo = rsMenu("Tipo").Value
    Menu = rsMenu("Menu").Value
    Disco = rsMenu("Disco").Value
    Interno = rsMenu("Interno").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

