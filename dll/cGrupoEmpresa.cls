VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrupoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoGrupo As Integer
Private mCodigoEmpresa As Integer
Private mNomeGrupo As String
Private mNomeEmpresaReduzido As String

Private rsGrupoEmpresa As New adodb.Recordset
'Const errNao_Consegui_Localizar = vbObjectError + 512 + 2
'Const errNao_Consegui_Salvar_Dados = vbObjectError + 512 + 3



'Inicio Propriedades da Classe
Public Property Get CodigoGrupo() As Integer
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As Integer)
    mCodigoGrupo = Valor
End Property


Public Property Get CodigoEmpresa() As Integer
    CodigoEmpresa = mCodigoEmpresa
End Property
Public Property Let CodigoEmpresa(ByVal Valor As Integer)
    mCodigoEmpresa = Valor
End Property


Public Property Get NomeGrupo() As String
    NomeGrupo = mNomeGrupo
End Property
Public Property Let NomeGrupo(ByVal Valor As String)
    mNomeGrupo = Valor
End Property


Public Property Get NomeEmpresaReduzido() As String
    NomeEmpresaReduzido = mNomeEmpresaReduzido
End Property
Public Property Let NomeEmpresaReduzido(ByVal Valor As String)
    mNomeEmpresaReduzido = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarGrupo(ByVal nCodigoEmpresa As Integer) As Boolean

On Error GoTo trata_erro
    
    LocalizarGrupo = False
    gSQL = PreparaSQL("WHERE [Codigo da Empresa] = " & nCodigoEmpresa, "ORDER BY [Codigo da Empresa] ASC")
    LocalizarGrupo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarGrupoEmpresa(ByVal nCodigoGrupo As Integer, ByVal nCodigoEmpresa As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarGrupoEmpresa = False
    xCondicao = " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    xCondicao = xCondicao & " AND [Codigo da Empresa] = " & nCodigoEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Empresa] ASC")
    LocalizarGrupoEmpresa = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior(ByVal nCodigoGrupo As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarAnterior = False
    xCondicao = " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    xCondicao = xCondicao & " AND [Codigo da Empresa] < " & mCodigoEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Empresa] DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nCodigoGrupo As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    xCondicao = xCondicao & " AND [Codigo da Empresa] > " & 0
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Empresa] ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo(ByVal nCodigoGrupo As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarProximo = False
    xCondicao = " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    xCondicao = xCondicao & " AND [Codigo da Empresa] > " & mCodigoEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Empresa] ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nCodigoGrupo As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarUltimo = False
    xCondicao = " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    xCondicao = xCondicao & " AND [Codigo da Empresa] < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Empresa] DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Grupo_Empresa ( [Codigo do Grupo], [Codigo da Empresa], [Nome do Grupo], [Nome da Empresa Reduzido] ) VALUES ( "
    Call sqlNumero(1, mCodigoGrupo, ", ")
    Call sqlNumero(1, mCodigoEmpresa, ", ")
    Call sqlTexto(1, mNomeGrupo, ", ")
    Call sqlTexto(1, mNomeEmpresaReduzido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigoGrupo As Integer, ByVal nCodigoEmpresa As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Grupo_Empresa SET "
    Call sqlTexto(2, "[Nome do Grupo] = ", mNomeGrupo)
    Call sqlTexto(2, ", [Nome da Empresa Reduzido] = ", mNomeEmpresaReduzido)
    gSQL = gSQL & " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    gSQL = gSQL & " AND [Codigo da Empresa] = " & nCodigoEmpresa
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigoGrupo As Integer, ByVal nCodigoEmpresa As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Grupo_Empresa"
    gSQL = gSQL & " WHERE [Codigo do Grupo] = " & nCodigoGrupo
    gSQL = gSQL & " AND [Codigo da Empresa] = " & nCodigoEmpresa
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Grupo], [Codigo da Empresa], [Nome do Grupo], [Nome da Empresa Reduzido] FROM Grupo_Empresa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsGrupoEmpresa = New adodb.Recordset
    rsGrupoEmpresa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsGrupoEmpresa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsGrupoEmpresa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoGrupo = rsGrupoEmpresa("Codigo do Grupo").Value
    CodigoEmpresa = rsGrupoEmpresa("Codigo da Empresa").Value
    NomeGrupo = rsGrupoEmpresa("Nome do Grupo").Value
    NomeEmpresaReduzido = rsGrupoEmpresa("Nome da Empresa Reduzido").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsGrupoEmpresa.Close
    Set rsGrupoEmpresa = Nothing
End Function

