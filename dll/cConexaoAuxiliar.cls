VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConexaoAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cConexao As adodb.Connection

Public Property Get Conexao() As adodb.Connection
    Set Conexao = cConexao
End Property
Private Sub Class_Terminate()
    cConexao.Close
    Set cConexao = Nothing
End Sub


'Inicio das Propriedades da Classe
'Fim das Propriedades da Classe


'Início dos Métodos da Classe
Public Function AbreConexao(ByVal pStringConexao As String) As Boolean
On Error GoTo trata_erro
    
    AbreConexao = False
    
    Set cConexao = New adodb.Connection
    
    'If bdAccess Then
    '    mConexaoConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
    'ElseIf bdSqlServer Then
    '    mConexaoConnectionString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & gNomeBancoDados & ";INITIAL CATALOG=" & "sgp_data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
    'End If
    cConexao.ConnectionString = pStringConexao
    cConexao.Open
    AbreConexao = True
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function RsConexao(ByVal pSQL As String) As adodb.Recordset

On Error GoTo trata_erro
    
    Set RsConexao = New adodb.Recordset
    RsConexao.CursorLocation = adUseClient 'adUseServer  'adUseClient
    RsConexao.Open pSQL, cConexao, adOpenForwardOnly, adLockReadOnly
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ExecutaSql(ByVal pSQL As String) As Long
Dim xRegistroAfetado As Long

On Error GoTo trata_erro
    
    cConexao.Execute pSQL, xRegistroAfetado, adCmdText + adExecuteNoRecords
    ExecutaSql = xRegistroAfetado
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim dos Métodos da Classe

