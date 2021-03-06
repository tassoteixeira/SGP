VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConexaoB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cConexaoB As adodb.Connection
Private mConexaoDrive As String
Private mConexaoDiretorio As String
Private mConexaoBaseDados As String
Private mConexaoConnectionString As String
Public Property Get Conexao() As adodb.Connection
    Set Conexao = cConexaoB
End Property
Private Sub Class_Initialize()
    Set cConexaoB = New adodb.Connection
    mConexaoDrive = ChamaDrive
    mConexaoDiretorio = "\vb5\sgp\data\"
    mConexaoBaseDados = "sgp_data_baixa.mdb"
    bdAccess = True
    bdMySql = False
    bdSqlServer = False
    bdOracle = False
    If bdAccess Then
        mConexaoConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
    End If
    '''cConexao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
    cConexaoB.ConnectionString = mConexaoConnectionString
    cConexaoB.Open
End Sub
Private Sub Class_Terminate()
    cConexaoB.Close
    Set cConexaoB = Nothing
End Sub
'Inicio das Propriedades da Classe
Public Property Get Drive() As String
    Drive = mConexaoDrive
End Property

Public Property Get Diretorio() As String
    Diretorio = mConexaoDiretorio
End Property

Public Property Get BaseDados() As String
    BaseDados = mConexaoBaseDados
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mConexaoConnectionString
End Property
'Fim das Propriedades da Classe


'In�cio dos M�todos da Classe
Public Function RsConexao(ByVal xSQL As String) As adodb.Recordset 'Object

On Error GoTo trata_erro
    
    Set RsConexao = New adodb.Recordset
    RsConexao.CursorLocation = adUseClient
    RsConexao.Open xSQL, cConexaoB, adOpenForwardOnly, adLockReadOnly
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim dos M�todos da Classe
