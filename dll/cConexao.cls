VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cConexao As adodb.Connection
Private mConexaoDrive As String
Private mConexaoDiretorio As String
Private mConexaoBaseDados As String
Private mConexaoConnectionString As String
Public Property Get Conexao() As adodb.Connection
    Set Conexao = cConexao
End Property
Public Function AbreConexao() As Boolean
    Dim xNomeGerenciadorBancoDados As String
    Dim xNomeInternoBD As String
    Dim xNivelSenha As String
    
    On Error GoTo trata_erro

    Call CriaLogCadastroDll("cConexao.AbreConexao - Inicia Abrir Conexao")
    AbreConexao = False
    
    
    gArquivoIni = "c:\Cerrado.Net\Sgp.ini"
    'gArquivoIni = "C:\CerradoPafEcf\NFe\NuvemNFe\NuvemNFe.ini"
    Set cConexao = New adodb.Connection
    mConexaoDrive = ChamaDrive
    mConexaoDiretorio = gDiretorioData
    mConexaoBaseDados = gNomeBancoDados
    bdAccess = False
    bdMySql = False
    bdSqlServer = False
    bdSqlServerAzure = False
    bdOracle = False
    xNomeGerenciadorBancoDados = ReadINI("SGBD", "Gerenciador de Banco de Dados", gArquivoIni)
    xNivelSenha = ReadINI("SGBD", "Senha", gArquivoIni)
    xNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", gArquivoIni)
    If xNomeInternoBD <> "sgp_data" Then
        xNomeInternoBD = "sgp_data" & xNomeInternoBD
    End If
    gNomeUsuarioBD = "sa"
    gSenhaBD = "cerrado72"
    If gNomeBancoDados Like "*cloudapp.net*" Or xNivelSenha = "padrao2" Then
        gSenhaBD = "cRr472#*$pst"
    End If
    
    If xNomeGerenciadorBancoDados = "ACCESS" Then
        bdAccess = True
    ElseIf xNomeGerenciadorBancoDados = "MYSQL" Then
        bdMySql = True
    ElseIf xNomeGerenciadorBancoDados = "SQLSERVER" Then
        bdSqlServer = True
    ElseIf xNomeGerenciadorBancoDados = "ORACLE" Then
        bdOracle = True
    End If
    If gNomeBancoDados Like "*cerrado.database.windows.net*" Then
        gNomeUsuarioBD = "cerrado"
        gSenhaBD = "cRr472#*$pst"
        bdSqlServerAzure = True
    End If
    
    
    If bdAccess Then
        mConexaoConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
        '''cConexao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
        cConexao.ConnectionString = mConexaoConnectionString
    ElseIf bdSqlServer Then
        'mConexaoConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=Tasso;Initial Catalog=Sgp;Data Source=SERVIDOR2000"
        'cConexao.ConnectionTimeout = 45
        mConexaoConnectionString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & gNomeBancoDados & ";INITIAL CATALOG=" & xNomeInternoBD & ";USER ID=" & gNomeUsuarioBD & ";PASSWORD=" & gSenhaBD & ";"
        cConexao.ConnectionString = mConexaoConnectionString
        Call CriaLogCadastroDll("cConexao.AbreConexao cConexao.ConnectionString=" & cConexao.ConnectionString)
    End If
    Set gConn = New adodb.Connection
    gConn = cConexao
    cConexao.Open
    gConn.Open
    
    
    AbreConexao = True
    Call CriaLogCadastroDll("cConexao.AbreConexao - Finaliza Abrir Conexao")
    Exit Function

trata_erro:
   Call CriaLogCadastroDll("cConexao.AbreConexao - Erro ao abrir Conexao. Erro Numero:" & Err.Number & " - Mensagem de Erro: " & Err.Description)
End Function
Public Function FinalizaConexao() As Boolean
    
    On Error GoTo trata_erro

    Call CriaLogCadastroDll("cConexao.FinalizaConexao - Inicia Fechar a Conexao")
    FinalizaConexao = False
    
    cConexao.Close
    gConn.Close
    
    
    FinalizaConexao = True
    Call CriaLogCadastroDll("cConexao.FinalizaConexao - Finaliza Fechar a Conexao")
    Exit Function

trata_erro:
   Call CriaLogCadastroDll("cConexao.FinalizaConexao - Erro ao fechar a Conexao. Erro Numero:" & Err.Number & " - Mensagem de Erro: " & Err.Description)
End Function
Private Sub Class_Initialize()
'    Dim xNomeGerenciadorBancoDados As String
'    Dim xNomeInternoBD As String
    On Error GoTo trata_erro
    
    CriaLogCadastroDll ("Class_Initialize")
    AbreConexao
    
'
'    gArquivoIni = "c:\Cerrado.Net\Sgp.ini"
'    'gArquivoIni = "C:\CerradoPafEcf\NFe\NuvemNFe\NuvemNFe.ini"
'    Set cConexao = New adodb.Connection
'    mConexaoDrive = ChamaDrive
'    mConexaoDiretorio = gDiretorioData
'    mConexaoBaseDados = gNomeBancoDados
'    bdAccess = False
'    bdMySql = False
'    bdSqlServer = False
'    bdOracle = False
'    xNomeGerenciadorBancoDados = ReadINI("SGBD", "Gerenciador de Banco de Dados", gArquivoIni)
'    xNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", gArquivoIni)
'    If xNomeInternoBD <> "sgp_data" Then
'        xNomeInternoBD = "sgp_data" & xNomeInternoBD
'    End If
'    gSenhaBD = "cerrado72"
'    If gNomeBancoDados Like "*cloudapp.net*" Then
'        gSenhaBD = "cRr472#*$pst"
'    End If
'
'    If xNomeGerenciadorBancoDados = "ACCESS" Then
'        bdAccess = True
'    ElseIf xNomeGerenciadorBancoDados = "MYSQL" Then
'        bdMySql = True
'    ElseIf xNomeGerenciadorBancoDados = "SQLSERVER" Then
'        bdSqlServer = True
'    ElseIf xNomeGerenciadorBancoDados = "ORACLE" Then
'        bdOracle = True
'    End If
'    If bdAccess Then
'        mConexaoConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
'        '''cConexao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & mConexaoDrive & mConexaoDiretorio & mConexaoBaseDados
'        cConexao.ConnectionString = mConexaoConnectionString
'    ElseIf bdSqlServer Then
'        'mConexaoConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=Tasso;Initial Catalog=Sgp;Data Source=SERVIDOR2000"
'        'cConexao.ConnectionTimeout = 45
'        mConexaoConnectionString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & gNomeBancoDados & ";INITIAL CATALOG=" & xNomeInternoBD & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
'        cConexao.ConnectionString = mConexaoConnectionString
'    End If
'    Set gConn = New adodb.Connection
'    gConn = cConexao
'    cConexao.Open
'    gConn.Open
'
   CriaLogCadastroDll ("Class_Initialize - CONEXÃO ABERTA OK")

    Exit Sub

trata_erro:
End Sub
Private Sub Class_Terminate()
    On Error GoTo trata_erro
    CriaLogCadastroDll ("Class_Terminate")
    cConexao.Close
    gConn.Close
    Set cConexao = Nothing
    Set gConn = Nothing
    CriaLogCadastroDll ("Class_Terminate - CONEXÃO FECHADA")
    Exit Sub

trata_erro:
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


'Início dos Métodos da Classe
Public Function RsConexao(ByVal pSQL As String) As adodb.Recordset

On Error GoTo trata_erro
    
    Set RsConexao = New adodb.Recordset
    RsConexao.CursorLocation = adUseClient 'adUseServer  'adUseClient
    RsConexao.Open pSQL, cConexao, adOpenForwardOnly, adLockReadOnly
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function RsConexaoTimeOut(ByVal pSQL As String, ByVal pTimeOut) As adodb.Recordset

On Error GoTo trata_erro
    
    Set RsConexaoTimeOut = New adodb.Recordset
    cConexao.CommandTimeout = pTimeOut
    RsConexaoTimeOut.CursorLocation = adUseClient 'adUseServer  'adUseClient
    RsConexaoTimeOut.Open pSQL, cConexao, adOpenForwardOnly, adLockReadOnly
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


Public Function ConexaoAtiva() As Boolean
On Error GoTo trata_erro

    ConexaoAtiva = False
    If gConn.State = 1 Then
        ConexaoAtiva = True
    End If
    Exit Function

trata_erro:
End Function


Public Sub IniciaTransacao()
On Error GoTo trata_erro

    gConn.BeginTrans

    Exit Sub

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Sub CancelaTransacao()
On Error GoTo trata_erro

    gConn.RollbackTrans

    Exit Sub

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Sub ConfirmaTransacao()
On Error GoTo trata_erro

    gConn.CommitTrans

    Exit Sub

trata_erro:
    gConn.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description
End Sub
'Fim dos Métodos da Classe
