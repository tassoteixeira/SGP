VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cControleSolicitacaoTef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim rsConfiguracao As New ADODB.Recordset
Dim lConn As ADODB.Connection
Dim lBdAccess As Boolean
Dim lBdMySql As Boolean
Dim lBdSqlServer As Boolean
Dim lBdOracle As Boolean
Dim lBdPostgreSql As Boolean

Public Function ProximaSolicitacaoTEF(ByVal nEmpresa As Integer) As Long
    Dim lRecordsAffected As Long

On Error GoTo trata_erro
    
    ProximaSolicitacaoTEF = 1
    
    AbreConexao
    
    If lBdPostgreSql Then
        'Adiciona 1 para o "Controle de Solicitacao TEF"
        gSQL = "UPDATE impressora_fiscal SET "
        Call sqlNumero(2, "cont_sol_tef = cont_sol_tef + ", 1)
        'para o marcelo nao precisa testar a empresa
        'gSQL = gSQL & " WHERE empresa = " & nEmpresa
        lConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        'Busca "Controle de Solicitacao TEF"
        'para o marcelo nao precisa testar a empresa
        'gSQL = "SELECT cont_sol_tef FROM impressora_fiscal WHERE empresa = " & nEmpresa
        gSQL = "SELECT cont_sol_tef FROM impressora_fiscal"
        'MsgBox "gSQL=" & gSQL
        If Localizar(1) Then
            ProximaSolicitacaoTEF = rsConfiguracao("cont_sol_tef").Value
        End If
    Else
        'Adiciona 1 para o "Controle de Solicitacao TEF"
        gSQL = "UPDATE Configuracao SET "
        Call sqlNumero(2, "[Controle de Solicitacao TEF] = [Controle de Solicitacao TEF] + ", 1)
        gSQL = gSQL & " WHERE Empresa = " & nEmpresa
        lConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        'Busca "Controle de Solicitacao TEF"
        gSQL = "SELECT [Controle de Solicitacao TEF] FROM Configuracao WHERE Empresa = " & nEmpresa
        If Localizar(1) Then
            ProximaSolicitacaoTEF = rsConfiguracao("Controle de Solicitacao TEF").Value
        End If
    End If
    
    rsConfiguracao.Close
    Set rsConfiguracao = Nothing
    lConn.Close
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConfiguracao = New ADODB.Recordset
    rsConfiguracao.CursorLocation = adUseClient
    i = Len(gSQL)
    If lBdPostgreSql Then
        If xQtdRegistro > 0 Then
            gSQL = gSQL & " LIMIT " & xQtdRegistro
        End If
    Else
        If xQtdRegistro > 0 Then
            gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
        End If
    End If

    'MsgBox "gSQL=" & gSQL
    rsConfiguracao.Open gSQL, lConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConfiguracao.EOF Then
        AtualizaRecordset = True
    End If
End Function

Private Function AbreConexao() As Boolean
    Dim xNomeGerenciadorBancoDados As String
    Dim xNomeBancoDados As String
    Dim xConexaoConnectionString As String
    Dim xDriveBD As String
    Dim xDiretorioBD As String
    Dim xNomeInternoBD As String
    Dim xPortaBD As String
    
    On Error GoTo trata_erro
    
    AbreConexao = False
    lBdAccess = False
    lBdMySql = False
    lBdSqlServer = False
    lBdOracle = False
    lBdPostgreSql = False
    xNomeGerenciadorBancoDados = ReadINI("SGBD", "Gerenciador de Banco de Dados", ArqSgpIni)
    
    If xNomeGerenciadorBancoDados = "ACCESS" Then
        lBdAccess = True
    ElseIf xNomeGerenciadorBancoDados = "MYSQL" Then
        lBdMySql = True
    ElseIf xNomeGerenciadorBancoDados = "SQLSERVER" Then
        lBdSqlServer = True
    ElseIf xNomeGerenciadorBancoDados = "ORACLE" Then
        lBdOracle = True
    ElseIf xNomeGerenciadorBancoDados = "POSTGRESQL" Then
        lBdPostgreSql = True
    End If
    
    xDriveBD = ReadINI("LOCAL", "Drive", ArqSgpIni)
    xDiretorioBD = ReadINI("LOCAL", "Diretorio BD", ArqSgpIni)
    xNomeBancoDados = ReadINI("LOCAL", "Nome do Banco de Dados", ArqSgpIni)
    xNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", ArqSgpIni)
    xPortaBD = ReadINI("LOCAL", "Porta", ArqSgpIni)
    If lBdSqlServer Then
        If xNomeInternoBD <> "sgp_data" Then
            xNomeInternoBD = "sgp_data" & xNomeInternoBD
        End If
    End If
    
    If lBdAccess Then
        xConexaoConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & xDriveBD & xDiretorioBD & xNomeBancoDados
    ElseIf lBdSqlServer Then
        xConexaoConnectionString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xNomeBancoDados & ";INITIAL CATALOG=" & xNomeInternoBD & ";USER ID=sa;PASSWORD=cerrado72;"
    ElseIf lBdPostgreSql Then
        'MsgBox "VOU ABRIR O POSTGRESQL"
        xConexaoConnectionString = "DRIVER={PostGreSQL ANSI};Port=" & xPortaBD & ";SERVER=" & xNomeBancoDados & ";DATABASE=" & xNomeInternoBD & ";UID=ADMIN;PWD=UN1S1S;OPTION=3;CursorType=1"
    End If
    Set lConn = New ADODB.Connection
    lConn.ConnectionString = xConexaoConnectionString
    lConn.Open
    'MsgBox "BANCO ABERTO"
    AbreConexao = True
    Exit Function

trata_erro:
    MsgBox "Erro ao conectar com o banco de dados", vbCritical, "Erro de Integridade"
'    Call CriaLogTEF("lBdPostgreSql=" & lBdPostgreSql)
'    Call CriaLogTEF("xConexaoConnectionString: " & xConexaoConnectionString)
End Function
Sub sqlNumero(ByVal pDelimitador As Integer, ByVal pString1 As String, ByVal pString2 As String)
    If lBdAccess Then
        If pDelimitador = 1 Then
            gSQL = gSQL & CLng(pString1) & pString2
        ElseIf pDelimitador = 2 Then
            gSQL = gSQL & pString1 & CLng(pString2)
        End If
    ElseIf lBdSqlServer Then
        If pDelimitador = 1 Then
            gSQL = gSQL & CLng(pString1) & pString2
        ElseIf pDelimitador = 2 Then
            gSQL = gSQL & pString1 & CLng(pString2)
        End If
    ElseIf lBdPostgreSql Then
        If pDelimitador = 1 Then
            gSQL = gSQL & CLng(pString1) & pString2
        ElseIf pDelimitador = 2 Then
            gSQL = gSQL & pString1 & CLng(pString2)
        End If
    End If
End Sub

