VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConexaoNuvem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mConexaoNuvem As adodb.Connection
Private mIpBancoDados As String
Private mConexaoConnectionString As String

'Inicio das Propriedades da Classe
Public Property Get Conexao() As adodb.Connection
    Set Conexao = mConexaoNuvem
End Property

Public Property Get IpBancoDados() As String
    IpBancoDados = mIpBancoDados
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
    RsConexao.Open pSQL, mConexaoNuvem, adOpenForwardOnly, adLockReadOnly
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ExecutaSql(ByVal pSQL As String) As Long
Dim xRegistroAfetado As Long

On Error GoTo trata_erro
    
    mConexaoNuvem.Execute pSQL, xRegistroAfetado, adCmdText + adExecuteNoRecords
    ExecutaSql = xRegistroAfetado
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ConexaoAtiva() As Boolean
On Error GoTo trata_erro

    ConexaoAtiva = False
    If gConnNuvem.State = 1 Then
        ConexaoAtiva = True
    End If
    Exit Function

trata_erro:
End Function


Public Sub IniciaTransacao()
On Error GoTo trata_erro

    gConnNuvem.BeginTrans

    Exit Sub

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Sub CancelaTransacao()
On Error GoTo trata_erro

    gConnNuvem.RollbackTrans

    Exit Sub

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Sub ConfirmaTransacao()
On Error GoTo trata_erro

    gConnNuvem.CommitTrans

    Exit Sub

trata_erro:
    gConnNuvem.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description
End Sub
'Fim dos Métodos da Classe





Private Sub Class_Initialize()
    Dim xNomeGerenciadorBancoDados As String
    Dim xNomeInternoBD As String
    
    On Error GoTo trata_erro
    
    Set mConexaoNuvem = New adodb.Connection
    mIpBancoDados = gNomeBancoDados
    'gArquivoIni = "C:\CerradoPafEcf\NFe\NuvemNFe\NuvemNFe.ini"
    xNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", gArquivoIni)
    If xNomeInternoBD <> "sgp_data" Then
        xNomeInternoBD = "sgp_nuvem" & xNomeInternoBD
    Else
        xNomeInternoBD = "sgp_nuvem"
    End If
    
    mConexaoConnectionString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & mIpBancoDados & ";INITIAL CATALOG=" & xNomeInternoBD & ";USER ID=" & gNomeUsuarioBD & ";PASSWORD=" & gSenhaBD & ";"
    mConexaoNuvem.ConnectionString = mConexaoConnectionString
    Set gConnNuvem = New adodb.Connection
    gConnNuvem = mConexaoNuvem
    mConexaoNuvem.Open
    gConnNuvem.Open
    Exit Sub

trata_erro:
End Sub
Private Sub Class_Terminate()
    On Error GoTo trata_erro
    
    mConexaoNuvem.Close
    gConnNuvem.Close
    Set mConexaoNuvem = Nothing
    Set gConnNuvem = Nothing
    Exit Sub

trata_erro:
End Sub


