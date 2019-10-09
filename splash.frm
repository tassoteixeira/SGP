VERSION 5.00
Begin VB.Form splash 
   Caption         =   "Sistema Gerenciador de Posto"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_mensagem 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1860
      TabIndex        =   0
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   0
      Top             =   1260
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   2160
   End
   Begin VB.Label lblVersao 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   6600
      Left            =   0
      Picture         =   "splash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8400
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x_tempo As Integer
Dim x_tempo2 As Integer
Dim x_flag As Boolean
Dim lConexaoSgoNuvem As New ADODB.Connection
Dim lConnConfiguracao As New ADODB.Connection

Dim lSQL As String

Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private dados As New cDados
Private Empresa As New cEmpresa
Private Sub SegurancaParaLocacao()
Dim lNumeroHd As String
Dim lNumeroLAN As String
Dim xString As String
Dim xNomeEmpresa As String
    ' Máscara da licença é:
    ' 9999      -> Ano da 1a Licença
    ' 9999      -> Número sequencial de Licença
    ' 99        -> Mes da 1a Licença
    ' 99        -> Dia da 1a Licença
    ' 999999999 -> Número de Série do HD
    lNumeroHd = DriveSerial(Left("C:", 1))
'    lNumeroLAN = GetMACAddress
    
    'MsgBox "->" & lNumeroHD & "<-"
    'MsgBox Len(lNumeroHD)
    
    If Not dados.LocalizarCodigo(1) Then
        MsgBox "Erro ao localizar dados.", vbInformation, "Erro de Integridade"
    End If
    
    If dados.Empresa2 <> 0 Then
        Call GravaAuditoria(1, Me.name, 27, "Locação Vencida")
        Dim xMensagem As String
        xMensagem = "A versão do executável do SGP está incompatível com o banco de dados atual." & Chr(13) & Chr(13)
        xMensagem = xMensagem & "O sistema não irá funcionar até que todos os computadores estejam com a versão atual do SGP." & Chr(13) & Chr(13)
        xMensagem = xMensagem & "Este procedimento irá impedir que seja danificada a integridade relacional do banco de dados." & Chr(13)
        xMensagem = xMensagem & "Entre em contato URGENTEMENTE com o suporte técnico para que seja atualizado o sistema." & Chr(13)
        xMensagem = xMensagem & "Sn. 2000-0037-12-01-" & lNumeroHd & Chr(13) & Chr(13)
        xMensagem = xMensagem & "Telefone do Suporte: (62) 3277-1017" & Chr(13) & Chr(13)
        xMensagem = xMensagem & "Cerrado Tecnologia - Soluções Inteligentes."
        MsgBox xMensagem, vbCritical, "Versão Incompatível com o Banco de Dados."
        End
    End If
    
    ''MsgBox "-->" & GetMACAddress & "<--"
    'If Not VerificaPlacaRede Then
    '    dados.Empresa2 = 9
    '    If Not dados.Alterar(1) Then
    '        MsgBox "Erro ao atualizar dados.", vbInformation, "Erro de Integridade"
    '    End If
    '    MsgBox "Este programa não está licenciado para esta empresa." & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroHd & Chr(34) & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroLAN & Chr(34), vbCritical, "Atenção! Pirataria é Crime."
    '    End
    'End If
    
    'Verifica HD
    If Not VerificaHD Then
        'Verifica se a Data Demonstração está 10 dias antes de travar
        'e avisa por email
        If DateDiff("d", Date, gDataDemonstracao) <= 10 Then
            xString = "Computador:" & GetIPHostName() & " - " & GetIPAddress() & vbCrLf
            xString = xString & "N.Serie=" & gNumeroHd & vbCrLf
            xString = xString & "Data da Demonstração=" & fMascaraData(gDataDemonstracao) & vbCrLf
            xNomeEmpresa = BuscaEmpresa
            Call EnviaMensagemEmail(1, xNomeEmpresa, "TRAV.EM " & str(DateDiff("d", Date, gDataDemonstracao)) & " DIAS", xString, True, 0)
        End If

        Call GravaAuditoria(1, Me.name, 27, "Falha de Registro! N.Serie=" & gNumeroHd)
        If ConfiguracaoDiversa.LocalizarCodigo(1, "Internet Banda Larga") Then
            If ConfiguracaoDiversa.Verdadeiro Then
                xString = "Computador:" & GetIPHostName() & " - " & GetIPAddress() & vbCrLf
                xString = xString & "N.Serie=" & gNumeroHd & vbCrLf
                xNomeEmpresa = BuscaEmpresa
                Call EnviaMensagemEmail(1, xNomeEmpresa, "Falha de Registro!", xString, True, 0)
            End If
        End If
    End If
    
    'Verifica se a gDataLimiteUso está 10 dias antes de travar
    'e avisa por email
    If DateDiff("d", Date, gDataLimiteUso) <= 10 Then
        xString = "Computador:" & GetIPHostName() & " - " & GetIPAddress() & vbCrLf
        xString = xString & "N.Serie=" & gNumeroHd & vbCrLf
        xString = xString & "Data Limite de Uso=" & fMascaraData(gDataLimiteUso) & vbCrLf
        xNomeEmpresa = BuscaEmpresa
        Call EnviaMensagemEmail(1, xNomeEmpresa, "TRAV.EM " & str(DateDiff("d", Date, gDataLimiteUso)) & " DIAS", xString, True, 0)
    End If
  
    
    'Data Demonstração
    If Date < gDataDemonstracao Then
        Exit Sub
    End If
    
    'Data Limite de Uso
    If Date > gDataLimiteUso Then
        dados.Empresa2 = 9
        If Not dados.Alterar(1) Then
            Call GravaAuditoria(1, Me.name, 27, "Erro ao Bloquear. DataLimite=" & Format(gDataLimiteUso, "dd/mm/yyyy"))
            MsgBox "Erro ao atualizar dados.", vbInformation, "Erro de Integridade"
        End If
        Call GravaAuditoria(1, Me.name, 27, "Bloqueado. DataLimite=" & Format(gDataLimiteUso, "dd/mm/yyyy"))
        MsgBox "A locação deste programa está vencida!" & Chr(13) & "Efetue o pagamento e entre em contato com o suporte técnico." & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroHd, vbCritical, "Atenção! A Locação está ATRASADA."
        End
    End If
    
    'Verifica HD
    If Not VerificaHD Then
        dados.Empresa2 = 9
        If Not dados.Alterar(1) Then
            Call GravaAuditoria(1, Me.name, 27, "Erro ao Bloquear. N.Serie=" & gNumeroHd)
            MsgBox "Erro ao atualizar dados.", vbInformation, "Erro de Integridade"
        End If
        Call GravaAuditoria(1, Me.name, 27, "Bloqueado. N.Serie=" & gNumeroHd)
        MsgBox "Este programa não está licenciado para esta empresa." & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroHd & Chr(34), vbCritical, "Atenção! Pirataria é Crime."
        End
    End If

End Sub
Private Sub VerificaBancoSgpNuvem()
    Dim xString As String
    Dim xIpBanco As String
    Dim xCreateDatabase As String
    Dim xRegistroAfetado As Integer
    Dim xNomeBanco As String
    Dim xRsTabela As New ADODB.Recordset
    Dim xVersaoBanco As String
    
    On Error GoTo trata_erro
    
    xNomeBanco = "sgp_nuvem"
    If gNomeInternoBD <> "sgp_data" Then
        xNomeBanco = xNomeBanco & gNomeInternoBD
    End If
    xIpBanco = ReadINI("Local", "Nome do Banco de Dados", gArquivoIni)
    'xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco & ",4949"
    xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco '& gPortaBanco
    xString = xString & ";INITIAL CATALOG=;USER ID=sa;PASSWORD=" & gSenhaBD & ";"
    lConexaoSgoNuvem.ConnectionString = xString
    lConexaoSgoNuvem.Open
    If lConexaoSgoNuvem.State = 1 Then
        
        'Verifica versão do Banco
        xVersaoBanco = ""
        lSQL = "Select @@Version"
        xRsTabela.CursorLocation = adUseClient
        xRsTabela.Open lSQL, lConexaoSgoNuvem, adOpenForwardOnly, adLockReadOnly
        If xRsTabela.RecordCount > 0 Then
            If xRsTabela(0).Value Like "*2014*" Then
                xVersaoBanco = "SQL SERVER 2014"
            ElseIf xRsTabela(0).Value Like "*2000*" Then
                xVersaoBanco = "MSDE 2000"
            ElseIf xRsTabela(0).Value Like "*2008 R2*" Then
                xVersaoBanco = "SQL SERVER 2008 R2"
            ElseIf xRsTabela(0).Value Like "*2012*" Then
                xVersaoBanco = "SQL SERVER 2012"
            Else
                xVersaoBanco = xRsTabela(0).Value
            End If
        End If
        xRsTabela.Close
        
        'Verifica se banco existente
        If xVersaoBanco = "MSDE 2000" Then
            lSQL = "SELECT CATALOG_NAME"
            lSQL = lSQL & " From INFORMATION_SCHEMA.SCHEMATA"
            lSQL = lSQL & " WHERE SCHEMA_NAME = " & preparaTexto("dbo")
            lSQL = lSQL & " AND CATALOG_NAME = " & preparaTexto(xNomeBanco)
        ElseIf xVersaoBanco = "SQL SERVER 2008 R2" Then
            lSQL = "SELECT name"
            lSQL = lSQL & " FROM sys.databases"
            lSQL = lSQL & "  WHERE name = " & preparaTexto(xNomeBanco)
        ElseIf xVersaoBanco = "SQL SERVER 2012" Then
            lSQL = "SELECT name"
            lSQL = lSQL & " FROM sys.databases"
            lSQL = lSQL & "  WHERE name = " & preparaTexto(xNomeBanco)
        ElseIf UCase(xVersaoBanco) Like "*SQL SERVER 2014*" Then
            lSQL = "SELECT name"
            lSQL = lSQL & " FROM sys.databases"
            lSQL = lSQL & "  WHERE name = " & preparaTexto(xNomeBanco)
        Else
            lSQL = "SELECT CATALOG_NAME"
            lSQL = lSQL & " From INFORMATION_SCHEMA.SCHEMATA"
            lSQL = lSQL & " WHERE SCHEMA_NAME = " & preparaTexto("dbo")
            lSQL = lSQL & " AND CATALOG_NAME = " & preparaTexto(xNomeBanco)
        End If
        xRsTabela.CursorLocation = adUseClient
        xRsTabela.Open lSQL, lConexaoSgoNuvem, adOpenForwardOnly, adLockReadOnly
        If xRsTabela.RecordCount = 0 Then
            'Cria Banco
            xCreateDatabase = "CREATE DATABASE " & xNomeBanco & " ON "
            xCreateDatabase = xCreateDatabase & "( NAME = " & xNomeBanco & ", FILENAME = 'C:\Cerrado.Net\Sgp\dataMSSQL$SGP_DATA\Data\" & xNomeBanco & ".mdf', SIZE = 10MB, MAXSIZE = 50MB, FILEGROWTH = 10MB )"
            xCreateDatabase = xCreateDatabase & " LOG ON "
            xCreateDatabase = xCreateDatabase & "( NAME = " & xNomeBanco & "_log, FILENAME = 'C:\Cerrado.Net\Sgp\dataMSSQL$SGP_DATA\Data\" & xNomeBanco & "_log.mdf', SIZE = 5MB, MAXSIZE = 25MB, FILEGROWTH = 5MB )"
            lConexaoSgoNuvem.Execute xCreateDatabase, xRegistroAfetado, adCmdText + adExecuteNoRecords
            If xRegistroAfetado = -1 Then
                xRsTabela.Close

                MsgBox "Banco " & xNomeBanco & " criado com sucesso.", vbInformation + vbOKOnly, "BD Criado!"
                lConexaoSgoNuvem.Close
                'xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco & ",4949"
                xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco
                xString = xString & ";INITIAL CATALOG=" & xNomeBanco & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
                lConexaoSgoNuvem.ConnectionString = xString
                lConexaoSgoNuvem.Open
                Call CriaTabelasSgpNuvem(xNomeBanco, "VersaoBancoDados")
                Call CriaTabelasSgpNuvem(xNomeBanco, "IntegracaoNuvem")
            Else
                xRsTabela.Close
            End If
        Else
            
            
            lConexaoSgoNuvem.Close
            
            'xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco & ",4949"
            xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco
            xString = xString & ";INITIAL CATALOG=" & xNomeBanco & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            lConexaoSgoNuvem.ConnectionString = xString
            lConexaoSgoNuvem.Open
            
            
            lSQL = "SELECT name FROM sysobjects WHERE type = 'U' AND name = 'MovimentoCat52'"
            xRsTabela.Open lSQL, lConexaoSgoNuvem, adOpenForwardOnly, adLockReadOnly
            If xRsTabela.RecordCount = 0 Then
                Call CriaTabelasSgpNuvem(xNomeBanco, "MovimentoCat52")
            End If
        End If
    Else
        MsgBox "Não foi possível conectar ao SGBD", vbCritical + vbInformation, "Erro com o SGBD"
        Exit Sub
    End If
    lConexaoSgoNuvem.Close
    Exit Sub
    
trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub CriaTabelasSgpNuvem(ByVal pNomeBanco As String, ByVal pNomeTabela As String)
    Dim xRsTabela As New ADODB.Recordset
    
    'Verifica se Tabela Existe
    lSQL = "SELECT TABLE_NAME"
    lSQL = lSQL & " From INFORMATION_SCHEMA.Tables"
    lSQL = lSQL & " WHERE TABLE_TYPE = " & preparaTexto("BASE TABLE")
    lSQL = lSQL & " AND TABLE_CATALOG = " & preparaTexto(pNomeBanco)
    lSQL = lSQL & " AND TABLE_NAME = " & preparaTexto(pNomeTabela)
    xRsTabela.CursorLocation = adUseClient
    xRsTabela.Open lSQL, lConexaoSgoNuvem, adOpenForwardOnly, adLockReadOnly
    If xRsTabela.RecordCount = 0 Then
        'Cria Tabela
        If pNomeTabela = "VersaoBancoDados" Then
            lSQL = "CREATE TABLE VersaoBancoDados ("
            lSQL = lSQL & "    Versao VarChar(8) Not Null,"
            lSQL = lSQL & "    Ordem SmallInt nOT Null,"
            lSQL = lSQL & "    [String de Comando] Text Not Null,"
            lSQL = lSQL & "    Descricao VarChar(150) Not Null,"
            lSQL = lSQL & "    Atualizado Bit Not Null"
            lSQL = lSQL & "    )"
            lSQL = lSQL & " ;"
            lSQL = lSQL & " CREATE UNIQUE CLUSTERED INDEX idVersao On VersaoBancoDados"
            lSQL = lSQL & "       ("
            lSQL = lSQL & "       Versao,"
            lSQL = lSQL & "       Ordem"
            lSQL = lSQL & "       )"
        ElseIf pNomeTabela = "IntegracaoNuvem" Then
            lSQL = "CREATE TABLE IntegracaoNuvem ("
            lSQL = lSQL & "    Empresa SmallInt Not Null,"
            lSQL = lSQL & "    Data DateTime Not Null,"
            lSQL = lSQL & "    [Nome da Tabela] VarChar(100) Not Null,"
            lSQL = lSQL & "    [Chave de Acesso] VarChar(100) Not Null,"
            lSQL = lSQL & "    [Tipo de Operacao] VarChar(30) Not Null,"
            lSQL = lSQL & "    [Integrado Em] DateTime Null"
            lSQL = lSQL & "    )"
            lSQL = lSQL & " ;"
            lSQL = lSQL & " CREATE UNIQUE CLUSTERED INDEX idData On IntegracaoNuvem"
            lSQL = lSQL & "       ("
            lSQL = lSQL & "       Empresa,"
            lSQL = lSQL & "       Data,"
            lSQL = lSQL & "       [Nome da Tabela],"
            lSQL = lSQL & "       [Chave de Acesso],"
            lSQL = lSQL & "       [Tipo de Operacao]"
            lSQL = lSQL & "       )"
        ElseIf pNomeTabela = "MovimentoCat52" Then
            lSQL = "CREATE TABLE MovimentoCat52 ("
            lSQL = lSQL & "   Empresa SmallInt Not Null,"
            lSQL = lSQL & "   [Codigo da ECF] SmallInt Not Null,"
            lSQL = lSQL & "   Data DateTime Not Null,"
            lSQL = lSQL & "   [Nome do Arquivo] VarChar(100) Not Null,"
            lSQL = lSQL & "   Imagem Image Null"
            lSQL = lSQL & "   )"
            lSQL = lSQL & ";"
            lSQL = lSQL & "CREATE UNIQUE NONCLUSTERED INDEX idCodigo On MovimentoCat52"
            lSQL = lSQL & "      ("
            lSQL = lSQL & "      Empresa,"
            lSQL = lSQL & "      [Codigo da ECF],"
            lSQL = lSQL & "      Data"
            lSQL = lSQL & "      )"
        End If
        lConexaoSgoNuvem.Execute lSQL, xRegistroAfetado, adCmdText + adExecuteNoRecords
        If xRegistroAfetado = -1 Then
            MsgBox "Tabela " & pNomeTabela & " criado com sucesso.", vbInformation + vbOKOnly, "Tabela Criada!"
            If pNomeTabela = "VersaoBancoDados" Then
                InsereRegistrosVersaoBancoDados
            End If
        End If
    End If
    xRsTabela.Close
End Sub
Private Function BuscaEmpresa() As String
    BuscaEmpresa = ""
    If Empresa.LocalizarCodigo(1) Then
        BuscaEmpresa = Empresa.Nome
    Else
        If Empresa.LocalizarUltimo Then
            BuscaEmpresa = Empresa.Nome
        End If
    End If
End Function
Private Sub InsereRegistrosVersaoBancoDados()
    'Insere Registro 'VersaoBancoDados
    lSQL = "INSERT INTO VersaoBancoDados VALUES ( "
    lSQL = lSQL & preparaTexto("20110714") & ", "
    lSQL = lSQL & "1, "
    
    lSQL = lSQL & "'CREATE TABLE VersaoBancoDados ("
    lSQL = lSQL & "    Versao VarChar(8) Not Null,"
    lSQL = lSQL & "    Ordem SmallInt nOT Null,"
    lSQL = lSQL & "    [String de Comando] Text Not Null,"
    lSQL = lSQL & "    Descricao VarChar(150) Not Null,"
    lSQL = lSQL & "    Atualizado Bit Not Null"
    lSQL = lSQL & "    )"
    lSQL = lSQL & " ;"
    lSQL = lSQL & " CREATE UNIQUE CLUSTERED INDEX idVersao On VersaoBancoDados"
    lSQL = lSQL & "       ("
    lSQL = lSQL & "       Versao,"
    lSQL = lSQL & "       Ordem"
    lSQL = lSQL & "       )', "
    lSQL = lSQL & preparaTexto("Cria tabela VersaoBancoDados") & ", "
    lSQL = lSQL & preparaBooleano(False) & " )"
    
    lConexaoSgoNuvem.Execute lSQL, xRegistroAfetado, adCmdText + adExecuteNoRecords
    If xRegistroAfetado > 0 Then
        MsgBox "Registro inserido com sucesso", vbInformation + vbOKOnly, "Registro Gravado!"
    End If
    
    
    'Insere Registro 'IntegracaoNuvem
    lSQL = "INSERT INTO VersaoBancoDados VALUES ( "
    lSQL = lSQL & preparaTexto("20110714") & ", "
    lSQL = lSQL & "2, "
    
    lSQL = lSQL & "'CREATE TABLE IntegracaoNuvem ("
    lSQL = lSQL & "    Empresa SmallInt Not Null,"
    lSQL = lSQL & "    Data DateTime Not Null,"
    lSQL = lSQL & "    [Nome da Tabela] VarChar(100) Not Null,"
    lSQL = lSQL & "    [Chave de Acesso] VarChar(100) Not Null,"
    lSQL = lSQL & "    [Tipo de Operacao] VarChar(30) Not Null,"
    lSQL = lSQL & "    [Integrado Em] DateTime Null"
    lSQL = lSQL & "    )"
    lSQL = lSQL & " ;"
    lSQL = lSQL & " CREATE UNIQUE CLUSTERED INDEX idData On IntegracaoNuvem"
    lSQL = lSQL & "       ("
    lSQL = lSQL & "       Empresa,"
    lSQL = lSQL & "       Data,"
    lSQL = lSQL & "       [Nome da Tabela],"
    lSQL = lSQL & "       [Chave de Acesso],"
    lSQL = lSQL & "       [Tipo de Operacao]"
    lSQL = lSQL & "       )', "
    lSQL = lSQL & preparaTexto("Cria tabela IntegracaoNuvem") & ", "
    lSQL = lSQL & preparaBooleano(False) & " )"
    
    lConexaoSgoNuvem.Execute lSQL, xRegistroAfetado, adCmdText + adExecuteNoRecords
    If xRegistroAfetado > 0 Then
        MsgBox "Registro inserido com sucesso", vbInformation + vbOKOnly, "Registro Gravado!"
    End If
End Sub
Function AccessPassword(ByVal Filename As String) As String

    Dim MaxSize, NextChar, MyChar, secretpos, TempPwd
    Dim secret(13)

    secret(0) = (&H86)
    secret(1) = (&HFB)
    secret(2) = (&HEC)
    secret(3) = (&H37)
    secret(4) = (&H5D)
    secret(5) = (&H44)
    secret(6) = (&H9C)
    secret(7) = (&HFA)
    secret(8) = (&HC6)
    secret(9) = (&H5E)
    secret(10) = (&H28)
    secret(11) = (&HE6)
    secret(12) = (&H13)

    secretpos = 0

    Open Filename For Input As #1 ' Abre o arquivo para escrita


    For NextChar = 67 To 79 Step 1 'Le a senha criptografada
        Seek #1, NextChar ' define a posição
        MyChar = Input(1, #1) ' Lê o caractere.
        TempPwd = TempPwd & Chr(Asc(MyChar) Xor secret(secretpos)) 'Faz a decriptação
        secretpos = secretpos + 1 'incrementa o ponteiro
    Next NextChar

    Close #1 ' fecha o arquivo.
    AccessPassword = TempPwd

End Function
Private Sub Form_Load()
    Dim xNomeBancoDados As String
    Dim xArquivo As TextStream
    Dim xNomeBd_Porta As Variant
    Dim xNivelSenha As String
    
    bdAccess = False
    bdMySql = False
    bdSqlServer = False
    bdOracle = False
    
    lblVersao.Caption = "Versão " & gVersaoSGP
    
    'DESCOBRE SENHA DE ACCESS
    'Dim xteste As String
    'xteste = AccessPassword("C:\Arquivos de programas\Horus\Horus.mdb")
    
    xNomeBancoDados = ReadINI("SGBD", "Gerenciador de Banco de Dados", gArquivoIni)
    xNivelSenha = ReadINI("SGBD", "Senha", gArquivoIni)
    If xNomeBancoDados = "ACCESS" Then
        bdAccess = True
    ElseIf xNomeBancoDados = "MYSQL" Then
        bdMySql = True
    ElseIf xNomeBancoDados = "SQLSERVER" Then
        bdSqlServer = True
    ElseIf xNomeBancoDados = "ORACLE" Then
        bdOracle = True
    End If
    CentraForm Me
    x_tempo = 0
    x_tempo2 = 0
    x_flag = False
    gPortaBanco = ",4949"
    If bdSqlServer Then
        gIpBanco = ReadINI("Local", "Nome do Banco de Dados", gArquivoIni)
        gIpBanco = Mid(gIpBanco, 1, Len(gIpBanco) - 5)
        If VerificaConexoesMultiplas Then
            xNomeBd_Porta = Split(gNomeInternoBD, ",")
            If UBound(xNomeBd_Porta) = 1 Then
                gNomeInternoBD = xNomeBd_Porta(0)
                gPortaBanco = "," & xNomeBd_Porta(1)
            End If
            'Call WriteINI("Local", "Nome do Banco de Dados", gIpBanco & ",4949", gArquivoIni)
            Call WriteINI("Local", "Nome do Banco de Dados", gIpBanco & gPortaBanco, gArquivoIni)
            Call WriteINI("Local", "Nome Interno BD", gNomeInternoBD, gArquivoIni)
            If gArqTxt.FileExists("c:\CheqPosto.ini") Then
                Call WriteINI("Local", "Nome do Banco de Dados", gIpBanco & gPortaBanco, "c:\CheqPosto.ini")
                Call WriteINI("Local", "Nome Interno BD", gNomeInternoBD, "c:\CheqPosto.ini")
            End If
        End If
    End If
    gNomeUsuarioBD = "sa"
    gSenhaBD = "cerrado72"
    If gIpBanco Like "*cloudapp.net*" Or xNivelSenha = "padrao2" Then
        gSenhaBD = "cRr472#*$pst"
    End If
    
    If gIpBanco Like "*cerrado.database.windows.net*" Then
        gNomeUsuarioBD = "cerrado"
        gSenhaBD = "cRr472#*$pst"
        bdSqlServerAzure = True
    Else
        bdSqlServerAzure = False
    End If
    Timer1.Enabled = True
    Timer1.Interval = 900
End Sub

Private Sub Timer1_Timer()
    x_tempo = x_tempo + 1
    If x_tempo = 1 Then
        ChamaDrive
        
        'gDiretorioData = ":\vb5\sgp\data\"
        'ChDir "\VB5\SGP\DATA"
        Set Conectar = New CConexao
        Set ConectarNuvem = New cConexaoNuvem
        If Conectar.ConexaoAtiva Then
            Call GravaAuditoria(1, Me.name, 1, "S.G.P. " & gVersaoSGP)
        Else
            MsgBox "Não foi possível estabelecer a conexão com o banco de dados!" & vbCrLf + "Endereço de Conexão: " & Conectar.BaseDados, vbCritical + vbOKOnly, "Erro de Conexão!"
            End
        End If
    ElseIf x_tempo = 2 Then
        If bdAccess Then
            Set bd_sgp = OpenDatabase(gNomeInternoBD & ".MDB")
        End If
    ElseIf x_tempo = 3 Then
        If bdAccess Then
            'Set bd_sgp_b = OpenDatabase("SGP_DATA_BAIXA.MDB")
        End If
    ElseIf x_tempo = 4 Then
        If bdAccess Then
            'Set bd_sgp_m = OpenDatabase("SGP_DATA_MOVIMENTO.MDB")
            'Set ConectarM = New CConexaoM
            'Set ConectarB = New CConexaoB
            'cnnSGPb.Mode = adModeRead
            'Set cnnSGPb = New adodb.Connection
            'cnnSGPm.Mode = adModeRead
            'Set cnnSGPm = New adodb.Connection
        End If
        cnnSGP.Mode = adModeRead
        Set cnnSGP = New ADODB.Connection
        If bdAccess Then
            ''teste para acesso remoto
            'gConnectionString = "Provider=MS Remote;Remote Server=http://postoesmeralda.myvnc.com:1;Remote Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gDrive & gDiretorioData & gNomeBancoDados '"Sgp_data.mdb"
            
           ' cnnSGP.Open "Provider=MS Remote;" & _
           '"Remote Server=http://postoesmeralda.myvnc.com:2;" & _
           '"Remote Provider=Microsoft.Jet.OLEDB.4.0;" & _
           '"Data Source=" & gDrive & gDiretorioData & gNomeBancoDados, _
           ' "admin", ""
            'gConnectionString = "Provider=MS Remote.1;Data Source=C:\vb5\sgp\data\Sgp_Data.Mdb;User ID=admin;Remote Server=http://postoesmeralda.myvnc.com:2;Remote Provider=Microsoft.Jet.OLEDB.4.0;Internet Timeout=300000;Transact Updates=True"
            
            
            
            ''fim teste para acesso remoto
            gConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & gDrive & gDiretorioData & gNomeBancoDados '"Sgp_data.mdb"
            cnnSGP.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gDiretorioData & gNomeBancoDados & ";Uid=Admin;Pwd=;"
            'cnnSGPb.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gDiretorioData & "sgp_data_baixa.mdb;Uid=Admin;Pwd=;"
            'cnnSGPm.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gDiretorioData & "sgp_data_movimento.mdb;Uid=Admin;Pwd=;"
        'ElseIf bdOracle Then
        '    'conexao ORACLE
        '    'cnnSGP.Open "Provider=msdaora;Data Source=CLIENTE;User Id=agroquima;Password=agro;"
        End If
        
        
'        'teste de conexao com azure
'        Dim xConnAzure As New adodb.Connection
'        Dim xString As String
'
'        xString = "Driver={SQL Server Native Client 10.0};Server=tcp:dlq29esbhj.database.windows.net;Database=cerradotef;Uid=cerradotef@dlq29esbhj;Pwd=tlara87&*;Encrypt=yes;"
'        xString = "Provider=SQLNCLI10;Server=tcp:dlq29esbhj.database.windows.net;Database=cerradotef;Uid=cerradotef@dlq29esbhj; Pwd=tlara87&*;"
'        xString = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & gNomeBancoDados & ";INITIAL CATALOG=" & "sgp_data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
'
'        xConnAzure.ConnectionString = xString
'        xConnAzure.Open
'        xConnAzure.Close

    ElseIf x_tempo = 5 Then
        gDataDemonstracao = CDate("20/01/2013")
        gDataLimiteUso = CDate("20/05/2013")
        'gDataDemonstracao = CDate("10/01/2010")
        'gDataLimiteUso = CDate("10/01/2010")
        If Not bdSqlServerAzure Then
            Call VerificaBancoSgpNuvem
        End If
        'Call SegurancaParaLocacao
        If dados.LocalizarCodigo(1) Then
            'If Dados.Empresa2 <> 0 Then
            '    MsgBox "A locação deste programa está vencida!" & Chr(13) & "Efetue o pagamento e entre em contato com o suporte técnico." & Chr(13) & "Fone: (0xx62) 8436-4444 - Tasso Teixeira", vbCritical, "Atenção! A Locação está ATRASADA."
            '    End
            'End If
        End If
        Set ConfiguracaoDiversa = Nothing
        Set dados = Nothing
        Set Empresa = Nothing
        Unload Me
    End If
End Sub
Private Sub Timer2_Timer()
    Dim x_mensagem As String
    x_mensagem = Space(90) & "Aguarde! Abrindo banco de dados..."
    x_tempo2 = x_tempo2 + 1
    If x_tempo2 = 1 Then
        lbl_mensagem = x_mensagem
        txt_mensagem = x_mensagem
    Else
        If x_tempo2 <= Len(x_mensagem) Then
            lbl_mensagem = Mid(x_mensagem, x_tempo2, Len(x_mensagem) - x_tempo2)
            txt_mensagem = Mid(x_mensagem, x_tempo2, Len(x_mensagem) - x_tempo2)
        Else
            x_tempo2 = 0
        End If
    End If
End Sub



Function SetDbSenha(DBPath As String, novaSenha As String) As Boolean

    If Dir(DBPath) = "" Then
        Exit Function
    End If

Dim db As DAO.Database

On Error Resume Next

    Set db = OpenDatabase(DBPath, True)

    If Err.Number <> 0 Then
        Exit Function
    End If

    db.NewPassword "", novaSenha

    SetDbSenha = Err.Number = 0

    db.Close

End Function

Private Function VerificaConexoesMultiplas() As Boolean
    Dim rsTabela As New ADODB.Recordset
    Dim xNomeBanco As String
    Dim xNomeEmpresa As String
    Dim i As Integer
    
    On Error GoTo FileError

    VerificaConexoesMultiplas = False
'[CONEXOES]
'Conexao 001=Esmeralda|@|crresmeralda.ddns.com.br|@|False|@|sgp_data|@|
'Conexao 002=Rubi|@|crrrubi.ddns.com.br|@|False|@|sgp_data|@|
    
    Dim xString As String
    For i = 1 To 100
        xString = ReadINI("CONEXOES", "Conexao " & Format(i, "000"), gArquivoIni)
        If xString = "" Then
            Exit For
        End If
        If UCase(fRetiraString(xString, 3)) = "TRUE" Then
            VerificaConexoesMultiplas = True
            gIpBanco = fRetiraString(xString, 2)
            gNomeInternoBD = fRetiraString(xString, 4)
            Exit Function
        End If
        g_string = g_string & fRetiraString(xString, 2) & "|@|"
        g_string = g_string & fRetiraString(xString, 1) & "|@|"
    Next
    g_string = "Selecione a Empresa para Conexão!|@|" & i - 1 & "|@|" & g_string
    
    opcaoGeral.Show 1
    xNomeEmpresa = RetiraGString(2)
    g_string = ""
    
    
    For i = 1 To 100
        xString = ReadINI("CONEXOES", "Conexao " & Format(i, "000"), gArquivoIni)
        If xString = "" Then
            Exit For
        End If
        
        If fRetiraString(xString, 1) = xNomeEmpresa Then
            gIpBanco = fRetiraString(xString, 2)
            gNomeInternoBD = fRetiraString(xString, 4)
            VerificaConexoesMultiplas = True
            Exit For
        End If
    Next
    
    
    
    
    
'    'Verifica se Existe Coluna "Nome do Banco"
'    'Caso não exista, cria a mesma.
'    VerificaCriaColunaNomeDoBanco
'
'    'Abre Conexao com Banco Configuracao.Mdb
'    xNomeBanco = "C:\Cerrado.Net\Sgp\Data\Configuracao.Mdb"
'    lConnConfiguracao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & xNomeBanco
'    lConnConfiguracao.Open
'
'
'
'    'Localiza Ip da Empresa Automatica
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Endereco IP] AS IP, [Nome do Banco]"
'    lSQL = lSQL & "  FROM Conexao"
'    lSQL = lSQL & " WHERE Automatico = -1"
'    Set rsTabela = New adodb.Recordset
'    rsTabela.Open lSQL, lConnConfiguracao, adOpenForwardOnly, adLockReadOnly
'    If Not rsTabela.EOF Then
'        rsTabela.MoveFirst
'        VerificaConexoesMultiplas = True
'        'gIpBanco = rsTabela("IP").Value & ",4949"
'        gIpBanco = rsTabela("IP").Value
'        gNomeInternoBD = rsTabela("Nome do Banco").Value
'        rsTabela.Close
'        lConnConfiguracao.Close
'        Set rsTabela = Nothing
'        Set lConnConfiguracao = Nothing
'        Exit Function
'    Else
'        Set rsTabela = Nothing
'    End If
'
'
'    'Cria Relação das Empresas
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Nome da Empresa] AS Nome, [Endereco IP] AS IP"
'    lSQL = lSQL & "  FROM Conexao"
'    lSQL = lSQL & " ORDER BY [Nome da Empresa]"
'
'    'Abre RecordSet
'    rsTabela.Open lSQL, lConnConfiguracao, adOpenForwardOnly, adLockReadOnly
'
'    g_string = "Selecione a Empresa para Conexão!|@|"
'    If rsTabela.EOF = False Then
'        i = 0
'        rsTabela.MoveFirst
'        Do Until rsTabela.EOF
'            i = i + 1
'            rsTabela.MoveNext
'        Loop
'        g_string = g_string & i & "|@|"
'
'        rsTabela.MoveFirst
'        Do Until rsTabela.EOF
'            g_string = g_string & rsTabela("IP").Value & "|@|"
'            g_string = g_string & rsTabela("Nome").Value & "|@|"
'            rsTabela.MoveNext
'        Loop
'
'        opcaoGeral.Show 1
'        xNomeEmpresa = RetiraGString(2)
'        g_string = ""
'
'        rsTabela.Close
'        'Localiza Ip da Empresa Selecionada
'        lSQL = ""
'        lSQL = lSQL & "SELECT [Endereco IP] AS IP, [Nome do Banco]"
'        lSQL = lSQL & "  FROM Conexao"
'        lSQL = lSQL & " WHERE [Nome da Empresa] = " & preparaTexto(xNomeEmpresa)
'        rsTabela.Open lSQL, lConnConfiguracao, adOpenForwardOnly, adLockReadOnly
'        rsTabela.MoveFirst
'        If Not rsTabela.EOF Then
'            VerificaConexoesMultiplas = True
'            gIpBanco = rsTabela("IP").Value
'            gNomeInternoBD = rsTabela("Nome do Banco").Value
'        End If
'
'    End If
'
'    rsTabela.Close
'    lConnConfiguracao.Close
'    Set rsTabela = Nothing
'    Set lConnConfiguracao = Nothing
    Exit Function
FileError:
    If Err = 3044 Then
        MsgBox "Não foi possível localizar o banco de dados " & xNomeBanco, vbCritical, "Erro de Integridade!"
    Else
        MsgBox Error
    End If
    Exit Function
End Function

Private Function VerificaCriaColunaNomeDoBanco() As Boolean
    MsgBox "teste 1a"
    Dim rsTabela As New ADODB.Recordset
    Dim i As Integer
    Dim xString1 As String
    Dim xString2 As String
    Dim xRecordsAffected As Long
    Dim xNomeBanco As String
    Dim xConnConfiguracao As New ADODB.Connection
        
    On Error GoTo FileError

    MsgBox "teste 1b"
    
    xNomeBanco = "C:\Cerrado.Net\Sgp\Data\Configuracao.Mdb"
    xConnConfiguracao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & xNomeBanco
    xConnConfiguracao.Open
    
    MsgBox "teste 1c"
    
    
    lSQL = ""
    lSQL = lSQL & "SELECT *" '[Nome do Banco]"
    lSQL = lSQL & "  FROM Conexao"
    MsgBox "teste 1d"
    
    'Abre RecordSet
    rsTabela.Open lSQL, xConnConfiguracao, adOpenForwardOnly, adLockReadOnly
    MsgBox "teste 1e"
    If rsTabela.Fields.Count = 3 Then
        xString1 = "ALTER TABLE Conexao ADD COLUMN `Nome do Banco` VARCHAR(100) NOT NULL"
        'xString1 = "ALTER TABLE Conexao ADD COLUMN NomeDoBanco TEXT 100"
        xString2 = "UPDATE Conexao SET `Nome do Banco` = " & preparaTexto("sgp_data")
        rsTabela.Close
        Set rsTabela = Nothing
        xConnConfiguracao.Execute xString1
        xConnConfiguracao.Execute xString2, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    Else
        rsTabela.Close
        Set rsTabela = Nothing
    End If
    
    xConnConfiguracao.Close
    Set xConnConfiguracao = Nothing
    
    Exit Function
FileError:
    MsgBox Error
    Exit Function
End Function


