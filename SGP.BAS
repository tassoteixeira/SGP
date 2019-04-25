Attribute VB_Name = "sgp_variaveis_globais"
Option Explicit

Global bd_sgp As Database
Global bd_disco As Database
Global cnnSGP As New ADODB.Connection
Global gConnGic As New ADODB.Connection
Public Conectar As CConexao
Public ConectarNuvem As cConexaoNuvem
Public ConexaoAuxiliar As cConexaoAuxiliar
Public gNomeUsuarioBD As String
Public gSenhaBD As String
Public Const gQUANTIDADE_MAXIMA_BICO As Integer = 48


'habilita/desabilida CTRL+ALT+DEL e CTRL+ESC
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'Arquivo INI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SHOWDROPDOWN = &H14F

Global DefaultWorkspace As Workspace
Global tabela As TableDef
Global campo As Field
Global Indice As Index

Global tbl_acm_venda_combustivel As Table
Global tbl_baixa_duplicata_receber As Table
Global tbl_baixa_nota As Table
Global tbl_baixa_pagar As Table
Global tbl_banco As Table
Global tbl_bomba As Table
Global tbl_cheque As Table
Global tbl_cheque_baixado As Table
Global tbl_cheque_emitido As Table
Global tbl_cheque_folha As Table
Global tbl_cliente As Table
Global tbl_cliente_conveniado As Table
Global tbl_combustivel As Table
Global tbl_configuracao As Table
'Global tbl_conta As Table
Global tbl_conta_bancaria As Table
Global tbl_contas_pagar As Table
Global tbl_convenio As Table
Global tbl_dados As Table
Global tbl_duplicata_receber As Table
Global tbl_empresa As Table
Global tbl_entrada_comb As Table
Global tbl_estoque As Table
Global tbl_fornecedor As Table
Global tbl_funcionario As Table
Global tbl_historico As Table
Global tbl_local_cobranca As Table
Global tbl_med_comb As Table
Global tbl_menu As Table
Global tbl_mov_bomba As Table
Global tbl_movimento_nota As Table
Global tbl_mov_historico As Table
Global tbl_pagina_lmc As Table
Global tbl_pedido_comb As Table
Global tbl_produto As Table
Global tbl_programa As Table
Global tbl_saldo_conta As Table
Global tbl_tipodoc As Table
Global tbl_usuario As Table
Global tbl_vencimento As Table

Global g_nivel_acesso As Integer
Global g_lmc As Integer
Global g_usuario As Integer
Global g_empresa As Integer
Global g_nome_empresa As String
Global g_caixa_unificado As Boolean
Global g_cidade_empresa As String
Global gUfEmpresa As String
Global g_nome_conta As String
Global g_historico As Long
Global g_banco As Integer
Global g_tipodoc As Integer
Global g_tamanho_impressora As Integer
Global g_conta_bancaria As String * 10
Global g_data As Date
Global g_data_def As Date
Global g_valor As Currency
Public g_nome_tabela As String
Public g_nome_usuario As String
Public g_string As String
Public gStringChamada As String
Public gStringChamadaSangria As String
Public gDrive As String
Public gDiretorioData As String
Public gDiretorioAplicativo As String
Public gNomeBancoDados As String
Public g_impressora_matricial As Boolean
Public gTefString As String
Public gTefResposta As Boolean
Public gArqTxt As New FileSystemObject
Public gArquivoTXT As TextStream
Public gArquivoTMP As TextStream
Public gArquivoTMV As TextStream
Public gArquivoHTML As TextStream
Public g_automacao As Boolean
Public gQtdPeriodo As Integer
Public gDataDemonstracao As Date
Public gDataLimiteUso As Date
Public gNumeroHd As String
Public gNumeroEmailInicial As Long
Public gInternetBandaLarga As Boolean
Public gEmpresaGlobal As Integer
Public gUsuarioGlobal As Integer
Public gVersaoSGP As String
Public gArquivoIni As String
Public gTipoAmbienteNFCe As String


Public gSQL As String
Public bdAccess As Boolean
Public bdMySql As Boolean
Public bdSqlServer As Boolean
Public bdSqlServerAzure As Boolean
Public bdOracle As Boolean
Public gConnectionString As String
Public gIpBanco As String
Public gPortaBanco As String
Public gNomeInternoBD As String

Public g_cfg_data_i As Date
Public g_cfg_data_f As Date
Public g_cfg_empresa_i As Integer
Public g_cfg_empresa_f As Integer
Public g_cfg_periodo_i As String * 1
Public g_cfg_periodo_f As String * 1


Public gIDImpressoraPadrao As Byte
Public gNomeImpressoraSelecionada As String
Public gTipoImpressoraSelecionada As String
Public gCNPJEmpresa As String


Global Const ASC_AA = 65, ASC_ZZ = 90
Global Const ASC_a = 97, ASC_z = 122
Global Const ASC_0 = 48, ASC_9 = 57


Public Const ArqSgpIni As String = "c:\Cerrado.Net\sgp.ini"


'Início de Variáveis para Criar DSN
Private Const ODBC_ADD_DSN = 1 ' Add data source
Private Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3 ' Remove data source
Private Const vbAPINull As Long = 0& ' NULL Pointer


Const REG_SZ = 1
Const KEY_ALL_ACCESS = &H2003F
Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'API Command to create a Data Source Name, not used in this example
Public Declare Function SQLCreateDataSource Lib "ODBCCP32.DLL" (ByVal hWnd&, ByVal lpszDS$) As Boolean
'API to modify/Edit/Create a Data Source Name
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWnd As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
'Fim de Variáveis para Criar DSN

Public Sub AguardaMS(ByVal pMileSegundos As Integer)
    Sleep (pMileSegundos)
End Sub

Public Function PetromovelEstaAtivo() As Boolean
    Dim xMovAtividade As New cMovAtividadeProg
    
    PetromovelEstaAtivo = False
    
    If xMovAtividade.LocalizarUltimoTipoPrograma("", GetIPAddress, "PETROMOVEL_AUTO") Then
        If xMovAtividade.Tipo_MovAtividadeProgramaExterno = "EM_EXECUCAO" Then
            PetromovelEstaAtivo = True
        End If
    End If
    
End Function

Public Function AtualizaVariaveisGlobaisPetromovelAuto() As Boolean

    AtualizaVariaveisGlobaisPetromovelAuto = False
    Dim xMovAutomatizacao As New CadastroDLL.cMovAutomatizacao
    Dim xDados As String
        
   On Error GoTo FileError


    'Verifico antes se está ativo
'    If Not ComunicaPetromovelAutoBD Then
'        Exit Function
'    End If
    
    
    If Not PetromovelEstaAtivo Then
        Exit Function
    End If
    
    xDados = ""
    xDados = xDados & g_empresa & "|@|"
    xDados = xDados & g_nome_empresa & "|@|"
    xDados = xDados & g_cidade_empresa & "|@|"
    xDados = xDados & gEmpresaGlobal & "|@|"
    xDados = xDados & gCNPJEmpresa & "|@|"
    xDados = xDados & g_usuario & "|@|"
    xDados = xDados & g_nome_usuario & "|@|"
    xDados = xDados & g_nivel_acesso & "|@|"
    xDados = xDados & g_data_def & "|@|"
    xDados = xDados & g_lmc & "|@|"
    xDados = xDados & IIf(g_impressora_matricial, "true", "false") & "|@|"
    xDados = xDados & IIf(g_caixa_unificado, "true", "false") & "|@|"
    xDados = xDados & g_conta_bancaria & "|@|"
    xDados = xDados & g_nome_conta & "|@|"
    xDados = xDados & g_cfg_empresa_i & "|@|"
    xDados = xDados & g_cfg_empresa_f & "|@|"
    xDados = xDados & g_cfg_data_i & "|@|"
    xDados = xDados & g_cfg_data_f & "|@|"
    xDados = xDados & g_cfg_periodo_i & "|@|"
    xDados = xDados & g_cfg_periodo_f & "|@|"
    xDados = xDados & IIf(g_automacao, "true", "false") & "|@|"
    
    xMovAutomatizacao.NSU_MovAutomatizacao = 0
    xMovAutomatizacao.IdEstabelecimento_MovAutomatizacao = g_empresa
    xMovAutomatizacao.DataSolicitacao_MovAutomatizacao = Now
    xMovAutomatizacao.Operacao_MovAutomatizacao = "ATUALIZA_VARIAVEIS_GLOBAIS"
    xMovAutomatizacao.Dados_MovAutomatizacao = xDados
    xMovAutomatizacao.IPComputadorAC_MovAutomatizacao = GetIPAddress()
    xMovAutomatizacao.IPInternetAC_MovAutomatizacao = "200.???.???.???"
    xMovAutomatizacao.VersaoAC_MovAutomatizacao = gVersaoSGP
    xMovAutomatizacao.VersaoHost_MovAutomatizacao = ""
    xMovAutomatizacao.HoraAnalise_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.EtapaAndamento_MovAutomatizacao = 1
    xMovAutomatizacao.DescricaoAndamento_MovAutomatizacao = "GRAVADO"
    xMovAutomatizacao.HoraImportacaoDados_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.HoraImportacaoComplementar_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.HoraConciliacaoDados_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.HoraCancelamentoHost_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.HoraConclusao_MovAutomatizacao = CDate("00:00:00")
    xMovAutomatizacao.MensagemRetorno_MovAutomatizacao = ""
    xMovAutomatizacao.Observacao_MovAutomatizacao = ""
    
    AtualizaVariaveisGlobaisPetromovelAuto = xMovAutomatizacao.Incluir
    
    Exit Function
    
FileError:
    CriaLogSGP "AtualizaVariaveisGlobaisPetromovelAuto: Erro ao tentar gravar movautomatizacao", Error, ""

End Function
Function ComunicaPetromovelAutoBD() As Boolean
    Dim xTempo As Integer
    Dim xHoraInicial As Date
    Dim xNSU As Long
    Dim xMovSolicitacao As New CadastroDLL.cMovSolicitacaoFuncaoNFe

    On Error GoTo FileError

    ComunicaPetromovelAutoBD = False
    
    xTempo = 10
    
    xNSU = GravaSolicitacaoComunicacaoPetromovelAuto
    If xNSU = 0 Then
        'MsgBox "Não foi possível incluir registro de Solicitação de comunicação com PetromovelAuto!", vbCritical, "Erro de Integridade"
        Exit Function
    End If
    
    If Not AguardaSolicitacaoPetromovelAuto(xTempo, xNSU) Then
        Exit Function
    End If
    
    If xMovSolicitacao.LocalizarNSU(g_empresa, xNSU) Then
        If xMovSolicitacao.HoraAnalise_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            If xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe = "PETROMOVEL_ATIVO" And xMovSolicitacao.Mensagem_MovSolicitacaoFuncaoNFe = "OK" Then
                ComunicaPetromovelAutoBD = True
            Else
                Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Operação desconhecida!" & vbCrLf & "TipoOperação=" & xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe, "")
                MsgBox "Operação desconhecida!" & vbCrLf & "TipoOperação=" & xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe, vbCritical, "Erro de Comando!"
            End If
        Else
            Call xMovSolicitacao.DefineHoraCancelamentoHost(xNSU, Now, gVersaoSGP)
        End If
    Else
        Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Não foi possível localizar a Solicitação de comunicação com o programa PetromovelAuto!", "")
        MsgBox "Não foi possível localizar a Solicitação de comunicação com o programa PetromovelAuto!" & vbCrLf & "NSU=" & xNSU, vbCritical, "Solicitação Inexistente!"
    End If
    
    Exit Function

FileError:
    Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Erro ao tentar comunicação com o programa PetromovelAuto - Erro=" & Err.Number & " - " & Err.Description, "")
    'MsgBox "Erro ao tentar comunicação com o programa PetromovelAuto!", vbCritical, "Erro na emissão de NFCe!"
    Exit Function
End Function

Public Function GravaSolicitacaoComunicacaoPetromovelAuto() As Long
    
    GravaSolicitacaoComunicacaoPetromovelAuto = 0
    
    On Error GoTo TrataError
    
    Dim xMovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
    
    xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    xMovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    xMovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = "PETROMOVEL_ATIVO"
    xMovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    xMovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = "0" 'Verficar se pode ser utilizado este número
    xMovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    xMovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    xMovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    xMovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    xMovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    xMovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    xMovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
    
    If xMovSolicitacaoFuncaoNFe.Incluir Then
        GravaSolicitacaoComunicacaoPetromovelAuto = xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe
    End If
    
    Exit Function
TrataError:
    Call CriaLogSGP("[GravaSolicitacaoComunicacaoPetromovelAuto]", "Erro ao tentar gravar solicitação de comunicação com PetromovelAuto - Err.Description=" & Err.Description, "")
    MsgBox "Não foi possível incluir registro de Solicitação de comunicação com PetromovelAuto!", vbCritical, "Erro de Integridade"
End Function



Public Function AguardaSolicitacaoPetromovelAuto(ByVal pTempoEspera As Integer, ByVal pNSU As Long) As Boolean
    Dim xHoraInicial As Date
    Dim i As Integer
    Dim xMovSolicitacao As New CadastroDLL.cMovSolicitacaoFuncaoNFe

    AguardaSolicitacaoPetromovelAuto = False
    'lbl_mensagem.Caption = "Aguarde... Verificando Serviço PetromovelAuto."
    'DoEvents
    
    xHoraInicial = Time
    'Fica o tempo definido por parâmetro
    Do Until DateDiff("s", xHoraInicial, Time) >= pTempoEspera
        If xMovSolicitacao.VerificaSeEstaEmAnalise(pNSU) Then
            AguardaSolicitacaoPetromovelAuto = True
            'lbl_mensagem.Caption = xMovSolicitacao.BuscaMensagem(pNSU)
            'DoEvents
            Exit Do
        End If
        'lbl_mensagem.Caption = xMovSolicitacao.BuscaMensagem(pNSU)
        'DoEvents
        Call AguardaMS(500)
    Loop
    If AguardaSolicitacaoPetromovelAuto = False Then
        If xMovSolicitacao.DefineHoraCancelamentoAC(pNSU, Time) Then
            'MsgBox "Tempo de solicitação de Serviço do PetromovelAuto excedido.", vbCritical, "Tempo Excedido!"
            Call CriaLogSGP("[AguardaSolicitacaoPetromovelAuto]", "Tempo de solicitação de Serviço do PetromovelAuto excedido.", "")
            'lbl_mensagem.Caption = "Tempo de solicitação de Serviço do PetromovelAuto excedido."
            DoEvents
        Else
            'MsgBox "Não será possível definir cancelamento de Solicitação pela AC.", vbCritical, "Erro de Integridade!"
            Call CriaLogSGP("[AguardaSolicitacaoPetromovelAuto]", "Não será possível definir cancelamento de Solicitação pela AC.", "")
            'lbl_mensagem.Caption = "Não será possível definir cancelamento de Solicitação pela AC."
            'DoEvents
        End If
    End If
End Function


Public Function CriaDSN() As Boolean
    Dim intRet As Long
    Dim strDriver As String
    Dim strAttributes As String

    
    
    strDriver = "SQL Server"
    strAttributes = "DSN=sgp_data" & Chr$(0)
    intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)

    If intRet Then
        MsgBox "DSN Eliminado"
    Else
        MsgBox "Fallo en el borrado"
    End If









    
    'strAttributes = "SERVER=127.0.0.1,4949" & Chr$(0)
    strAttributes = "SERVER=127.0.0.1" & gPortaBanco & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=Temp DSN modificado" & Chr$(0)
    strAttributes = strAttributes & "DSN=DSN_TEMP" & Chr$(0)
    strAttributes = strAttributes & "DATABASE=sgp_data" & Chr$(0)
    strAttributes = strAttributes & "UID=sa" & Chr$(0)
    strAttributes = strAttributes & "PWD=" & gSenhaBD & Chr$(0)
    'Para mostrar el diálogo usar Form1.Hwnd en vez de vbAPINull.
    intRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_DSN, strDriver, strAttributes)
    If intRet Then
        MsgBox "DSN Modificado"
    Else
        MsgBox "Fallo en la modificacion"
    End If

    'Para crear un DSN :
    
    'usamos el driver de SQL Server porque es el mas comun
    strDriver = "SQL Server"
    'Asignamos los parametros separados por null.
    'strAttributes = "SERVER=127.0.0.1,4949" & Chr$(0)
    strAttributes = "SERVER=127.0.0.1" & gPortaBanco & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=sgp_data teste" & Chr$(0)
    strAttributes = strAttributes & "DSN=sgp_data" & Chr$(0)
    strAttributes = strAttributes & "DATABASE=sgp_data" & Chr$(0)
    strAttributes = strAttributes & "UID=sa" & Chr$(0)
    strAttributes = strAttributes & "PWD=" & gSenhaBD & Chr$(0)
    'Para mostrar el diálogo usar Form1.Hwnd en vez de vbAPINull.
    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
    If intRet Then
        MsgBox "DSN Creado"
    Else
        MsgBox "Fallo en la creación"
    End If





End Function
Function SelecionaImpressoraEpson(ByVal x_form As Form) As Boolean
    g_tamanho_impressora = 132
    SelecionaImpressoraEpson = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) Like "*EPSON*" And Not UCase(Printer.DeviceName) Like "*DUPLICATA*" Then
        g_impressora_matricial = False 'True Alex - desabilitando o uso de matricial
        Exit Function
    End If
    x_form.CommonDialog1.ShowPrinter
    x_form.CommonDialog1.PrinterDefault = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) <> "*EPSON*" And Not UCase(Printer.DeviceName) Like "*DUPLICATA*" Then
        g_impressora_matricial = False
        Exit Function
    End If
    MsgBox "Não será possível imprimir na impressora " & Printer.DeviceName & ".", vbInformation
    SelecionaImpressoraEpson = False
End Function
Public Sub SelecionaImpressoraPadrao(ByVal pTitulo As String, ByVal pMensagem As String, ByVal pSuperior As Currency, ByVal pEsquerda As Currency, ByVal pLargura As Currency, ByVal pAltura As Currency)
    Dim xFrmSelecionaImpressora As New frm_seleciona_impressora
    
    Call xFrmSelecionaImpressora.MostraMensagens(pTitulo, pMensagem, pSuperior, pEsquerda, pLargura, pAltura)
    xFrmSelecionaImpressora.Show 1
    Set xFrmSelecionaImpressora = Nothing
End Sub
Function SelecionaImpressoraEpsonDuplicata(ByVal x_form As Form) As Boolean
    g_tamanho_impressora = 132
    SelecionaImpressoraEpsonDuplicata = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) Like "*EPSON*" And UCase(Printer.DeviceName) Like "*DUPLICATA*" Then
        g_impressora_matricial = True
        Exit Function
    End If
    x_form.CommonDialog1.ShowPrinter
    x_form.CommonDialog1.PrinterDefault = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) <> "*EPSON*" And UCase(Printer.DeviceName) Like "*DUPLICATA*" Then
        g_impressora_matricial = False
        Exit Function
    End If
    MsgBox "Não será possível imprimir na impressora " & Printer.DeviceName & ".", vbInformation
    SelecionaImpressoraEpsonDuplicata = False
End Function
Function SelecionaImpressoraHP(ByVal x_form As Form) As Boolean
    SelecionaImpressoraHP = True
    g_impressora_matricial = True
    If UCase(Printer.DeviceName) Like "*HP*" Or UCase(Printer.DeviceName) Like "*SAMSUNG*" Then
        g_impressora_matricial = False
        Exit Function
    End If
    x_form.CommonDialog1.ShowPrinter
    x_form.CommonDialog1.PrinterDefault = True
    If UCase(Printer.DeviceName) Like "*HP*" Or UCase(Printer.DeviceName) Like "*SAMSUNG*" Then
        g_impressora_matricial = False
        Exit Function
    End If
    'MsgBox "Não será possível imprimir na impressora " & Printer.DeviceName & ".", vbInformation
    'SelecionaImpressoraHP = False
End Function
Public Function ReadINI(Section As String, Key As String, Filename As String) As String
'Filename=nome do arquivo ini
'section=O que esta entre []
'key=nome do que se encontra antes do sinal de igual
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), Filename)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
Public Sub WriteINI(Section As String, Key As String, Text As String, Filename As String)
'Filename=nome do arquivo ini
'section=O que esta entre []
'key=nome do que se encontra antes do sinal de igual
'text= valor que vem depois do igual
    WritePrivateProfileString Section, Key, Text, Filename
End Sub
Private Function verificaCaracterEspecial(ByVal xString As String) As String
Dim i As Integer
    verificaCaracterEspecial = ""
    For i = 1 To Len(xString)
        verificaCaracterEspecial = verificaCaracterEspecial & Mid(xString, i, 1)
        If Mid(xString, i, 1) = "'" Then
            verificaCaracterEspecial = verificaCaracterEspecial & "'"
        End If
    Next
End Function
    Function fUltimoDiaMes(ByVal pData As Date) As Integer
        Dim xData As String
        Dim xDia As Integer
        
        xData = Format(pData, "dd/mm/yyyy")
        xDia = 32
        Mid(xData, 1, 2) = xDia
        Do Until IsDate(xData)
            xDia = xDia - 1
            Mid(xData, 1, 2) = xDia
        Loop
        fUltimoDiaMes = xDia
    End Function
Sub ValidaInteiro(ByRef tecla As Integer)
    Dim char As String
    
    char = Chr$(tecla)
    If char < "0" Or char > "9" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Sub ValidaLetra(ByRef tecla As Integer)
    Dim char As String
    
    char = Chr$(tecla)
    If char < "A" Or char > "Z" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Sub ValidaInteiroQtd(ByRef tecla As Integer)
    Dim char As String
    
    char = Chr$(tecla)
    If (char < "0" Or char > "9") And char <> "*" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Sub ValidaValor(ByRef pTecla As Integer)
    '3=^C, 22=^V, 24=^X
    If pTecla = 3 Or pTecla = 22 Or pTecla = 24 Then
        Exit Sub
    End If
    If Chr(pTecla) = "." Or Chr(pTecla) = "," Then
        Exit Sub
    End If
    If (Chr(pTecla) < "0" Or Chr(pTecla) > "9") Then
        If pTecla = 8 Then
            Exit Sub
        End If
        pTecla = 0
    End If
End Sub
Sub ValidaValorSinal(ByRef pTecla As Integer)
    '3=^C, 22=^V, 24=^X
    If pTecla = 3 Or pTecla = 22 Or pTecla = 24 Then
        Exit Sub
    End If
    If Chr(pTecla) = "." Or Chr(pTecla) = "," Or Chr(pTecla) = "-" Then
        Exit Sub
    End If
    If (Chr(pTecla) < "0" Or Chr(pTecla) > "9") Then
        If pTecla = 8 Then
            Exit Sub
        End If
        pTecla = 0
    End If
End Sub
Public Function VerificaHD() As Boolean
     
     VerificaHD = False
    ' Máscara da licença é:
    ' 9999      -> Ano da 1a Licença
    ' 9999      -> Número sequencial de Licença
    ' 99        -> Mes da 1a Licença
    ' 99        -> Dia da 1a Licença
    ' 999999999 -> Número de Série do HD
    gNumeroHd = DriveSerial(Left("C:", 1))
    
    'MsgBox "->" & gNumeroHd & "<-"
    'MsgBox Len(gNumeroHd)
    
    'Cerrado Informatica MV Aleister          "928722960"
    'Cerrado Informatica Aleister-PC          "2050478445"
    'Cerrado Informatica Pc-Suporte-2 (Juca)  "2085741908"
    If gNumeroHd = "928722960" Or gNumeroHd = "2050478445" Or gNumeroHd = "2085741908" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Recepcao daiany      "874335376"
    'Cerrado Informatica Recepcao daiany m.v. "613485991"
    'Cerrado Informatica Xen Serv. VeV servidor-sgbd "-1124578330"
    If gNumeroHd = "874335376" Or gNumeroHd = "613485991" Or gNumeroHd = "-1124578330" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Tasso Win XP         "1011313379"
    'Cerrado Informatica PC-DARUMA            "273085721"
    'Cerrado Informatica Tasso-PC Win-8 XP    "711249877"
    If gNumeroHd = "1011313379" Or gNumeroHd = "273085721" Or gNumeroHd = "711249877" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica SPED-01 Matheus      "-832546760"
    'Cerrado Informatica SPED-02              "-1333951473"
    'Cerrado Informatica MV-BD2008 ALEISTER RESTAURAR BD 2008 "944466570"
    If gNumeroHd = "-832546760" Or gNumeroHd = "-1333951473" Or gNumeroHd = "944466570" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Francisco mv-recebe-dados-vev           "-1929867039"
    'Cerrado Informatica Francisco mv-testa-cartao               "752438561"
    'Cerrado Informatica Francisco/mateus SPED-PC                "-528031409"
    'Cerrado Informatica Juarez (Serv.Maker)                     "-1872364250"
    'Cerrado Informatica Francisco(Res)                          "-276172724"
    If gNumeroHd = "-1929867039" Or gNumeroHd = "752438561" Or gNumeroHd = "-528031409" Or gNumeroHd = "-1872364250" Or gNumeroHd = "-276172724" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Mateus/Sped          "1612327772"
    'Cerrado Informatica Yago/Sped            "-1767431678"
    If gNumeroHd = "1612327772" Or gNumeroHd = "-1767431678" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica NoteBook Luis MV     "1486954119"
    'Cerrado Informatica MV SRV-UNISIS        "-265820933"
    'Cerrado Informatica PC-ISAC              "1779830539"
    If gNumeroHd = "1486954119" Or gNumeroHd = "-265820933" Or gNumeroHd = "1779830539" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Tasso Win XP MV      "-259173964"
    If gNumeroHd = "-259173964" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica Notebook ACER         "474179554"
    'Cerrado Informatica Servidor Tef Cerrado  "1150081308"
    'Cerrado Informatica PC DENICE             "-1735892654"
    If gNumeroHd = "474179554" Or gNumeroHd = "1150081308" Or gNumeroHd = "-1735892654" Then
        VerificaHD = True
    End If
    
    'Cerrado Informatica MV-jaime "1275590636"
    'Cerrado Informatica Dev04VM-Luiz "-1673518219"
    If gNumeroHd = "1275590636" Or gNumeroHd = "-1673518219" Then
        VerificaHD = True
    End If
    
     'Cerrado Informatica Not-Oliveira (empresa test) "1487533641"
    If gNumeroHd = "1487533641" Then
        VerificaHD = True
    End If
    
       
    'Posto Costa Escritorio Pista Serv BD  "782271162"
    'Posto Costa Notebook NFe Weberson     "1350114636"
    'Posto Costa Servidor PayGo            "174538327"
    If gNumeroHd = "782271162" Or gNumeroHd = "1350114636" Or gNumeroHd = "174538327" Then
        VerificaHD = True
    End If
    
    'Posto Costa PC-Cristiane            "-1065880334"
    'Posto Costa Werberson pc-camera     "681993716"
    'Posto Costa Notebook Ivander (novo) "-299440927"
    'Posto Costa Netbook (reserva)       "418540460"
    If gNumeroHd = "-1065880334" Or gNumeroHd = "681993716" Or gNumeroHd = "-299440927" Or gNumeroHd = "418540460" Then
        VerificaHD = True
    End If
    
    'CONV Costa CAIXA            "-1296795267"
    If gNumeroHd = "-1296795267" Then
        VerificaHD = True
    End If
    
         'J. A OLIVEIRA
    'Posto Costa Igor-PC         "-332401574"
    If gNumeroHd = "-332401574" Then
        VerificaHD = True
    End If
    
    
    'BOSQUE Pista TEF           "1422395015"
    'BOSQUE Servidor BD         "1818306558"
    'If gNumeroHd = "1422395015" Or gNumeroHd = "1818306558" Then
    '    VerificaHD = True
    'End If
    
    
    'BOSQUE Esc (Luzia)         "146436050"
    'BOSQUE Esc (Haroldo)       "1084333964"
    'BOSQUE Esc (Alessandra) notebook    "-464646810"
    'If gNumeroHd = "146436050" Or gNumeroHd = "1084333964" Or gNumeroHd = "-464646810" Then
    '    VerificaHD = True
    'End If
   
    'Posto Everest ECF         "-2007438058"
    If gNumeroHd = "-2007438058" Then
        VerificaHD = True
    End If
    
    
    'Posto Mutirao ECF               "480744591"
    'Posto Colorado ECF              "-1258427048"
    'Posto Solex ECF                 "686753962"
    'Posto Alto do Vale ECF          "274785215"
    'Posto Alto do Vale Adriano ESC  "-668327021"
    If gNumeroHd = "480744591" Or gNumeroHd = "-1258427048" Or gNumeroHd = "686753962" Or gNumeroHd = "274785215" Or gNumeroHd = "-668327021" Then
        VerificaHD = True
    End If
    
    'Central Postos  (Servidor)     "-2070071156"
    'Central Postos  (Denise)       "-597316237"
    'Central Postos  (Gilberto )    "576261776"
    If gNumeroHd = "-2070071156" Or gNumeroHd = "-597316237" Or gNumeroHd = "576261776" Then
        VerificaHD = True
    End If
    
    'Central Postos Wester          "2021790118"
    'Central Postos pc-recebe-dados "272608878"
    If gNumeroHd = "2021790118" Or gNumeroHd = "272608878" Then
        VerificaHD = True
    End If
    
    'Posto Belem ECF           "1153331749"
    'Posto Belem ESC           "541322182"
    If gNumeroHd = "1153331749" Or gNumeroHd = "541322182" Then
        VerificaHD = True
    End If
    
    'Posto Podium Ap.Goiânia   (est.)  "224533471"
    'Posto Podium Ap.Goiânia   (serv.) "979179491"
    'If gNumeroHd = "224533471" Or gNumeroHd = "979179491" Then
    '    VerificaHD = True
    'End If
    
    'Posto Papillon            "742594018"
    'If gNumeroHd = "742594018" Then
    '    VerificaHD = True
    'End If
    
    'Posto Cruzeiro ECF        "413189084"
    'Posto Rio Dourados TEF    "-60253195"
    'Posto Cruzeiro Paulo      "-1532769842"
    If gNumeroHd = "413189084" Or gNumeroHd = "-60253195" Or gNumeroHd = "-1532769842" Then
        VerificaHD = True
    End If
    
    'Posto Rio Formoso TEF PISTA     "-1404108374"
    'Posto Rio Formoso CAIXA         "-1308051231"
    'Posto Cruzeiro Paulo            "-334720585"
    If gNumeroHd = "-1404108374" Or gNumeroHd = "-1308051231" Or gNumeroHd = "-334720585" Then
        VerificaHD = True
    End If
    
    'Posto Tiradentes ECF      "-1671731998"
    'Posto Cidade Livre ECF    "-993454395"
    If gNumeroHd = "-1671731998" Or gNumeroHd = "-993454395" Then
        VerificaHD = True
    End If
    
    'Posto Atlanta ECF Servidor        "1618583259"
    'Posto Atlanta ECF Pista 2         "-2071409776"
    'Posto Atlanta Conveniencia        "-2145775857"
    If gNumeroHd = "1618583259" Or gNumeroHd = "-2071409776" Or gNumeroHd = "-2145775857" Then
        VerificaHD = True
    End If
    
    'Posto Atlanta PISTA 01        "1618583259"
    If gNumeroHd = "1618583259" Then
        VerificaHD = True
    End If
    
    
    'Posto V & V ECF                                   "-1560424844" 'Formatado pela cerrado e sgp instalado em 20/09/2012
    'Posto Vera Cruz                                   "1557225298"
    'Posto Vila Souza (Rosa Batista e Batista)         "-127579335"
    If gNumeroHd = "-1560424844" Or gNumeroHd = "1557225298" Or gNumeroHd = "-127579335" Then
        VerificaHD = True
    End If
    
    'Posto B & S ECF           "215194701"
    If gNumeroHd = "215194701" Then
        VerificaHD = True
    End If
    
    'Posto Pedro Ludovico ECF                "405649785"
    'Posto Presidente ECF                    "-805114323"
    If gNumeroHd = "405649785" Or gNumeroHd = "-805114323" Then
        VerificaHD = True
    End If
    
    'Posto 3 Irmaos (TRINDADE) Pista-1 Servidor BD  "138881428"
    'Posto 3 Irmaos (TRINDADE) Pista-2 Pay-GO       "-728608838"
    If gNumeroHd = "138881428" Or gNumeroHd = "-728608838" Then
        VerificaHD = True
    End If
    
            'POSTO BATISTA E BATISTA (BARAO)
    'PISTA - SVR BD                    "1650385426"
    If gNumeroHd = "1650385426" Then
        VerificaHD = True
    End If
    
    '# INICIO - Cidade Empresarial (Novo Controle)
    'Fernanda                   "-25830781"
    'Alberto                    "1088416630"
    'Kamila-PC                  "-655083637"
    If gNumeroHd = "-25830781" Or gNumeroHd = "1088416630" Or gNumeroHd = "-655083637" Then
        VerificaHD = True
    End If
    '# FIM - Cidade Empresarial (Novo Controle)
    
    
    'Computadores do Escritório da Cidade Empresarial
    'Estacao1                   "1325756496" Primeiro ao entrar
    'ServidorVeV                "1082127466" Segundo  ao Entrar
    'Gustavo                    "1176690536"                             "-959344634"
    If gNumeroHd = "1325756496" Or gNumeroHd = "1082127466" Or gNumeroHd = "1176690536" Then
        VerificaHD = True
    End If
    
    'Computadores do Escritório da Cidade Empresarial
    'PC-FABIANY                 "1947661654" Fabiany
    If gNumeroHd = "1947661654" Then
        VerificaHD = True
    End If
    
    'Computadores do Escritório da Cidade Empresarial
    'Kleber                     "204935238" (Mesa Alexandre)
    'PCKLEBER (megaware)        "-1362290915" (Mesa Kleber)
    'Notebook do Klebar         "-1293837818" (Levou pra Residendia, nao sei se utiliza)
    If gNumeroHd = "204935238" Or gNumeroHd = "-1362290915" Or gNumeroHd = "-1293837818" Then
        VerificaHD = True
    End If
    
    'Computadores da Cidade Empresarial
    'Notebook Alexandre        "-268104214"
    'pc-Alexandre              "449210400"
    'PC-BRUNO                  "-922839813"
    If gNumeroHd = "-268104214" Or gNumeroHd = "449210400" Or gNumeroHd = "-922839813" Then
        VerificaHD = True
    End If
    
    'Vila Alzira TEF           "-1071665400"
    'Bairro Ilda ECF           "1725140257"
    'BS ECF                    "872545438"
    If gNumeroHd = "-1071665400" Or gNumeroHd = "1725140257" Or gNumeroHd = "872545438" Then
        VerificaHD = True
    End If
    
    'Gaviao ECF                "1086513127"
    If gNumeroHd = "1086513127" Then
        VerificaHD = True
    End If
    
    'PC-Reserva I Rede VeV     "-1743147156"
    'PC-Reserva II Rede VeV    "-991389283"
    'PC-Reserva III Rede VeV   "1747620118"
    If gNumeroHd = "-1743147156" Or gNumeroHd = "-991389283" Or gNumeroHd = "1747620118" Then
        VerificaHD = True
    End If
    
    'Posto União ESC           "407573996"
    'Posto União TEF           "693704412"
    'If gNumeroHd = "407573996" Or gNumeroHd = "693704412" Then
    '    VerificaHD = True
    'End If
    
    'Posto Igualdade Res/helmane      "1818224842"
    'Posto Igualdade Notebook Helmane "184393060"
    'Posto Igualdade 1 Ecf            "1320207168" 'Garavelo
    If gNumeroHd = "1818224842" Or gNumeroHd = "184393060" Or gNumeroHd = "1320207168" Then
        VerificaHD = True
    End If
    
    'Posto Igualdade 2 Ecf (servidor)     "-1700889873" 'Goiá
    'Posto Igualdade 2 Escritorio-camera  "1309912941" 'Goiá  'Formatado por terceiros e sgp instalado em 20/09/2012
    If gNumeroHd = "-1700889873" Or gNumeroHd = "1309912941" Then
        VerificaHD = True
    End If
    
    'Posto Igualdade 3 Ecf     "-2003707064"   'Anel Viário
    'Posto Igualdade 4 Ecf     "-1743597838"  'Independência
    If gNumeroHd = "-2003707064" Or gNumeroHd = "-1743597838" Then
        VerificaHD = True
    End If
            'POSTO BOM SUCESSO (OLOANA)
    'Posto Bom Sucesso RESID.(BD ANTIGO)     "1156763908"
    'Posto Bom Sucesso ECF                   "-1663847479"
    If gNumeroHd = "1156763908" Or gNumeroHd = "-1663847479" Then
        VerificaHD = True
    End If
    
    'Posto Rubi TEF            "-1292335901"
    'Posto Rubi PAY GO         "1023284830"
    'Posto Rubi ESC Juliano    "1381981526"
    If gNumeroHd = "-1292335901" Or gNumeroHd = "1023284830" Or gNumeroHd = "1381981526" Then
        VerificaHD = True
    End If
    
    'Posto Rubi Note-Juliano      "-1907670450"
    'Posto Rubi ESC MARCO AURELIO "-1804988580"
    'Posto Rubi PISTA             "-1292335901"
    If gNumeroHd = "-1907670450" Or gNumeroHd = "-1804988580" Or gNumeroHd = "-1292335901" Then
        VerificaHD = True
    End If
    
    'Posto Esmeralda Esc            "-1945379034"
    'Posto Esmeralda PISTA-1 TEF    "1245025101"
    'Posto Esmeralda NoteBook       "671480579"
    If gNumeroHd = "-1945379034" Or gNumeroHd = "1245025101" Or gNumeroHd = "671480579" Then
        VerificaHD = True
    End If
    
    'Posto Esmeralda TEF Serv pay-go     "-1668111990"
    'Posto Esmeralda PISTA-2 TEF         "748752674"
    'Posto Esmeralda Marco Aurelio       "-1473057162"
    If gNumeroHd = "-1668111990" Or gNumeroHd = "748752674" Or gNumeroHd = "-1473057162" Then
        VerificaHD = True
    End If
    
    'Conveniencia Esmeralda Caixa    "215614881"
    'Conveniencia Esmeralda Notebook "2028036911"
    'Conveniencia Esmeralda Netbook (LG) "1788251903"
    If gNumeroHd = "215614881" Or gNumeroHd = "2028036911" Or gNumeroHd = "1788251903" Then
        VerificaHD = True
    End If
    
    
    'Posto Sol Nascente ECF pista "976819207"
    'Posto Sol Nascente ECF LOJA  "695145216"
    'Posto Sol Nascente ESC       "-1536732893"
    'If gNumeroHd = "976819207" Or gNumeroHd = "695145216" Or gNumeroHd = "-1536732893" Then
    '    VerificaHD = True
    'End If
    
    'Escritorio de Contabilidade do Vandeir
    'Notebook HP 4GB 500GB Vandeir       "-1237000432"
    If gNumeroHd = "-1237000432" Then
        VerificaHD = True
    End If
    
    'PetroPuro Esc      "-1264906999"
    'PetroPuro ECF      "-1971808768"
    If gNumeroHd = "-1971808768" Or gNumeroHd = "-1264906999" Then
        VerificaHD = True
    End If
    
            'Serve Bem - ideal
    'Serve Bem ECF      "552734078"
    'Serve Bem Esc      "2060647775"
    If gNumeroHd = "552734078" Or gNumeroHd = "2060647775" Then
        VerificaHD = True
    End If
    
    '3T Escritorio      "-1662694784"
    '3T ECF             "351465373"
    '3T Res.Teuler      "282763366"
    If gNumeroHd = "-1662694784" Or gNumeroHd = "351465373" Or gNumeroHd = "282763366" Then
        VerificaHD = True
    End If
    
    'AROEIRA Esc        "-1742087126"
    'AROEIRA TEF        "881265288"
'    If gNumeroHd = "-1742087126" Or gNumeroHd = "881265288" Then
'        VerificaHD = True
'    End If
    
    'HS Informática     "476976599"
'    If gNumeroHd = "476976599" Then
'        VerificaHD = True
'    End If
    
    'Ventania Esc                "2120807908"
    'Ventania TEF Serv.BD/PISTA  "2090857296"
    'Ventania Escritorio 2       "1686731313"
    If gNumeroHd = "2120807908" Or gNumeroHd = "2090857296" Or gNumeroHd = "1686731313" Then
        VerificaHD = True
    End If
    
    'Ventania Res Andrea       "1487086925"
    'Ventania Notebook Andrea  "1719721608"
    'Ventania Caixa NFE        "-1943120484"
    If gNumeroHd = "1487086925" Or gNumeroHd = "1719721608" Or gNumeroHd = "-1943120484" Then
        VerificaHD = True
    End If
    
    'Churrascaria Monjolo      "-1498394830"
    If gNumeroHd = "-1498394830" Then
        VerificaHD = True
    End If
    
    'Parque(bsb) ECF Serv. Pay-GO  "2029601401"
    'Parque Esc Serv.BD            "1623891176"
    If gNumeroHd = "2029601401" Or gNumeroHd = "1623891176" Then
        VerificaHD = True
    End If
    
    'Recreio Esc Serv.BD           "-2144430562" Serv.BD
    'Recreio Pista ECF TEF         "877565761"
    If gNumeroHd = "-2144430562" Or gNumeroHd = "877565761" Then
        VerificaHD = True
    End If
    
    'Supermercado Campos       "-2134537283"
'    If gNumeroHd = "-2134537283" Then
'        VerificaHD = True
'    End If
    
    'Conv.Free Store - Loja    "725226721"
    'Conv.Free Store - pc-camera "807061341" 14/08/2009
    'Conv.Free Store - Escritorio2 "683500759" 25/09/2009
    'If gNumeroHd = "725226721" Or gNumeroHd = "807061341" Or gNumeroHd = "683500759" Then
    '    VerificaHD = True
    'End If
    
    'Conv."AV. MUTIRAO IPIRANGA"  - Loja    "1748433760" 10/09/2009
    'If gNumeroHd = "1748433760" Then
    '    VerificaHD = True
    'End If
    
    'Conv.TJ     - Loja    "2091198530"
    If gNumeroHd = "2091198530" Then
        VerificaHD = True
    End If
    
    'Dist. Pimenta      "-385996109"
'    If gNumeroHd = "-385996109" Then
'        VerificaHD = True
'    End If



        'POSTO SOL-MJ
    'Posto Sol - Automacao EzTech/ECF  "1894374233"
    'Posto Sol - PC Pista              "-593971708"
    If gNumeroHd = "1894374233" Or gNumeroHd = "-593971708" Then
        VerificaHD = True
    End If
    
        'AUTO POSTO SARA
    'Posto Sara computador:posto    "1959080686"
    If gNumeroHd = "1959080686" Then
        VerificaHD = True
    End If
    
    'Escritorio de Contabilidade do Posto Sara
    'MarkPerson                   "-261723008"
    'Wesley                       "-867206099"
    'If gNumeroHd = "-261723008" Or gNumeroHd = "-867206099" Then
    '    VerificaHD = True
    'End If

    'Escritorio de Contabilidade do Emivaldo (pedro ludovico)
    'Emivaldo                     "794498309"
    'If gNumeroHd = "794498309" Then
    '    VerificaHD = True
    'End If

    'Posto Ratinho TEF Pista-01            "1117012064"
    'Posto Ratinho TEF Pista-02            "71471171"
    'Posto Ratinho Servidor-BD             "-857498964" Maquina Virtual
    If gNumeroHd = "1117012064" Or gNumeroHd = "71471171" Or gNumeroHd = "-857498964" Then
        VerificaHD = True
    End If

    'Posto Ratinho Escritorio 1 "1492233534" Camila
    'Posto Ratinho Escritorio 2 "939963216" Cameras (formatado 09/02/2010)
    'Posto Ratinho NetBookJoseEduardo "-21293415" NetBook HP (novo em 16/12/2009)
    'Posto Ratinho PC Taina     "1258163121"
    If gNumeroHd = "1492233534" Or gNumeroHd = "939963216" Or gNumeroHd = "-21293415" Or gNumeroHd = "1258163121" Then
        VerificaHD = True
    End If

    'Posto Ratinho Serv-Automacao           "-1933426647"  'Serv.Automacao'
    'Posto Ratinho/Esmeralda/Rubi Reserva   "819094858"
    'Posto Ratinho/Esmeralda/Rubi NOTEBOOK JOSE EDUARDO SONY VAIO "1682760774"
    'Posto Ratinho/Esmeralda/Rubi Reserva 2 "1619477437"
    If gNumeroHd = "-1933426647" Or gNumeroHd = "819094858" Or gNumeroHd = "1682760774" Or gNumeroHd = "1619477437" Then
        VerificaHD = True
    End If

    ' ** EX-CLIENTE ** Posto Serra Dourada TEF   "878857714"
'    If gNumeroHd = "878857714" Then
'        VerificaHD = True
'    End If

    'P. Rios Mendes - Flavia Cont.            "886068550"
    'P. Rios Mendes - Flavia Cont. Notebook   "1022171256"
    'P. Rios Mendes - Janaina Notebook        "-91481095"
    If gNumeroHd = "886068550" Or gNumeroHd = "1022171256" Or gNumeroHd = "-91481095" Then
        VerificaHD = True
    End If

    'P. Mozarlandia Miron II  all in one acer  SERV. BD-CUPOM "-55460152"
    'P. Mozarlandia - Miron NA LOJA DE VEICULOS               "-2011493561"
    If gNumeroHd = "-55460152" Or gNumeroHd = "-2011493561" Then
        VerificaHD = True
    End If

    'Posto Rios Mendes Ltda    "1756751910"
    'Posto Rios Mendes Ltda (Miron - Garagem)   "875184173"
    If gNumeroHd = "1756751910" Or gNumeroHd = "875184173" Then
        VerificaHD = True
    End If

    'Sao Mateus TEF            "1146623901"
    'Sao Mateus ESC            "-1937398005"
    'Sao Mateus Conv.          "1017438851"
'    If gNumeroHd = "1146623901" Or gNumeroHd = "-1937398005" Or gNumeroHd = "1017438851" Then
'        VerificaHD = True
'    End If


    'P AM Esc Fernanda         "-61267268"
    'P AM TEF PISTA            "1759942121"
'    If gNumeroHd = "-61267268" Or gNumeroHd = "1759942121" Then
'        VerificaHD = True
'    End If

    'Zenit - ECF    "-1669898741"
    'If gNumeroHd = "-1669898741" Then
    '    VerificaHD = True
    'End If

    'Divertel Pista TEF             "-200105308"
    'Divertel Escritorio (Ludimila) "75102294"
    'Divertel Escritorio (1 na Entrada) "477319347"
    'If gNumeroHd = "-200105308" Or gNumeroHd = "75102294" Or gNumeroHd = "477319347" Then
    '    VerificaHD = True
    'End If

    'Divertel Servidor BD           "-1199596521"
    'If gNumeroHd = "-1199596521" Then
    '    VerificaHD = True
    'End If

    'Brasilian Conveniencia       "2026336322"
    If gNumeroHd = "2026336322" Then
        VerificaHD = True
    End If
        
        'POSTO CENTRAL I
    'PA-Central I ECF Automacao           "1491985196"
    'PA-Central I Escritorio Maria Jose   "1051646906"
    'PA-Central I Escritorio Gilmark      "-561549260"
    If gNumeroHd = "1491985196" Or gNumeroHd = "1051646906" Or gNumeroHd = "-561549260" Then
        VerificaHD = True
    End If

    'PA-Central I Escritorio Walgney        "-396844061"
    'PA-Central I Escritorio Weber          "-1672457803"
    'PA-Central I Escritorio Manuel-PC      "1752860510"
    If gNumeroHd = "-396844061" Or gNumeroHd = "-1672457803" Or gNumeroHd = "1752860510" Then
        VerificaHD = True
    End If

    'PA-Central I Notebook Negão          "548567400"
    'PA-Central I Recepcao Regina         "-1737583111"
    'PA-Central I mv-servidor  banco      "404049645"
    If gNumeroHd = "548567400" Or gNumeroHd = "-1737583111" Or gNumeroHd = "404049645" Then
        VerificaHD = True
    End If

    'PA-Central I Conveniencia            "-631092339"
    'PA-Central I Samyla                  "2021881120"
    If gNumeroHd = "-631092339" Or gNumeroHd = "2021881120" Then
        VerificaHD = True
    End If
    
        'POSTO GRANADA II
    'PA-Granada II Pista-ECF              "348424244"
    'PA-Granada II Escritorio             "-2077831345"
    'PA-Granada II Cleuto Filho Note      "1578153943"
    If gNumeroHd = "348424244" Or gNumeroHd = "-2077831345" Or gNumeroHd = "1578153943" Then
        VerificaHD = True
    End If
    
    'PA-Granada II PC DO CAIXA            "1222217159"
    If gNumeroHd = "1222217159" Then
        VerificaHD = True
    End If
    
        'POSTO GRANADA I
    'PA-Granada I Escritorio              "14721637"
    'PA-Granada I Pista-ECF               "2116169992"
    'PA-Granada I Servidor 2              "-693313759"
    If gNumeroHd = "14721637" Or gNumeroHd = "2116169992" Or gNumeroHd = "-693313759" Then
        VerificaHD = True
    End If
    
    'PA-Granada I Marilene                "561249260"
    'PA-Granada I Servidor                "-693313759"
    If gNumeroHd = "561249260" Or gNumeroHd = "-693313759" Then
        VerificaHD = True
    End If
    
        'POSTO REDENÇÃO
    'PA-Redencao Pista-ECF                "-800489349"
    'PA-Redencao Escritorio               "-1060004516"
    If gNumeroHd = "-800489349" Or gNumeroHd = "-1060004516" Then
        VerificaHD = True
    End If
            
        'POSTO CONCEIÇÃO
    'PA-Conceição Pista-ECF                "1053564204"
    'PA-Conceição Zé Mineiro               "1053564204"
    If gNumeroHd = "1053564204" Or gNumeroHd = "1053564204" Then
        VerificaHD = True
    End If

    '- Posto Ricardo Paranhos (Valposto)
    'P Valposto Pista1                             "1011751311"
    'P Valposto Pista-2                            "1280488698"
    'P Valposto PC-RESERVA  (Escritorio-Givanildo) "1615477594"
    'P Valposto Serv.Xen    (Servidor BD/Tef)      "1421752457"
    If gNumeroHd = "1011751311" Or gNumeroHd = "1280488698" Or gNumeroHd = "1615477594" Or gNumeroHd = "1421752457" Then
        VerificaHD = True
    End If
    
    '- Posto Ricardo Paranhos (Valposto)
    'P Valposto Notebook HP (Fabiano)        "651050852"
    'P Valposto Posto (Escritorio 2o Andar)  "-1540256466"
    'P Valposto Notebook (Fernando)          "1483351151"
    If gNumeroHd = "651050852" Or gNumeroHd = "-1540256466" Or gNumeroHd = "1483351151" Then
        VerificaHD = True
    End If

    '- Posto Ricardo Paranhos (Valposto)
    'P Valposto PC-GIVANILDO (Residencia Givanildo) "-1676583747"
    'P Valposto Escritorio                          "82459011"
    If gNumeroHd = "-1676583747" Or gNumeroHd = "82459011" Then
        VerificaHD = True
    End If
    
            'Posto Ricardo Paranhos (Valposto) "Conv. Emporio Marista"
    'Emporio Marista caixa                      "-1428807243"
    If gNumeroHd = "-1428807243" Then
        VerificaHD = True
    End If

    'P Rio Verde Pista TEF              "553515613"
    'P Rio Verde Escritorio (Servidor)  "-594330800"
    'P Rio Verde (NoteBook Leonardo)    "487526869"
    If gNumeroHd = "553515613" Or gNumeroHd = "-594330800" Or gNumeroHd = "487526869" Then
        VerificaHD = True
    End If

    'Contabilidade rede VW - Fabia   "277168767"
    If gNumeroHd = "277168767" Then
        VerificaHD = True
    End If

    'Conv. JC Circular                "1485141650"
    If gNumeroHd = "1485141650" Then
        VerificaHD = True
    End If

    'Rest. Santos e Murate Caixa      "-663680762"
    If gNumeroHd = "-663680762" Then
        VerificaHD = True
    End If

    'P. Talismã Escritorio Cameras     "1"
    'P. Talismã Escritorio (Jaqueline) "1827317347"
    'P. Talismã Servidor Pista         "616854414"
    If gNumeroHd = "1" Or gNumeroHd = "1827317347" Or gNumeroHd = "616854414" Then
        VerificaHD = True
    End If
    
    'Rede AM
    'Posto JH (Jardim Helvécia) Saída para Nerópolis - São Judas.
    'P. AM (Alex) Pista ECF           "-524233484"
    'P. AM (Alex) Escritório          "-892332545"
    'Rede AM - Contador (Fabrício)    "-151297353"
    'Rede AM - ALEX NOTEBOOK PESSOAL  "-1166822417"
    If gNumeroHd = "-524233484" Or gNumeroHd = "-892332545" Or gNumeroHd = "-151297353" Or gNumeroHd = "-1166822417" Then
        VerificaHD = True
    End If

    'Rede AM
    'Posto JH da Rio Verde do Alex
    'no jardim Helvécio
    'P. JH (Alex) Pista ECF           "-2000863285"
    'P. JH (Alex) Escritorio          "67699329"
    If gNumeroHd = "-2000863285" Or gNumeroHd = "67699329" Then
        VerificaHD = True
    End If


    'Rede AM - Posto SERRA DOURADA do Alex
    'Posto SERRA DOURADA Pista ECF           "276128426"
    'Posto SERRA DOURADA Escritorio          "-689957191"
    If gNumeroHd = "276128426" Or gNumeroHd = "-689957191" Then
        VerificaHD = True
    End If

    
    'Tigrão Escritorio/Camera  "-1126116035"
    'Tigrão Pista EFC          "707189010"
    If gNumeroHd = "-1126116035" Or gNumeroHd = "707189010" Then
        VerificaHD = True
    End If
    
    'Tigrão SOMA Contabilidade (Gardenia)      "1490317609"
    'Tigrão SOMA Contabilidade (Luiz Henrique) "-1595992920"
    'Tigrão SOMA Contabilidade (Idelma)        "815383212"
    If gNumeroHd = "1490317609" Or gNumeroHd = "-1595992920" Or gNumeroHd = "815383212" Then
        VerificaHD = True
    End If
    
                   'POSTO 55 FLF
    'Posto 55 Escritorio         "-194983863"
    'Posto 55 Pista EFC          "547754071"
    'Posto 55 (Conveniencia FAMILY) Caixa Serv BD  "1424503042"
    If gNumeroHd = "-194983863" Or gNumeroHd = "547754071" Or gNumeroHd = "1424503042" Then
        VerificaHD = True
    End If

    'Posto 55 Esc Brasilia       "1546734580"
    'Posto 55 Esc PISTA 2       "-931907446"
    If gNumeroHd = "1546734580" Or gNumeroHd = "-931907446" Then
        VerificaHD = True
    End If
    
    'P. AM Bela Vista - Pista ECF           "174007410"
    If gNumeroHd = "174007410" Then
        VerificaHD = True
    End If

    'Posto Duarte - Pista ECF           "-1262061109"
    'Posto Duarte - Servidor BD         "519586913"
    If gNumeroHd = "-1262061109" Or gNumeroHd = "519586913" Then
        VerificaHD = True
    End If

    'P. Eldorado - Pista ECF            "744626998"
    'P. Eldorado - NFe                  "-1735530550"
    If gNumeroHd = "744626998" Or gNumeroHd = "-1735530550" Then
        VerificaHD = True
    End If

    'Posto Portal da Chapada - Pista ECF           "-1806804679"
    'Posto Portal da Chapada - Escritorio          "140111827"
    If gNumeroHd = "-1806804679" Or gNumeroHd = "140111827" Then
        VerificaHD = True
    End If

    'Posto Aragoiania - ECF                      "-860509781"
    If gNumeroHd = "-860509781" Then
        VerificaHD = True
    End If
    
    'Posto Vera Cruz (Rodovia)  PC-PISTA   "-2112673729"
    'Posto Vera Cruz SVR-BANCODADOS        "1691921771"
    'Posto Vera Cruz PC-ALEXANDRA          "-2070358589"
    If gNumeroHd = "-2112673729" Or gNumeroHd = "1691921771" Or gNumeroHd = "-2070358589" Then
        VerificaHD = True
    End If
    
    'Posto Vera Cruz Airton                "744885599"
    'Posto Vera Cruz Maria Alice           "1116663977"
    'Posto Vera Cruz CAIO                  "-222836388"
    
    If gNumeroHd = "744885599" Or gNumeroHd = "1116663977" Or gNumeroHd = "-222836388" Then
        VerificaHD = True
    End If
    
    'Posto Paineiras ECF                   "616941636"
    'Posto Paineiras Notebook-Alisson      "1554386276"
    'Posto Paineiras Conveniencia          "1879941863"
    'Posto Paineiras Escritorio            "1980263712"
    If gNumeroHd = "616941636" Or gNumeroHd = "1554386276" Or gNumeroHd = "1879941863" Or gNumeroHd = "1980263712" Then
        VerificaHD = True
    End If
    
    'Posto Guaira ECF PISTA SERV "-1006542841"
    'Posto Guaira Escritório     "742239039"
    If gNumeroHd = "-1006542841" Or gNumeroHd = "742239039" Then
        VerificaHD = True
    End If
    
    'Posto França (Caixa) ECF/TEF           "-453537385"
    'Posto França Escritorio Servidor BD    "615444953"
    If gNumeroHd = "-453537385" Or gNumeroHd = "615444953" Then
        VerificaHD = True
    End If

    'Restaurante Cabeca de Mel              "-1877484435"
    If gNumeroHd = "-1877484435" Then
        VerificaHD = True
    End If

    'Posto AM ANEL VIARIO - ECF                     "-1398333983"
    'Posto AM ANEL VIARIO - ALL IN 1 (ACER) ALEX    "-1535010949"
    'Posto AM ANEL VIARIO - SERVIDOR BD             "1488136080"
    'Posto AM ANEL VIARIO - SECRETARIA              "-39484525"
    If gNumeroHd = "-1398333983" Or gNumeroHd = "-1535010949" Or gNumeroHd = "1488136080" Or gNumeroHd = "-39484525" Then
        VerificaHD = True
    End If
    
    'Posto AM ANEL VIARIO - PC NABIL             "-1795361983"
    If gNumeroHd = "-1795361983" Then
        VerificaHD = True
    End If
    
    'Posto Pendao - Conveniencia              "-121947569"
    'Posto Pendao - Pista ecf servidor bd     "13194430"
    'Posto Pendao - Escritorio de outro posto "548770641"
    If gNumeroHd = "-121947569" Or gNumeroHd = "13194430" Or gNumeroHd = "548770641" Then
        VerificaHD = True
    End If
    
    'Posto Santa Helena Pista Serv BD   "988119492"
    'Posto Santa Helena Escritorio      "76988877"
    If gNumeroHd = "988119492" Or gNumeroHd = "76988877" Then
        VerificaHD = True
    End If
    
    'Conveniencia Galinha Serv BD       "-16907318"
    'If gNumeroHd = "-16907318" Then
    '    VerificaHDb = True
    'End If
    
    'Monaco Pista Serv BD               "-1096297208"
    'Monaco Escritorio                  "-996510105"
    If gNumeroHd = "-1096297208" Or gNumeroHd = "-996510105" Then
        VerificaHD = True
    End If
    
    'Brisa - Pista Serv BD              "-1865106262"
    'Brisa - Escritório                 "141403610"
    If gNumeroHd = "-1865106262" Or gNumeroHd = "141403610" Then
        VerificaHD = True
    End If
    
    'VeV Lubrificantes - Servidor       "-2139736781"
    'VeV Lubrificantes - Paulo-PC       "1658863407"
    If gNumeroHd = "-2139736781" Or gNumeroHd = "1658863407" Then
        VerificaHD = True
    End If

    'Posto Novo Horizonte Pista BD   "751875553"
    'Posto Novo Horizonte Escritorio "1017231235"
    If gNumeroHd = "751875553" Or gNumeroHd = "1017231235" Then
        VerificaHD = True
    End If

    'Caiapo Pio XII - Pista ECF      "1580825054"
    'Felicidade     - Pista ECF      "-1642450018"
    If gNumeroHd = "1580825054" Or gNumeroHd = "-1642450018" Then
        VerificaHD = True
    End If
    
    'Caiapo / Felicidade - Contabilidade: Cida      "1015354162"
    'Caiapo / Felicidade - Contabilidade: JUNIOR    "546154551"
    'Caiapo / Felicidade - Contabilidade: recepçao  "-932127927"
    If gNumeroHd = "1015354162" Or gNumeroHd = "546154551" Or gNumeroHd = "-932127927" Then
        VerificaHD = True
    End If
    
    'Caiapo / SVRREDECAIAPO - Contabilidade         "217039471"
    If gNumeroHd = "217039471" Then
        VerificaHD = True
    End If
    
    'T-13 PISTA-01 - calibrador                     "-1709838527"
    'T-13 SVR-BD PISTA-02                           "942040818"
    'T-13 PISTA-03                                  "-1002410244"
    If gNumeroHd = "-1709838527" Or gNumeroHd = "942040818" Or gNumeroHd = "-1002410244" Then
        VerificaHD = True
    End If

    'T-13 ** COMPUTADORES DO ESCRITORIO
    'T-13 PC-NFE/JORDANA                            "-2035996501"
    'T-13 ESCRITORIO (ATUAL SERVIDOR REZENDE)       "1822925500"
    'T-13 CONTABIL (TEORICAMENTE NAO USA)           "-1670033919"
    If gNumeroHd = "-2035996501" Or gNumeroHd = "1822925500" Or gNumeroHd = "-1670033919" Then
        VerificaHD = True
    End If

    
    'Conveniencia T-13 PC-LAURO                     "1244948758"
    'Conveniencia T-13 SVR-BD                       "-1340265950"
    'Conveniencia T-13 LOJA2                        "-2144516003"
    'Conveniencia T-13 LOJA1                        "-265958053"
    If gNumeroHd = "1244948758" Or gNumeroHd = "-1340265950" Or gNumeroHd = "-2144516003" Or gNumeroHd = "-265958053" Then
        VerificaHD = True
    End If

    'Conveniencia T-13 CONV4/ESCRITORIO             "686083416"
    If gNumeroHd = "686083416" Then
        VerificaHD = True
    End If
    
        'AUTO POSTO INDUSTRIAL
    'Auto Posto Industrial Not-Oliveira    "1487533641"
    'Auto Posto Industrial Pista-Pc        "2103352553"
    'Auto Posto Industrial Servidor-BD     "-2103352553"
    If gNumeroHd = "1487533641" Or gNumeroHd = "2103352553" Or gNumeroHd = "-2103352553" Then
        VerificaHD = True
    End If
    
         'SUPERMECADO SUDOESTE (E A MARTINS ARM)
    'Supermecado Sudoeste Caixa1           "1421175095"
    'Supermecado Sudoeste Caixa2           "1742550999"
    'Supermecado Sudoeste Escritorio       "-1742550999"
    If gNumeroHd = "1421175095" Or gNumeroHd = "1742550999" Or gNumeroHd = "-1742550999" Then
        VerificaHD = True
    End If
    
    'Supermecado Sudoeste Givanildo-PC     "1891254976"
    If gNumeroHd = "1891254976" Then
        VerificaHD = True
    End If
    
          'Rede ADM COMBUSTIVEL
    'ANTONIO MELO-PC                      "373374433"
    'FREDERICO-STI                        "-1804406464"
    'ANDREA-PC                            "978693774"
    If gNumeroHd = "373374433" Or gNumeroHd = "-1804406464" Or gNumeroHd = "978693774" Then
        VerificaHD = True
    End If
    
    'PC-SERVIDOR                       "1246987327"
    If gNumeroHd = "1246987327" Then
        VerificaHD = True
    End If
    
            'MOREIRA E TOSTA (TIGRAO)
    'PC-ESCRITORIO                     "1490739413"
    'PC-TIGRAOPT                       "-2137045885"
    If gNumeroHd = "1490739413" Or gNumeroHd = "-2137045885" Then
        VerificaHD = True
    End If
        
            'POSTO CORDEIRO (GO 118)
    'PC-PISTA                   "136495726"
    'PC-ESCRITORIO              "-332505978"
    If gNumeroHd = "136495726" Or gNumeroHd = "-332505978" Then
        VerificaHD = True
    End If
    
            'POSTO BRITO BARROS
    'PC-PISTA                   "-1635412902"
    'PC-ESCRITORIO              "2026120169"
    If gNumeroHd = "-1635412902" Or gNumeroHd = "2026120169" Then
        VerificaHD = True
    End If
    
            'POSTO V8 (RPI)
    'PC-ESCRITORIO                  "-636794512"
    If gNumeroHd = "-636794512" Then
        VerificaHD = True
    End If
    
    
End Function


Public Function VerificaPlacaRede() As Boolean
    Dim xNumeroLAN As String
    VerificaPlacaRede = False
    xNumeroLAN = GetMACAddress
    
    'MsgBox "->" & xNumeroLAN & "<-"
    
    'Cerrado Informatica NoteBook             "679837946"
    If xNumeroLAN = "00 C0 9F 57 D7 BA" Then
        VerificaPlacaRede = True
    End If
    
    'HS Informatica NoteBook             "679837946"
    If xNumeroLAN = "00 D0 59 93 1A 4C" Then
        VerificaPlacaRede = True
    End If
    

End Function


Sub sqlData(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(35) & Format(xString1, "mm/dd/yyyy") & Chr(35) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(35) & Format(xString2, "mm/dd/yyyy") & Chr(35)
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            If xString1 = "00:00:00" Then
                gSQL = gSQL & "Null" & xString2
            Else
                gSQL = gSQL & Chr(39) & Format(xString1, "dd/mm/yyyy") & Chr(39) & xString2
            End If
        ElseIf xDelimitador = 2 Then
            If xString2 = "00:00:00" Then
                gSQL = gSQL & xString1 & "Null"
            Else
                gSQL = gSQL & xString1 & Chr(39) & Format(xString2, "dd/mm/yyyy") & Chr(39)
            End If
        End If
    End If
End Sub
Sub sqlHora(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(35) & Format(xString1, "hh:mm:ss") & Chr(35) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(35) & Format(xString2, "hh:mm:ss") & Chr(35)
        End If
    End If
End Sub
Sub sqlNumero(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & CLng(xString1) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & CLng(xString2)
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            gSQL = gSQL & CLng(xString1) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & CLng(xString2)
        End If
    End If
End Sub
Sub sqlTexto(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    If bdAccess Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(39) & verificaCaracterEspecial(xString1) & Chr(39) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(39) & verificaCaracterEspecial(xString2) & Chr(39)
        End If
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            gSQL = gSQL & Chr(39) & verificaCaracterEspecial(xString1) & Chr(39) & xString2
        ElseIf xDelimitador = 2 Then
            gSQL = gSQL & xString1 & Chr(39) & verificaCaracterEspecial(xString2) & Chr(39)
        End If
    End If
End Sub
Sub sqlValor(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    Dim x_valor As String
    If bdAccess Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "0000000000.00")
            Mid(xString1, 11, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "0000000000.00")
            Mid(xString2, 11, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    ElseIf bdSqlServer Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "0000000000.0000;-000000000.0000")
            Mid(xString1, 11, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "0000000000.0000;-000000000.0000")
            Mid(xString2, 11, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    End If
End Sub
Sub sqlValor4(ByVal xDelimitador As Integer, ByVal xString1 As String, ByVal xString2 As String)
    Dim x_valor As String
    If bdAccess Then
        If xDelimitador = 1 Then
            xString1 = Format(CCur(xString1), "00000000.0000")
            Mid(xString1, 9, 1) = "."
        ElseIf xDelimitador = 2 Then
            xString2 = Format(CCur(xString2), "00000000.0000")
            Mid(xString2, 9, 1) = "."
        End If
        gSQL = gSQL & xString1 & xString2
    End If
End Sub
Function BioCriaImprime() As String
    Dim xNomeArquivo As String
    xNomeArquivo = Format(Day(Date), "00") & Format(Time, "hhmmss") & ".TMV"
    Set gArquivoTMV = gArqTxt.CreateTextFile(xNomeArquivo, True)
    Mid(xNomeArquivo, 10, 3) = "TMP"
    Set gArquivoTMP = gArqTxt.CreateTextFile(xNomeArquivo, True)
    BioCriaImprime = xNomeArquivo
End Function
Sub BioImprime(ByVal pString As String)
    gArquivoTMP.WriteLine (pString)
    If pString = "@@Printer.NewPage" Or pString = "@@Printer.CurrentY = 0" Then
        gArquivoTMV.WriteLine (" ")
    ElseIf Mid(pString, 1, 2) <> "@@" Then
        gArquivoTMV.WriteLine (Mid(pString, 16, Len(pString) - 15))
    End If
End Sub
Sub BioFechaImprime()
    gArquivoTMV.Close
    gArquivoTMP.Close
    Set gArquivoTMV = Nothing
    Set gArquivoTMP = Nothing
End Sub
Function BuscaConfiguracaoConsulta(ByVal pPrograma As String) As Boolean
    Dim lSQL As String
    Dim rst As New ADODB.Recordset
    Dim xPrograma As String
    
    On Error GoTo FileError
    
    BuscaConfiguracaoConsulta = False
    
    xPrograma = Mid(pPrograma, 1, 40)
    lSQL = "SELECT Campo, Operador, Condicao"
    lSQL = lSQL & " FROM Configuracao_Pesquisa"
    lSQL = lSQL & " WHERE Programa = " & preparaTexto(xPrograma)
    'rst.CursorLocation = adUseClient
    'rst.MaxRecords = 1
    'rst.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Set rst = Conectar.RsConexao(lSQL)
    If Not rst.EOF Then
        BuscaConfiguracaoConsulta = True
        g_string = rst!campo & "|@|"
        g_string = g_string & rst!operador & "|@|"
        g_string = g_string & rst!Condicao & "|@|"
    End If
    rst.Close
    Set rst = Nothing
    Exit Function
FileError:
    rst.Close
End Function
Public Function CalculaDataAbastecimentoVencimento(ByVal mCodigoVencimento As Integer, ByVal mDataHoje As Date) As String
    Dim Vencimento As New cVencimento
    Dim xDataI As String
    Dim xDataF As String
    Dim xDataV As String
    Dim xDiaI As Integer
    Dim xDiaF As Integer
    Dim xQtdDias As Integer
    Dim i As Integer
    On Error GoTo FileError
    
    CalculaDataAbastecimentoVencimento = ""
    If Vencimento.LocalizarCodigo(mCodigoVencimento) Then
        If Day(mDataHoje) > Vencimento.Dia1Final Then
            xDiaI = Vencimento.Dia1Inicial
            xDiaF = Vencimento.Dia1Final
            xQtdDias = Vencimento.Dia1Vencimento
        ElseIf Day(mDataHoje) > Vencimento.Dia2Final Then
            xDiaI = Vencimento.Dia2Inicial
            xDiaF = Vencimento.Dia2Final
            xQtdDias = Vencimento.Dia2Vencimento
        Else
            xDiaI = Vencimento.Dia3Inicial
            xDiaF = Vencimento.Dia3Final
            xQtdDias = Vencimento.Dia3Vencimento
        End If
        xDataI = Format(xDiaI, "00") & "/" & Format(mDataHoje, "mm") & "/" & Format(mDataHoje, "yyyy")
        xDataF = Format(xDiaF, "00") & "/" & Format(mDataHoje, "mm") & "/" & Format(mDataHoje, "yyyy")
        Do Until IsDate(xDataF)
            i = Mid(xDataF, 1, 2) - 1
            Mid(xDataF, 1, 2) = Format(i, "00")
        Loop
        xDataV = Format(CDate(xDataF) + xQtdDias, "dd/mm/yyyy")
        If CDate(xDataF) > mDataHoje Then
            xDataF = Format(fMudaMes(xDataF, "-", 1), "dd/mm/yyyy")
        End If
        Do Until IsDate(xDataI)
            i = Mid(xDataI, 1, 2) - 1
            Mid(xDataI, 1, 2) = Format(i, "00")
        Loop
        If CDate(xDataI) > CDate(xDataF) Then
            xDataI = Format(fMudaMes(xDataI, "-", 1), "dd/mm/yyyy")
        End If
        If Val(Mid(xDataF, 1, 2)) > 27 Then
            Mid(xDataF, 1, 2) = "32"
        End If
        Do Until IsDate(xDataF)
            i = Mid(xDataF, 1, 2) - 1
            Mid(xDataF, 1, 2) = Format(i, "00")
        Loop
        xDataV = Format(CDate(xDataF) + xQtdDias, "dd/mm/yyyy")
        If CDate(xDataV) < CDate(xDataF) Then
            xDataV = Format(fMudaMes(xDataV, "+", 1), "dd/mm/yyyy")
        End If
        Do Until IsDate(xDataV)
            i = Mid(xDataV, 1, 2) - 1
            Mid(xDataV, 1, 2) = Format(i, "00")
        Loop
    End If
    Set Vencimento = Nothing
    CalculaDataAbastecimentoVencimento = xDataI & "|@|" & xDataF & "|@|" & xDataV & "|@|"
    Exit Function
FileError:
    Set Vencimento = Nothing
End Function
Function CalculaDigito5(ByVal x_codigo As Long) As Integer
    Dim x_numero As String
    Dim x_digito As String
    Dim soma As Integer
    Dim tabela As Integer
    x_numero = Format(x_codigo, "00000")
    soma = Mid(x_numero, 1, 1) * 6 + Mid(x_numero, 2, 1) * 5 + Mid(x_numero, 3, 1) * 4 + Mid(x_numero, 4, 1) * 3 + Mid(x_numero, 5, 1) * 2
    tabela = 0
    Do Until soma < tabela
        tabela = tabela + 11
    Loop
    x_digito = tabela - soma
    If x_digito = 0 Then
        x_digito = 1
    End If
    x_digito = Format(x_digito, "00000")
    x_digito = Mid(x_digito, 5, 1)
    CalculaDigito5 = Val(x_digito)
End Function
Function CalculaDigitoCNPJ(ByVal pCNPJ As String) As Boolean
    Dim x_soma As Currency
    Dim i As Integer
    Dim x_digito_1 As Integer
    Dim x_digito_2 As Integer
    x_soma = 0
    For i = 1 To 4
        x_soma = x_soma + (Val(Mid(pCNPJ, i, 1)) * (6 - i))
    Next
    For i = 5 To 12
        x_soma = x_soma + (Val(Mid(pCNPJ, i, 1)) * (14 - i))
    Next
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_1 = 0
    Else
        x_digito_1 = 11 - x_soma
    End If
    x_soma = 0
    For i = 1 To 5
        x_soma = x_soma + (Val(Mid(pCNPJ, i, 1)) * (7 - i))
    Next
    For i = 6 To 12
        x_soma = x_soma + (Val(Mid(pCNPJ, i, 1)) * (15 - i))
    Next
    x_soma = x_soma + (x_digito_1 * 2)
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_2 = 0
    Else
        x_digito_2 = 11 - x_soma
    End If
    If Mid(pCNPJ, 13, 1) = x_digito_1 And Mid(pCNPJ, 14, 1) = x_digito_2 Then
        CalculaDigitoCNPJ = True
    Else
        CalculaDigitoCNPJ = False
        MsgBox "O CNPJ: " & Mid(pCNPJ, 1, 2) & "." & Mid(pCNPJ, 3, 3) & "." & Mid(pCNPJ, 6, 3) & "/" & Mid(pCNPJ, 9, 4) & "-" & Mid(pCNPJ, 13, 2) & " está com o dígito errado!" & Chr(10) & "O dígito correto é: " & Chr(34) & x_digito_1 & x_digito_2 & Chr(34), vbInformation, "Dígito de CGC Inválido!"
    End If
End Function
Function CalculaDigitoCPF(ByVal x_cpf As String) As Boolean
    Dim x_soma As Currency
    Dim i As Integer
    Dim x_digito_1 As Integer
    Dim x_digito_2 As Integer
    x_soma = 0
    For i = 1 To 9
        x_soma = x_soma + (Val(Mid(x_cpf, i, 1)) * (11 - i))
    Next
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_1 = 0
    Else
        x_digito_1 = 11 - x_soma
    End If
    x_soma = 0
    For i = 1 To 9
        x_soma = x_soma + (Val(Mid(x_cpf, i, 1)) * (12 - i))
    Next
    x_soma = x_soma + (x_digito_1 * 2)
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_2 = 0
    Else
        x_digito_2 = 11 - x_soma
    End If
    If Mid(x_cpf, 10, 1) = x_digito_1 And Mid(x_cpf, 11, 1) = x_digito_2 Then
        CalculaDigitoCPF = True
    Else
        CalculaDigitoCPF = False
        MsgBox "O CPF: " & Mid(x_cpf, 1, 3) & "." & Mid(x_cpf, 4, 3) & "." & Mid(x_cpf, 7, 3) & "-" & Mid(x_cpf, 10, 2) & " está com o dígito errado!" & Chr(10) & "O dígito correto é: " & Chr(34) & x_digito_1 & x_digito_2 & Chr(34), vbInformation, "Dígito de CPF Inválido!"
    End If
End Function
Public Function Cat52ConverteDiaBema(ByVal pDia As Integer) As String
    Dim xAsc As Integer
    
    xAsc = 55
    Cat52ConverteDiaBema = pDia
    If pDia >= 10 Then
        xAsc = xAsc + pDia
        Cat52ConverteDiaBema = Chr(xAsc)
    End If
End Function
Public Function Cat52ConverteModeloBema(ByVal pModelo As String) As String
    Cat52ConverteModeloBema = "BEX"
    If pModelo Like "MP-20 FI II*" Then
        Cat52ConverteModeloBema = "BE6"
    ElseIf pModelo Like "MP-40 FI II*" Then
        Cat52ConverteModeloBema = "BE8"
    ElseIf pModelo Like "MP-25 FI" Then
        Cat52ConverteModeloBema = "BEB"
    ElseIf pModelo Like "MP-50 FI*" Then
        Cat52ConverteModeloBema = "BEC"
    ElseIf pModelo Like "MP-2000 TH FI*" Then
        Cat52ConverteModeloBema = "BED"
    ElseIf pModelo Like "MP-2100 TH FI*" Then
        Cat52ConverteModeloBema = "BEE"
    ElseIf pModelo Like "MP-3000 TH FI*" Then
        Cat52ConverteModeloBema = "BEF"
    ElseIf pModelo Like "MP-6000 TH FI*" Then
        Cat52ConverteModeloBema = "BEG"
    ElseIf pModelo Like "MP-6100 TH FI*" Then
        Cat52ConverteModeloBema = "BEH"
    ElseIf pModelo Like "MP-4000 TH FI*" Then
        Cat52ConverteModeloBema = "BEJ"
    End If
End Function
Function fCalculaValores(ByVal pValor As String) As Currency
    Dim xPosicao As Integer
    Dim xTamanho As Integer
    Dim xFlag As Boolean
    Dim xValor1 As Currency
    Dim xValor2 As Currency
    Dim xOperacao As String
    
    fCalculaValores = 0
    If Trim(pValor) = "" Then
        Exit Function
    End If
    
    xValor1 = 0
    xValor2 = 0
    
    xFlag = False
    xTamanho = Len(pValor)
    xPosicao = Len(pValor)
    Do Until xPosicao = 0
        If Not IsNumeric(Mid(pValor, xPosicao, 1)) Then
            If Mid(pValor, xPosicao, 1) = "-" Or Mid(pValor, xPosicao, 1) = "+" Then
                xOperacao = Mid(pValor, xPosicao, 1)
                Exit Do
            End If
        End If
        xPosicao = xPosicao - 1
    Loop
    xValor1 = CCur(Mid(pValor, 1, xPosicao - 1))
    xValor2 = CCur(Mid(pValor, xPosicao + 1, xTamanho - xPosicao))
        
    If xOperacao = "-" Then
        fCalculaValores = xValor1 - xValor2
    ElseIf xOperacao = "+" Then
        fCalculaValores = xValor1 + xValor2
    Else
        fCalculaValores = xValor1
    End If
    
End Function
Public Sub fConsultaPreencheOperador(ByRef pComboBox As ComboBox)
    pComboBox.Clear
    pComboBox.AddItem "Diferente"
    pComboBox.AddItem "Igual"
    pComboBox.AddItem "Maior"
    pComboBox.AddItem "Maior Igual"
    pComboBox.AddItem "Menor"
    pComboBox.AddItem "Menor Igual"
    pComboBox.AddItem "Semelhante"
End Sub
Public Function fConsultaPreparaCondicao(ByVal pNomeCampo As String, ByVal pTipoCampo As Integer, ByVal pOperador As String, ByVal pCondicao As String) As String
    'Prepara pCondicao
    If pTipoCampo = 7 Or pTipoCampo = 8 Or pTipoCampo = 135 Then
        'Data
        pCondicao = preparaData(CDate(pCondicao))
    ElseIf (pTipoCampo = 200 Or pTipoCampo = 202 Or pTipoCampo = 203) And pOperador = "Semelhante" Then
        pCondicao = Chr(39) & "%" & pCondicao & "%" & Chr(39)
    ElseIf (pTipoCampo = 200 Or pTipoCampo = 202 Or pTipoCampo = 203) And pOperador <> "Semelhante" Then
        pCondicao = preparaTexto(pCondicao)
    End If
    
    'Prepara pOperador
    If pOperador = "Diferente" Then
        pOperador = "<>"
    ElseIf pOperador = "Igual" Then
        pOperador = "="
    ElseIf pOperador = "Maior" Then
        pOperador = ">"
    ElseIf pOperador = "Maior Igual" Then
        pOperador = ">="
    ElseIf pOperador = "Menor" Then
        pOperador = "<"
    ElseIf pOperador = "Menor Igual" Then
        pOperador = "<="
    ElseIf pOperador = "Semelhante" Then
        pOperador = "Like"
    End If
    
    fConsultaPreparaCondicao = "[" & pNomeCampo & "] " & pOperador & " " & pCondicao
End Function
Function fContemLetra(ByVal x_string As String) As Boolean
    Dim i As Integer
    fContemLetra = False
    i = Len(x_string)
    If i = 0 Then
        Exit Function
    End If
    For i = 1 To Len(x_string)
        If Mid(x_string, i, 1) >= "A" And Mid(x_string, i, 1) <= "Z" Or Mid(x_string, i, 1) >= "a" And Mid(x_string, i, 1) <= "z" Then
            fContemLetra = True
            Exit For
        End If
    Next
End Function
Function GravaAuditoria(ByVal pSequencia As Integer, ByVal pNomeInternoPrograma As String, ByVal pOperacao As Integer, ByVal pObservacao As String) As Boolean
    Dim Auditoria As New cAuditoria
    Auditoria.CodigoUsuario = g_usuario
    Auditoria.Data = Format(Date, "dd/mm/yyyy")
    Auditoria.Hora = Format(Time, "hh:mm:ss")
    Auditoria.Sequencia = pSequencia
    Auditoria.NomeInternoPrograma = pNomeInternoPrograma
    Auditoria.Operacao = pOperacao
    Auditoria.Computador = GetIPHostName() & " - " & GetIPAddress()
    If Len(Auditoria.Computador) > 30 Then
        Dim xQtdMaior As Integer
        Dim xQtdNomePc As Integer
        xQtdMaior = Len(Auditoria.Computador) - 30
        xQtdNomePc = Len(GetIPHostName())
        Auditoria.Computador = Mid(GetIPHostName(), 1, xQtdNomePc - xQtdMaior) & " - " & GetIPAddress()
    End If
    Auditoria.Observacao = pObservacao
    GravaAuditoria = Auditoria.Incluir
    Set Auditoria = Nothing
End Function
Function GravaConfiguracaoConsulta(ByVal pPrograma As String, ByVal pCampo As String, ByVal pOperador As String, ByVal pCondicao As String) As Boolean
    Dim lSQL As String
    Dim xPrograma As String
    
    On Error GoTo FileError
    
    xPrograma = Mid(pPrograma, 1, 40)
    GravaConfiguracaoConsulta = False
    lSQL = "DELETE FROM Configuracao_Pesquisa"
    lSQL = lSQL & " WHERE Programa = " & preparaTexto(xPrograma)
    Conectar.ExecutaSql (lSQL)
    
    lSQL = "INSERT INTO Configuracao_Pesquisa ( Programa, Campo, Operador, Condicao) VALUES ("
    lSQL = lSQL & preparaTexto(xPrograma) & ", "
    lSQL = lSQL & preparaTexto(pCampo) & ", "
    lSQL = lSQL & preparaTexto(pOperador) & ", "
    lSQL = lSQL & preparaTexto(pCondicao) & ")"
    If Conectar.ExecutaSql(lSQL) > 0 Then
        GravaConfiguracaoConsulta = True
    End If
    Exit Function
FileError:
    MsgBox Error, Err
End Function
Sub Imprime(ByVal f_coluna As Currency, ByVal f_linha As Currency, ByVal f_string As String, ByVal f_local As Integer)
    If f_local = 0 Then
        'emissao_movimentacao_diaria.Picture1.CurrentX = f_coluna '/ 2
        'emissao_movimentacao_diaria.Picture1.CurrentY = f_linha '/ 2
        'emissao_movimentacao_diaria.Picture1.Print f_string
    Else
        Printer.CurrentX = f_coluna
        Printer.CurrentY = f_linha
        Printer.Print f_string
    End If
End Sub
Sub centerform(AnyForm As Form)
    If AnyForm.WindowState = 0 Then
        AnyForm.Top = (Screen.Height - AnyForm.Height) / 2
        AnyForm.Left = (Screen.Width - AnyForm.Width) / 2
    End If
End Sub
Function ImprimeCentralizadoB(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim x_string As String
    x_string = f_string
    If Val(x_string) = 0 Then
        x_string = ""
    End If
    ImprimeCentralizadoB = x_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(x_string)) / 2
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeCentralizadoB, f_local
End Function
Function ImprimeString(ByVal f_string As String, ByVal f_coluna As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeString = f_string
    Printer.CurrentX = f_coluna
    Printer.CurrentY = f_linha
    Imprime f_coluna, f_linha, ImprimeString, f_local
End Function
Function fMudaMes(ByVal mData As String, mOperacao As String, mQuantidade As Integer) As Date
    Dim xDia As Integer
    Dim xMes As Integer
    Dim xAno As Integer
    Dim xDataTeste As String
    xDia = Day(CDate(mData))
    xMes = Month(CDate(mData))
    xAno = Year(CDate(mData))
    If mOperacao = "+" Then
        If xMes = 12 Then
            xMes = 0
            xAno = xAno + 1
        End If
        xMes = xMes + 1
    Else
        If xMes = 1 Then
            xMes = 13
            xAno = xAno - 1
        End If
        xMes = xMes - 1
    End If
    xDataTeste = Format(xDia, "00") & "/" & Format(xMes, "00") & "/" & Format(xAno, "0000")
    Do Until IsDate(xDataTeste)
        xDia = xDia - 1
        Mid(xDataTeste, 1, 2) = Format(xDia, "00")
    Loop
    fMudaMes = CDate(xDataTeste)
End Function
Function fRetiraString(ByVal pString As String, ByVal pNumero As Integer) As String
    fRetiraString = ""
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(pString) > 0 Then
        Do Until x_index > Len(pString)
            x_index = x_index + 1
            If Mid(pString, x_index, 3) = "|@|" Then
                If x_numero = pNumero Then
                    fRetiraString = Mid(pString, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Function fTiraCodigoProduto5(ByVal f_codigo As String, ByVal f_tipo As String) As Long
    fTiraCodigoProduto5 = 0
    If Len(f_codigo) >= 4 Then
        If Len(f_codigo) = 4 Then
            f_codigo = "0" & f_codigo
        End If
        If f_tipo = "g" Then
            fTiraCodigoProduto5 = Val(Mid(f_codigo, 1, 2))
        ElseIf f_tipo = "c" Then
            fTiraCodigoProduto5 = Val(Mid(f_codigo, 3, 3))
        End If
    End If
End Function
Function ImprimeCentralizado(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeCentralizado = f_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(f_string)) / 2
    If f_local = 1 Then
        Printer.CurrentX = f_coluna_i
        Printer.CurrentY = f_linha
    Else
        'emissao_movimentacao_diaria.Picture1.CurrentX = f_coluna_i
        'emissao_movimentacao_diaria.Picture1.CurrentY = f_linha
    End If
    Imprime f_coluna_i, f_linha, ImprimeCentralizado, f_local
End Function
Function ImprimeValorB(ByVal f_valor As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim x_string As String
    x_string = f_valor
    If Val(x_string) = 0 Then
        x_string = ""
    End If
    ImprimeValorB = x_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(x_string) - 0.08)
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeValorB, f_local
End Function
Function fValidaValor(ByVal pValor As String) As Currency
    Dim x_posicao As Integer
    Dim x_tamanho As Integer
    Dim x_flag As Boolean
    Dim xNegativo As Boolean
    
    x_flag = False
    xNegativo = False
    If pValor <> "" Then
        If Mid(pValor, 1, 1) = "-" Then
            xNegativo = True
        End If
        x_tamanho = Len(pValor)
        x_posicao = Len(pValor)
        Do Until x_posicao = 0
            If Not IsNumeric(Mid(pValor, x_posicao, 1)) Then
                If x_flag = False Then
                    Mid(pValor, x_posicao, 1) = "."
                    x_flag = True
                Else
                    pValor = " " & Mid(pValor, 1, x_posicao - 1) & Mid(pValor, x_posicao + 1, x_tamanho - x_posicao)
                End If
            End If
            x_posicao = x_posicao - 1
        Loop
    End If
    fValidaValor = Val(Trim(pValor))
    If xNegativo Then
        fValidaValor = -fValidaValor
    End If
End Function
Function fValidaValor1(ByVal pValor As String) As Currency
    Dim i As Integer
    If pValor <> "" Then
        i = Len(Trim(pValor)) - 1
        Mid(pValor, i, 1) = "."
    End If
    If Len(Trim(pValor)) = 7 Then
        Mid(pValor, 1, 2) = " " & Mid(pValor, 1, 1)
    ElseIf Len(Trim(pValor)) = 8 Then
        Mid(pValor, 1, 3) = " " & Mid(pValor, 1, 2)
    ElseIf Len(Trim(pValor)) = 9 Then
        Mid(pValor, 1, 4) = " " & Mid(pValor, 1, 3)
    End If
    fValidaValor1 = Val(Trim(pValor))
End Function
Function fValidaValor2(ByVal pValor As String) As Currency
    Dim i As Integer
    If pValor <> "" Then
        i = Len(Trim(pValor)) - 2
        Mid(pValor, i, 1) = "."
    End If
    If Len(Trim(pValor)) = 8 Then
        Mid(pValor, 1, 2) = " " & Mid(pValor, 1, 1)
    ElseIf Len(Trim(pValor)) = 9 Then
        Mid(pValor, 1, 3) = " " & Mid(pValor, 1, 2)
    ElseIf Len(Trim(pValor)) = 10 Then
        Mid(pValor, 1, 4) = " " & Mid(pValor, 1, 3)
    End If
    fValidaValor2 = Val(Trim(pValor))
End Function
Function fValidaValor4(ByVal pValor As String) As Currency
    Dim i As Integer
    If pValor <> "" Then
        i = Len(Trim(pValor)) - 4
        Mid(pValor, i, 1) = "."
    End If
    If Len(Trim(pValor)) = 10 Then
        Mid(pValor, 1, 2) = " " & Mid(pValor, 1, 1)
    ElseIf Len(Trim(pValor)) = 11 Then
        Mid(pValor, 1, 3) = " " & Mid(pValor, 1, 2)
    ElseIf Len(Trim(pValor)) = 12 Then
        Mid(pValor, 1, 4) = " " & Mid(pValor, 1, 3)
    End If
    fValidaValor4 = Val(Trim(pValor))
End Function
Function fVerificaSomaSubtracao(ByVal pValor As String) As Boolean
    Dim i As Integer
    
    fVerificaSomaSubtracao = False
    If Len(Trim(pValor)) = 0 Then
        Exit Function
    End If
    For i = 1 To Len(Trim(pValor))
        If Mid(pValor, i, 1) = "-" Or Mid(pValor, i, 1) = "+" Then
            fVerificaSomaSubtracao = True
            Exit Function
        End If
    Next
End Function
Function fMascaraCEP(ByVal pCEP As String) As String
    If Len(Trim(pCEP)) = 8 Then
        fMascaraCEP = Format(pCEP, "@@.@@@-@@@")
    Else
        fMascaraCEP = pCEP
    End If
End Function
Function fMascaraCPF(ByVal pCPF As String) As String
    If Len(Trim(pCPF)) = 11 Then
        fMascaraCPF = Format(pCPF, "@@@.@@@.@@@-@@")
    Else
        fMascaraCPF = pCPF
    End If
End Function
Function fMascaraData(ByVal pData As String) As String
    Dim i As Integer

    For i = 1 To Len(pData)
        If Mid(pData, i, 1) = "/" Then
            fMascaraData = Format(CDate(pData), "dd/mm/yyyy")
            Exit Function
        End If
    Next
    i = Len(pData)
    If i = 8 Then
        fMascaraData = Mid(pData, 1, 2) & "/" & Mid(pData, 3, 2) & "/" & Mid(pData, 5, 4)
    ElseIf i = 7 Then
        fMascaraData = "0" & Mid(pData, 1, 1) & "/" & Mid(pData, 2, 2) & "/" & Mid(pData, 4, 4)
    ElseIf i = 6 Then
        fMascaraData = Mid(pData, 1, 2) & "/" & Mid(pData, 3, 2) & "/20" & Mid(pData, 5, 4)
    ElseIf i = 5 Then
        fMascaraData = "0" & Mid(pData, 1, 1) & "/" & Mid(pData, 2, 2) & "/20" & Mid(pData, 4, 4)
    Else
        fMascaraData = pData
    End If
End Function
Public Function fMascaraHora(ByVal pHora As String) As String
    fMascaraHora = ""
    If IsDate(pHora) Then
        fMascaraHora = Format(pHora, "HH:mm:ss")
    End If
End Function
Public Sub FinalizaProcessoCaixa()
    Dim xArquivoDiscoTMP As TextStream
    
    On Error GoTo FileError
    
    Set xArquivoDiscoTMP = gArqTxt.CreateTextFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP")
    xArquivoDiscoTMP.WriteLine ("[Outras]")
    xArquivoDiscoTMP.WriteLine ("gString=" & g_string)
    xArquivoDiscoTMP.Close
    
    Call gArqTxt.MoveFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP", "C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.INI")
    Exit Sub
FileError:
End Sub
Function fMascaraCNPJ(cnpj_x As String) As String
    If Len(Trim(cnpj_x)) = 14 Then
        fMascaraCNPJ = Mid(cnpj_x, 1, 2) & "." & Mid(cnpj_x, 3, 3) & "." & Mid(cnpj_x, 6, 3) & "/" & Mid(cnpj_x, 9, 4) & "-" & Mid(cnpj_x, 13, 2)
    Else
        fMascaraCNPJ = cnpj_x
    End If
End Function
Function fMascaraContaContabil(ByVal nConta As String) As String
    If Len(Trim(nConta)) = 1 Then
        fMascaraContaContabil = Mid(nConta, 1, 1)
    ElseIf Len(Trim(nConta)) = 2 Then
        fMascaraContaContabil = Mid(nConta, 1, 1) & "." & Mid(nConta, 2, 1)
    ElseIf Len(Trim(nConta)) = 3 Then
        fMascaraContaContabil = Mid(nConta, 1, 1) & "." & Mid(nConta, 2, 1) & "." & Mid(nConta, 3, 1)
    ElseIf Len(Trim(nConta)) = 5 Then
        fMascaraContaContabil = Mid(nConta, 1, 1) & "." & Mid(nConta, 2, 1) & "." & Mid(nConta, 3, 1) & "." & Mid(nConta, 4, 2)
    ElseIf Len(Trim(nConta)) = 9 Then
        fMascaraContaContabil = Mid(nConta, 1, 1) & "." & Mid(nConta, 2, 1) & "." & Mid(nConta, 3, 1) & "." & Mid(nConta, 4, 2) & "." & Mid(nConta, 6, 4)
    Else
        fMascaraContaContabil = nConta
    End If
End Function
Function fMascaraTelefone(ByVal telefone_x As String) As String
    If Len(Trim(telefone_x)) = 7 Then
        fMascaraTelefone = Mid(telefone_x, 1, 3) & "-" & Mid(telefone_x, 4, 4)
    ElseIf Len(Trim(telefone_x)) = 8 Then
        fMascaraTelefone = Mid(telefone_x, 1, 4) & "-" & Mid(telefone_x, 5, 4)
    ElseIf Len(Trim(telefone_x)) = 10 Then
        fMascaraTelefone = "(" & Mid(telefone_x, 1, 3) & ") " & Mid(telefone_x, 4, 3) & "-" & Mid(telefone_x, 7, 4)
    ElseIf Len(Trim(telefone_x)) = 11 Then
        fMascaraTelefone = "(" & Mid(telefone_x, 1, 3) & ") " & Mid(telefone_x, 4, 4) & "-" & Mid(telefone_x, 8, 4)
    Else
        fMascaraTelefone = telefone_x
    End If
End Function
Function fDesmascaraContaContabil(ByVal nConta As String) As String
    If Len(Trim(nConta)) = 1 Then
        fDesmascaraContaContabil = Mid(nConta, 1, 1)
    ElseIf Len(Trim(nConta)) = 3 Then
        fDesmascaraContaContabil = Mid(nConta, 1, 1) & Mid(nConta, 3, 1)
    ElseIf Len(Trim(nConta)) = 6 Then
        fDesmascaraContaContabil = Mid(nConta, 1, 1) & Mid(nConta, 3, 1) & Mid(nConta, 5, 2)
    ElseIf Len(Trim(nConta)) = 9 Then
        fDesmascaraContaContabil = Mid(nConta, 1, 1) & Mid(nConta, 3, 1) & Mid(nConta, 5, 2) & Mid(nConta, 8, 2)
    ElseIf Len(Trim(nConta)) = 13 Then
        fDesmascaraContaContabil = Mid(nConta, 1, 1) & Mid(nConta, 3, 1) & Mid(nConta, 5, 2) & Mid(nConta, 8, 2) & Mid(nConta, 11, 3)
    Else
        fDesmascaraContaContabil = nConta
    End If
End Function
Function fDesmascaraCPF(ByVal pCPF As String) As String
    If Len(Trim(pCPF)) = 14 Then
        fDesmascaraCPF = Mid(pCPF, 1, 3) & Mid(pCPF, 5, 3) & Mid(pCPF, 9, 3) & Mid(pCPF, 13, 2)
    Else
        fDesmascaraCPF = pCPF
    End If
End Function
Function fDesmascaraData(ByVal pData As String) As String
    If Len(pData) = 10 Then
        fDesmascaraData = Mid(pData, 1, 2) & Mid(pData, 4, 2) & Mid(pData, 7, 4)
    Else
        fDesmascaraData = pData
    End If
End Function
Function fDataPrimeiroDiaMesAnterior(ByVal pData As Date) As String
    Dim xData As String
    On Error GoTo FileError
    
    xData = Format(pData, "dd/mm/yyyy")
    Mid(xData, 1, 2) = "15"
    xData = Format(CDate(xData) - 31, "dd/mm/yyyy")
    Mid(xData, 1, 2) = "01"
    fDataPrimeiroDiaMesAnterior = xData
    Exit Function
FileError:
    fDataPrimeiroDiaMesAnterior = Format(pData, "dd/mm/yyyy")
End Function
Function fDataUltimoDiaMesAnterior(ByVal pData As Date) As String
    Dim xData As String
    On Error GoTo FileError
    
    xData = Format(pData, "dd/mm/yyyy")
    Mid(xData, 1, 2) = "15"
    xData = Format(CDate(xData) - 31, "dd/mm/yyyy")
    Mid(xData, 1, 2) = "31"
    Do Until IsDate(xData)
        Mid(xData, 1, 2) = Val(Mid(xData, 1, 2)) - 1
    Loop
    
    fDataUltimoDiaMesAnterior = xData
    Exit Function
FileError:
    fDataUltimoDiaMesAnterior = Format(pData, "dd/mm/yyyy")
End Function
Function fDesmascaraCNPJ(ByVal pCNPJ As String) As String
    If Len(Trim(pCNPJ)) = 18 Then
        fDesmascaraCNPJ = Mid(pCNPJ, 1, 2) & Mid(pCNPJ, 4, 3) & Mid(pCNPJ, 8, 3) & Mid(pCNPJ, 12, 4) & Mid(pCNPJ, 17, 2)
    Else
        fDesmascaraCNPJ = pCNPJ
    End If
End Function
Function fDesmascaraNumero(ByVal pNumeroMascarado As String) As Long
    Dim xNumero As String
    Dim i As Integer
    
    xNumero = ""
    For i = 1 To Len(pNumeroMascarado)
        If Mid(pNumeroMascarado, i, 1) >= "0" And Mid(pNumeroMascarado, i, 1) <= "9" Then
            xNumero = xNumero & Mid(pNumeroMascarado, i, 1)
        End If
    Next
    fDesmascaraNumero = CLng(xNumero)
End Function
Function fDesmascaraNumeroString(ByVal pNumeroMascarado As String) As String
    Dim xNumero As String
    Dim i As Integer
    
    xNumero = ""
    For i = 1 To Len(pNumeroMascarado)
        If Mid(pNumeroMascarado, i, 1) >= "0" And Mid(pNumeroMascarado, i, 1) <= "9" Then
            xNumero = xNumero & Mid(pNumeroMascarado, i, 1)
        End If
    Next
    fDesmascaraNumeroString = xNumero
End Function
Function fDesmascaraTelefone(ByVal telefone_x As String) As String
    If Len(Trim(telefone_x)) = 8 Then
        fDesmascaraTelefone = Mid(telefone_x, 1, 3) & Mid(telefone_x, 5, 4)
    ElseIf Len(Trim(telefone_x)) = 9 Then
        fDesmascaraTelefone = Mid(telefone_x, 1, 4) & Mid(telefone_x, 6, 4)
    ElseIf Len(Trim(telefone_x)) = 14 Then
        fDesmascaraTelefone = Mid(telefone_x, 2, 3) & Mid(telefone_x, 7, 3) & Mid(telefone_x, 11, 4)
    ElseIf Len(Trim(telefone_x)) = 15 Then
        fDesmascaraTelefone = Mid(telefone_x, 2, 3) & Mid(telefone_x, 7, 4) & Mid(telefone_x, 12, 4)
    Else
        fDesmascaraTelefone = telefone_x
    End If
End Function
Function fUsaNFCe() As Boolean
    Dim dados As String
    
    fUsaNFCe = False
    
    On Error GoTo FileError
    
    
    dados = ReadINI("NFCe", "Emite NFCe", ArqSgpIni)
    If dados = "SIM" Then
        fUsaNFCe = True
    End If
    
    Exit Function
FileError:
    Exit Function
End Function
Function fEcfInstalada() As Boolean
    Dim dados As String
    'Dim NumeroArquivo As Integer
    'NumeroArquivo = FreeFile
    fEcfInstalada = False
    
    On Error GoTo FileError
    
    'Open "C:\VB5\SGP\CUPOM_DEMONSTRACAO.TXT" For Input As NumeroArquivo
    'Do Until EOF(NumeroArquivo)
    '    Line Input #NumeroArquivo, dados
    '    If Mid(dados, 1, 13) = "ECF INSTALADA" Then
    '        fEcfInstalada = True
    '    End If
    'Loop
    'Close #NumeroArquivo
    
    dados = ReadINI("CUPOM FISCAL", "ECF Instalada", ArqSgpIni)
    If dados = "SIM" Then
        fEcfInstalada = True
    End If
    
    
    Exit Function
FileError:
    Exit Function
End Function
Function ImprimeValor(ByVal f_valor As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeValor = f_valor
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(f_valor) - 0.08)
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeValor, f_local
End Function
Function ImprimeTexto(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim i As Integer
    ImprimeTexto = ""
    f_coluna_i = f_coluna_i + 0.08
    f_coluna_f = f_coluna_f - 0.08
    Do
        If Trim(f_string) = "" Then
            Exit Do
        End If
        i = i + 1
        If Printer.TextWidth(Mid(f_string, 1, i)) >= (f_coluna_f - f_coluna_i) Then
            Exit Do
        End If
        If Printer.TextWidth(f_string) = 0 Then
            Exit Do
        End If
        ImprimeTexto = Mid(f_string, 1, i)
        If Len(f_string) = i Then
            Exit Do
        End If
    Loop
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeTexto, f_local
End Function
Public Function EnviaMensagemEmail(ByVal pEmpresa As String, ByVal pNomeEmpresa As String, ByVal pAssunto As String, ByVal pMensagem As String, ByVal pConcluido As Boolean, ByVal pNumeroEmailInicial As Long) As Boolean
    Dim xString As String
    Dim xNomeArquivo As String
    Dim xArquivoTmp As String
    Dim xArquivoDiscoTMP As TextStream
    Dim retval As Long
    Dim EmailEnviado As New cEmailEnviado
    
    On Error GoTo FileError
    
    EnviaMensagemEmail = False
    If gNumeroEmailInicial = 0 Then
        EmailEnviado.Empresa = pEmpresa
        EmailEnviado.numero = 1
        EmailEnviado.Remetente = pNomeEmpresa
        EmailEnviado.Destinatario = "Cerrado Tecnologia"
        EmailEnviado.EmailRemetente = "enviaversao@cerradoinformatica.com.br"
        EmailEnviado.EmailDestinatario = "tassoteixeira@outlook.com"
        EmailEnviado.Assunto = Mid(pNomeEmpresa, 1, 15) & ":" & pAssunto
        EmailEnviado.Mensagem = pMensagem & vbCrLf & "Usuário:" & g_nome_usuario
        EmailEnviado.CriadoEm = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
        EmailEnviado.EnviadoEm = EmailEnviado.CriadoEm
        EmailEnviado.IPRemetente = GetIPAddress()
        EmailEnviado.EmailEnviado = False
        EmailEnviado.EmailConcluido = pConcluido
        EnviaMensagemEmail = EmailEnviado.Incluir
    Else
        If EmailEnviado.LocalizarCodigo(g_empresa, gNumeroEmailInicial) Then
            EmailEnviado.Mensagem = EmailEnviado.Mensagem & vbCrLf & pMensagem
            EmailEnviado.EmailConcluido = pConcluido
            pConcluido = EmailEnviado.Alterar(g_empresa, gNumeroEmailInicial)
        Else
            gNumeroEmailInicial = 0
        End If
    End If
    If pConcluido = False Then
        gNumeroEmailInicial = EmailEnviado.numero
    Else
        gNumeroEmailInicial = 0
    End If
    
    If gNumeroEmailInicial = 0 Then
        EmailEnviado.EmailDestinatario = "cerradoinfo@gmail.com"
        EmailEnviado.CriadoEm = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
        EmailEnviado.EnviadoEm = EmailEnviado.CriadoEm
        EnviaMensagemEmail = EmailEnviado.Incluir
    Else
        If EmailEnviado.LocalizarCodigo(g_empresa, gNumeroEmailInicial + 1) Then
            EmailEnviado.Mensagem = EmailEnviado.Mensagem & vbCrLf & pMensagem
            EmailEnviado.EmailConcluido = pConcluido
            pConcluido = EmailEnviado.Alterar(g_empresa, gNumeroEmailInicial + 1)
        End If
    End If
    Set EmailEnviado = Nothing
    
    
    xNomeArquivo = "C:\vb5\sgp\sgp_cadastro.ini"
    retval = Len(xNomeArquivo)
    xArquivoTmp = Mid(xNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    
    
    Set xArquivoDiscoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    xArquivoDiscoTMP.WriteLine ("[Empresa]")
    xArquivoDiscoTMP.WriteLine ("Empresa=" & Format(pEmpresa, "000"))
    xArquivoDiscoTMP.WriteLine ("NomeEmpresa=" & pNomeEmpresa)
    'xArquivoDiscoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    xArquivoDiscoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    
    xArquivoDiscoTMP.WriteLine (" ")
    xArquivoDiscoTMP.WriteLine ("[Programa]")
    xArquivoDiscoTMP.WriteLine ("Programa=" & "MovimentoEnviaMensagemEmailRotina")
    
    xArquivoDiscoTMP.WriteLine (" ")
    xArquivoDiscoTMP.WriteLine ("[Email]")
    xArquivoDiscoTMP.WriteLine ("Assunto=" & Mid(pNomeEmpresa, 1, 15) & ":" & pAssunto)
    xArquivoDiscoTMP.WriteLine ("Mensagem=" & pMensagem & Chr(10) & "Usuário:" & g_nome_usuario)
    
    xArquivoDiscoTMP.Close
    
    
    Call gArqTxt.MoveFile(xArquivoTmp, xNomeArquivo)

    EnviaMensagemEmail = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar Sgp_Cadastro.ini!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Sub ErroArquivo(ByVal x_arquivo As String, ByVal x_nome As String)
    x_arquivo = UCase(x_arquivo)
    If Err = 3022 Then
        MsgBox Mid(x_nome, 1, Len(x_nome) - 1) & " já Cadastrad" & Mid(x_nome, Len(x_nome), 1) & "!" & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3019 Then
        MsgBox "Índice não informado!" & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3058 Then
        MsgBox "Campo do índice não informado!" & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3265 Then
        MsgBox "Variável inexistente!" & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3315 Then
        MsgBox "Campo com valor nulo!" & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    Else
        MsgBox "Erro de número " & Err & Chr(10) & Error, vbInformation, "Erro na Tabela: " & x_arquivo
        Exit Sub
    End If
End Sub
Sub CentraForm(unForm As Form)
    unForm.Left = (Screen.Width - unForm.Width) / 2
    unForm.Top = (Screen.Height - unForm.Height) / 2
End Sub
Function ChamaDrive() As Boolean
    On Error GoTo FileError
    
    ChamaDrive = False
    gDrive = ReadINI("LOCAL", "Drive", ArqSgpIni)
    gDiretorioData = ReadINI("LOCAL", "Diretorio BD", ArqSgpIni)
    gDiretorioAplicativo = ReadINI("LOCAL", "Diretorio Aplicativo", ArqSgpIni)
    
    gNomeBancoDados = ReadINI("LOCAL", "Nome do Banco de Dados", ArqSgpIni)
    gNomeInternoBD = ReadINI("LOCAL", "Nome Interno BD", ArqSgpIni)
    ChDrive gDrive
    ChDir gDiretorioData
    ChamaDrive = True
    Exit Function

FileError:
    MsgBox "Não foi possível definir a unidade: " & gDrive & vbCrLf & "Se o problema continuar não será possível acessar o banco de dados." & Chr(10) & Chr(10) & "Pode ser que a unidade " & preparaTexto(gDrive) & " não esteja mapeada na rede." & vbCrLf & ArqSgpIni, vbCritical, "Erro de definição de Unidade!"
End Function
Function FazExtenso(ByVal pValor As Currency) As String
    FazExtenso = ""
    'Esta rotina faz o extenso até o limite de
    '900.000,00 (Novecentos Mil reais)
    Dim x_extenso As String
    Dim Valor As String
    'Cria variáveis
    Dim Unidade(1 To 9) As String
    Dim dezena_1(0 To 9) As String
    Dim dezena(2 To 9) As String
    Dim centena(1 To 9) As String
    'unidade
    Unidade(1) = "Um "
    Unidade(2) = "Dois "
    Unidade(3) = "Três "
    Unidade(4) = "Quatro "
    Unidade(5) = "Cinco "
    Unidade(6) = "Seis "
    Unidade(7) = "Sete "
    Unidade(8) = "Oito "
    Unidade(9) = "Nove "
    'dezena_1
    dezena_1(0) = "Dez "
    dezena_1(1) = "Onze "
    dezena_1(2) = "Doze "
    dezena_1(3) = "Treze "
    dezena_1(4) = "Quatrorze "
    dezena_1(5) = "Quinze "
    dezena_1(6) = "Dezesseis "
    dezena_1(7) = "Dezesete "
    dezena_1(8) = "Dezoito "
    dezena_1(9) = "Dezenove "
    'dezena
    dezena(2) = "Vinte "
    dezena(3) = "Trinta "
    dezena(4) = "Quarenta "
    dezena(5) = "Cinquenta "
    dezena(6) = "Sessenta "
    dezena(7) = "Setenta "
    dezena(8) = "Oitenta "
    dezena(9) = "Noventa "
    'centena
    centena(1) = "Cento "
    centena(2) = "Duzentos "
    centena(3) = "Trezentos "
    centena(4) = "Quatrocentos "
    centena(5) = "Quinhentos "
    centena(6) = "Seiscentos "
    centena(7) = "Setecentos "
    centena(8) = "Oitocentos "
    centena(9) = "Novecentos "
    
    x_extenso = ""
    Valor = Format(pValor, "0000000000.00")
    
    '*** início (1000 a 999000) ***
    'testa centena de milhar de (100 000 a 900 000)
    If Mid(Valor, 5, 1) > 0 Then
        If Mid(Valor, 5, 3) = 100 Then
            x_extenso = x_extenso & "Cem Mil "
        Else
            x_extenso = x_extenso & centena(Mid(Valor, 5, 1))
            If Mid(Valor, 6, 2) > 0 Then
                x_extenso = x_extenso & "e "
            End If
        End If
    End If
    'testa dezena de milhar de (10 a 19)
    If Mid(Valor, 6, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 7, 1)) & "Mil "
    End If
    'testa dezena de milhar (20 a 90)
    If Mid(Valor, 6, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 6, 1))
        If Mid(Valor, 7, 1) > 0 Then
            x_extenso = x_extenso & "e "
        Else
            x_extenso = x_extenso & "Mil "
        End If
    End If
    'testa unidade de milhar (1 a 9) e dezena de milhar <> 1
    If Mid(Valor, 7, 1) > 0 And Mid(Valor, 6, 1) <> 1 Then
        x_extenso = x_extenso & Unidade(Mid(Valor, 7, 1)) & "Mil "
    End If
    'testa de milhar(xxx xxx) e > 0 e unidade(xxx) > 0
    If Mid(Valor, 5, 3) > 0 And Mid(Valor, 8, 3) > 0 Then
        x_extenso = x_extenso & "e "
    End If
    '*** fim (1000 a 999000) ***
    
    
    '*** início (1 a 999) ***
    'testa centena de (100 a 900)
    If Mid(Valor, 8, 1) > 0 Then
        If Mid(Valor, 8, 3) = 100 Then
            x_extenso = x_extenso & "Cem "
        Else
            x_extenso = x_extenso & centena(Mid(Valor, 8, 1))
            If Mid(Valor, 9, 2) > 0 Then
                x_extenso = x_extenso & "e "
            End If
        End If
    End If
    'testa dezena de (10 a 19)
    If Mid(Valor, 9, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 10, 1))
    End If
    'testa dezena de (20 a 90)
    If Mid(Valor, 9, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 9, 1))
        If Mid(Valor, 10, 1) > 0 Then
            x_extenso = x_extenso & "e "
        End If
    End If
    'testa unidade de (1 a 9) e dezena <> 1
    If Mid(Valor, 10, 1) > 0 And Mid(Valor, 9, 1) <> 1 Then
        x_extenso = x_extenso & Unidade(Mid(Valor, 10, 1))
    End If
    '*** fim (1 a 999) ***
    
    
    'testa se é Real ou Reais
    If Mid(Valor, 1, 10) = 1 Then
        x_extenso = x_extenso & "Real "
    ElseIf Mid(Valor, 1, 10) > 1 Then
        x_extenso = x_extenso & "Reais "
    End If
    
    
    '*** início Centavos (1 a 99) ***
    'testa dezena de (10 a 19)
    If Mid(Valor, 12, 2) > 0 Then
        x_extenso = x_extenso & "e "
    End If
    'testa dezena de centavos de (10 a 19)
    If Mid(Valor, 12, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 13, 1))
    End If
    'testa dezena de centavos de (20 a 90)
    If Mid(Valor, 12, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 12, 1))
        If Mid(Valor, 13, 1) > 0 Then
            x_extenso = x_extenso & "e "
        End If
    End If
    'testa unidade de centavos de (1 a 9) e dezena de centavos <> 1
    If Mid(Valor, 13, 1) > 0 And Mid(Valor, 12, 1) <> 1 Then
        x_extenso = x_extenso & Unidade(Mid(Valor, 13, 1))
    End If
    '*** fim Centavos (1 a 99) ***
    
    
    'testa se é Centavo ou Centavos
    If Mid(Valor, 12, 2) = 1 Then
        x_extenso = x_extenso & "Centavo "
    ElseIf Mid(Valor, 12, 2) > 1 Then
        x_extenso = x_extenso & "Centavos "
    End If
    
    Mid(x_extenso, Len(Trim(x_extenso)) + 1, 1) = "."
    FazExtenso = x_extenso
End Function
Sub FazExtensoMes(ByVal data_x As String, mes_x As String)
    Select Case Month(data_x)
        Case 1
            mes_x = "Janeiro  "
        Case 2
            mes_x = "Fevereiro"
        Case 3
            mes_x = "Março    "
        Case 4
            mes_x = "Abril    "
        Case 5
            mes_x = "Maio     "
        Case 6
            mes_x = "Junho    "
        Case 7
            mes_x = "Julho    "
        Case 8
            mes_x = "Agosto   "
        Case 9
            mes_x = "Setembro "
        Case 10
            mes_x = "Outubro  "
        Case 11
            mes_x = "Novembro "
        Case 12
            mes_x = "Dezembro "
    End Select
End Sub
Public Function NivelAcesso(ByVal pNivelAcesso As Integer) As String
    If pNivelAcesso = 1 Then
        NivelAcesso = "Desenvolvimento"
    ElseIf pNivelAcesso = 2 Then
        NivelAcesso = "Diretoria"
    ElseIf pNivelAcesso = 3 Then
        NivelAcesso = "Gerência"
    ElseIf pNivelAcesso = 4 Then
        NivelAcesso = "Operação"
    ElseIf pNivelAcesso = 5 Then
        NivelAcesso = "Digitação"
    End If
End Function
Sub PaiFilho(ByRef pai_x As String, ByRef filho_x As String)
    Dim x_tamanho As Integer
    Dim x_posicao As Integer
    Dim x_posicao2 As Integer
    Dim x_pai As String
    Dim x_filho As String
    If pai_x = "\" Then
        pai_x = "\"
        filho_x = "\"
        Exit Sub
    End If
    x_tamanho = Len(Trim(pai_x))
    x_posicao = x_tamanho
    Do Until x_posicao = 0
        If Mid(pai_x, x_posicao, 1) = "\" Then
            x_filho = Mid(pai_x, (x_posicao + 1), (x_tamanho - x_posicao))
            Exit Do
        End If
        x_posicao = x_posicao - 1
    Loop
    x_tamanho = x_tamanho - (Len(Trim(x_filho)) + 1)
    x_posicao = x_tamanho
    Do Until x_posicao = 0
        If Mid(pai_x, x_posicao, 1) = "\" Then
            x_pai = Mid(pai_x, (x_posicao + 1), (x_tamanho - x_posicao))
            Exit Do
        End If
        x_posicao = x_posicao - 1
    Loop
    If pai_x = "" Then
        pai_x = "\"
    End If
    pai_x = x_pai
    filho_x = x_filho
End Sub
Function preparaArredonda(ByVal pString As String, ByVal pCasasDecimais As Integer) As String
    Dim i As Integer
    Dim xString As String
    preparaArredonda = ""
    If bdAccess Then
        xString = Chr(39) & "0000000000."
        For i = 1 To pCasasDecimais
            xString = xString & "0"
        Next
        xString = xString & Chr(39)
        preparaArredonda = "FORMAT(" & pString & ", " & xString & ")"
    ElseIf bdSqlServer Then
        preparaArredonda = "ROUND(" & pString & ", " & pCasasDecimais & ")"
    End If
End Function
Function preparaBooleano(ByVal pBooleano As Boolean) As String
    preparaBooleano = ""
    If bdAccess Then
        If pBooleano = True Then
            preparaBooleano = "-1"
        Else
            preparaBooleano = "0"
        End If
    ElseIf bdSqlServer Then
        If pBooleano = True Then
            preparaBooleano = "1"
        Else
            preparaBooleano = "0"
        End If
    End If
End Function
Function preparaBooleanoBUG(ByVal pBooleano As Boolean) As String
    preparaBooleanoBUG = ""
    If pBooleano = True Then
        preparaBooleanoBUG = "-1"
    Else
        preparaBooleanoBUG = "0"
    End If
End Function
Function preparaData(ByVal xData As Date) As String
    preparaData = ""
    If bdAccess Then
        preparaData = Chr(35) & Format(xData, "mm/dd/yyyy") & Chr(35)
    ElseIf bdSqlServerAzure Then
        preparaData = Chr(39) & Format(xData, "yyyy/mm/dd") & Chr(39)
    ElseIf bdSqlServer Then
        preparaData = Chr(39) & Format(xData, "dd/mm/yyyy") & Chr(39)
    End If
End Function
Function preparaHora(ByVal xHora As Date) As String
    preparaHora = ""
    If bdAccess Then
        preparaHora = Chr(35) & Format(xHora, "hh:mm:ss") & Chr(35)
    ElseIf bdSqlServer Then
        preparaHora = Chr(39) & Format(xHora, "hh:mm:ss") & Chr(39)
    End If
End Function
Function preparaTexto(ByVal xTexto As String) As String
    preparaTexto = ""
    If bdAccess Then
        preparaTexto = Chr(39) & xTexto & Chr(39)
    ElseIf bdSqlServer Then
        preparaTexto = Chr(39) & xTexto & Chr(39)
    Else
        preparaTexto = Chr(39) & xTexto & Chr(39)
    End If
End Function
Function preparaValor(ByVal pValor As Currency) As String
    Dim xString As String
    preparaValor = ""
    If bdAccess Then
        xString = Format(pValor, "0000000000.0000")
        Mid(xString, 11, 1) = "."
    ElseIf bdSqlServer Then
        xString = Format(pValor, "0000000000.0000")
        Mid(xString, 11, 1) = "."
    End If
    preparaValor = xString
End Function
Function RetiraGString(ByVal pNumero As Integer) As String
    Dim xIndex As Integer
    Dim xInicio As Integer
    Dim xNumero As Integer
    
    RetiraGString = ""
    xInicio = 1
    xNumero = 1
    If Len(g_string) > 0 Then
        Do Until xIndex > Len(g_string)
            xIndex = xIndex + 1
            If Mid(g_string, xIndex, 3) = "|@|" Then
                If xNumero = pNumero Then
                    RetiraGString = Mid(g_string, xInicio, xIndex - xInicio)
                    Exit Function
                End If
                xIndex = xIndex + 2
                xNumero = xNumero + 1
                xInicio = xIndex + 1
            End If
        Loop
    End If
End Function
Function RetiraString(ByVal pNumero As Integer, ByVal pString As String) As String
    Dim xIndex As Integer
    Dim xInicio As Integer
    Dim xNumero As Integer
    
    RetiraString = ""
    xInicio = 1
    xNumero = 1
    If Len(pString) > 0 Then
        Do Until xIndex > Len(pString)
            xIndex = xIndex + 1
            If Mid(pString, xIndex, 3) = "|@|" Then
                If xNumero = pNumero Then
                    RetiraString = Mid(pString, xInicio, xIndex - xInicio)
                    Exit Function
                End If
                xIndex = xIndex + 2
                xNumero = xNumero + 1
                xInicio = xIndex + 1
            End If
        Loop
    End If
End Function
Function ContaMascarada(ByVal mConta As String) As Boolean
    ContaMascarada = False
    If Len(mConta) = 1 Then
        ContaMascarada = True
    ElseIf Len(mConta) > 1 Then
        If Mid(mConta, 2, 1) = "." Then
            ContaMascarada = True
        End If
    End If
End Function
Function CopiaArquivoDataHora(ByVal pNomeOriginal As String) As Boolean
    Dim xArqTxt As New FileSystemObject
    Dim xNomeArquivoCopia As String
    Dim xPosicaoInicial As Integer
    Dim i As Integer
    
    On Error GoTo FileError
    
    CopiaArquivoDataHora = False
    '12345678901234567890
    'SgpNetTemporario.ini
    '123456789012345678901234
    'SgpNetTemporarioNFCE.ini
   
    xPosicaoInicial = Len(pNomeOriginal) - 3
    xNomeArquivoCopia = Mid(pNomeOriginal, 1, Len(pNomeOriginal) - 4) & "_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Time, "HHmmss") & Mid(pNomeOriginal, xPosicaoInicial, 4)
    If xArqTxt.FileExists(xNomeArquivoCopia) = True Then
        For i = 1 To 30
            xNomeArquivoCopia = Mid(pNomeOriginal, 1, Len(pNomeOriginal) - 4) & "_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Time, "HHmmss") & "__" & i & Mid(pNomeOriginal, xPosicaoInicial, 4)
            If xArqTxt.FileExists(xNomeArquivoCopia) = False Then
                Exit For
            End If
        Next
    End If
    Call xArqTxt.CopyFile(pNomeOriginal, xNomeArquivoCopia, True)
    CopiaArquivoDataHora = True
    Exit Function

FileError:
    CriaLogSGP "CopiaArquivoDataHora: Erro ao verificar conexão com internet", Error, ""
End Function
Sub CriaCampo(ByVal x_nome_campo As String, ByVal x_tipo_campo As String, ByVal x_tamanho_campo As Integer, ByVal x_campo_em_branco As String)
    Set campo = New Field
    campo.name = Trim(x_nome_campo)
    If Trim(x_tipo_campo) = "Data/Time" Then
        campo.Type = dbDate
    ElseIf Trim(x_tipo_campo) = "Text" Then
        campo.Type = dbText
    ElseIf Trim(x_tipo_campo) = "Memo" Then
        campo.Type = dbMemo
    ElseIf Trim(x_tipo_campo) = "Boolean" Then
        campo.Type = dbBoolean
    ElseIf Trim(x_tipo_campo) = "Integer" Then
        campo.Type = dbInteger
    ElseIf Trim(x_tipo_campo) = "Long" Then
        campo.Type = dbLong
    ElseIf Trim(x_tipo_campo) = "Currency" Then
        campo.Type = dbCurrency
    ElseIf Trim(x_tipo_campo) = "Single" Then
        campo.Type = dbSingle
    ElseIf Trim(x_tipo_campo) = "Double" Then
        campo.Type = dbDouble
    ElseIf Trim(x_tipo_campo) = "Byte" Then
        campo.Type = dbByte
    ElseIf Trim(x_tipo_campo) = "Long Binary" Then
        campo.Type = dbLongBinary
    End If
    If x_tamanho_campo > 0 Then
        campo.Size = x_tamanho_campo
    End If
    If Trim(x_tipo_campo) = "Text" Or Trim(x_tipo_campo) = "Memo" Then
        If x_campo_em_branco = "S" Then
            campo.AllowZeroLength = True
        End If
    End If
    tabela.Fields.Append campo
End Sub
Sub CriaIndice(ByVal x_nome_indice As String, ByVal x_campo_do_indice As String, ByVal x_indice_primario As String, x_indice_unico As String)
    Set Indice = New Index
    Indice.name = Trim(x_nome_indice)
    Indice.Fields = Trim(x_campo_do_indice)
    If Trim(x_indice_primario) = "S" Then
        Indice.Primary = True
    Else
        Indice.Primary = False
    End If
    If Trim(x_indice_unico) = "S" Then
        Indice.Unique = True
    Else
        Indice.Unique = False
    End If
    tabela.Indexes.Append Indice
End Sub
Sub CriaLogConveniencia(ByVal pLinhaDeLog As String, ByVal pMensagemErro As String, ByVal pDetalhe As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    Dim xHora As String
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "SGP_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "C:\VB5\SGP\DATA\Conveniencia_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    xHora = Format(Time, "HH:MM:SS") & " - "
    
    'Grava o log
    xArquivo.WriteLine (xHora & pLinhaDeLog)
    If pMensagemErro <> "" Then
        xArquivo.WriteLine (xHora & "Mensagem de Erro:" & pMensagemErro)
    End If
    If pDetalhe <> "" Then
        xArquivo.WriteLine (xHora & "Detalhe:" & pDetalhe)
    End If
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub

FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG CriaLogConveniencia: " & pLinhaDeLog, vbInformation, "Erro: CriaLogConveniencia"
    Exit Sub
End Sub
Sub CriaLogCupom(ByVal pTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "ECF_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "C:\VB5\SGP\DATA\ECF_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub

FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG ECF: " & pTipoLog, vbInformation, "Erro: CriaLogCupom"
    Exit Sub
End Sub
Sub CriaLogMonitoraBug(ByVal pTipoLog As String, ByVal pPrograma As String, ByVal pMensagemLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "MonitoraBUG_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "C:\VB5\SGP\DATA\MonitoraBUG_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log: Tipo
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pTipoLog)
    'Grava o log: Progama
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pPrograma)
    'Grava o log: Mensagem
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pMensagemLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub

FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG MonitoraBUG: " & pTipoLog, vbInformation, "Erro: CriaLogMonitoraBug"
    Exit Sub
End Sub
Sub CriaLogSGP(ByVal pLinhaDeLog As String, ByVal pMensagemErro As String, ByVal pDetalhe As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    Dim xHora As String
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "SGP_DD_MM_YYYY.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "C:\VB5\SGP\DATA\SGP_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    xHora = Format(Time, "HH:MM:SS") & " - "
    
    'Grava o log
    xArquivo.WriteLine (xHora & pLinhaDeLog)
    If pMensagemErro <> "" Then
        xArquivo.WriteLine (xHora & "Mensagem de Erro:" & pMensagemErro)
    End If
    If pDetalhe <> "" Then
        xArquivo.WriteLine (xHora & "Detalhe:" & pDetalhe)
    End If
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG SGP: " & pLinhaDeLog, vbInformation, "Erro: CriaLogSGP"
    Exit Sub
End Sub
Sub CriaLogAutomacao(ByVal pTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    
    On Error GoTo FileError
    
    'Define nome do arquivo no seguinte formato: "AUTOMACAO_DD_MM_YYYY_HH.Log"
    'onde DD é o dia, MM o mês e YYYY o ano
    xNomeArquivo = "C:\VB5\SGP\DATA\AUTOMACAO_" & Format(Date, "dd") & "_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & "_" & Format(Time, "HH") & ".LOG"
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG ECF: " & pTipoLog, vbInformation, "Erro: CriaLogCupom"
    Exit Sub
End Sub
Sub CriaLogAutomacaoEncerrante(ByVal pMarcaAutomacao As String, ByVal pTipoLog As String)
    Dim xNomeArquivo As String
    Dim xArquivo As TextStream
    
    On Error GoTo FileError
    'Log Específico de encerrante lidos da automação
    'Define nome do arquivo no seguinte formato: "AUTOMACAO_Encerrante_MM_YYYY.Log"
    'onde MM o mês e YYYY o ano
    If pMarcaAutomacao = "COMPANY" Then
        xNomeArquivo = "C:\Cerrado\AutoCerradoCompany\AUTOMACAO_Encerrante_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    ElseIf pMarcaAutomacao = "HOROUSTECH" Then
        xNomeArquivo = "C:\Cerrado.Net\AutoCerradoHorousTech\AUTOMACAO_Encerrante_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    ElseIf pMarcaAutomacao = "EZTECH" Then
        xNomeArquivo = "C:\Cerrado\AutoCerradoEZ\AUTOMACAO_Encerrante_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    ElseIf pMarcaAutomacao = "IONICS" Then
        xNomeArquivo = "C:\Cerrado\AutoCerradoIonics\AUTOMACAO_Encerrante_" & Format(Date, "mm") & "_" & Format(Date, "yyyy") & ".LOG"
    End If
    
    'Verifica se o arquivo existe, depois abre ou cria
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set xArquivo = gArqTxt.OpenTextFile(xNomeArquivo, ForAppending)
    Else
        Set xArquivo = gArqTxt.CreateTextFile(xNomeArquivo)
    End If
    
    'Grava o log
    xArquivo.WriteLine (Format(Time, "HH:MM:SS") & " - " & pTipoLog)
    
    'Fecha arquivo texto
    xArquivo.Close
    Set xArquivo = Nothing
    Exit Sub
FileError:
    'MsgBox Error
    'MsgBox "Erro ao criar LOG ECF: " & pTipoLog, vbInformation, "Erro: CriaLogAutomacaoEncerrante"
    Exit Sub
End Sub
Function CriaNovoMDB(ByVal nNomeArquivo As String, ByVal nFormato As Integer) As Boolean
Dim catDB As ADOX.Catalog

On Error GoTo trata_erro

    'abre o objeto catalogo no bd no qual vamos criar o indice
    Set catDB = New ADOX.Catalog
    catDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=" & nFormato & ";Data Source=" & nNomeArquivo
    CriaNovoMDB = True
    Exit Function
    
trata_erro:

    CriaNovoMDB = False
    
    ' o tratamento de erro você pode colocar na rotina que chama a função
    If Err.Number = -2147217897 Then
        MsgBox "O banco de dados: " & nNomeArquivo & " já existe.", vbInformation, "Banco não pode ser Criado!"
    Else
        MsgBox "Ocorreu um erro durante a criação do banco de dados " & nNomeArquivo & vbCrLf & _
        Err.Description & vbCrLf & _
        "Erro numero : " & Err.Number
    End If

End Function
Sub DisabelCtrlAltDel(bdisabled As Boolean)
  Dim X As Long
  X = SystemParametersInfo(97, bdisabled, CStr(1), 0)
End Sub

'Percorre impressoras instaladas no spooler do windows e define a primeira térmica encontrada com padrão
Public Function DefineImpressoraTermicaComoPadrao() As Boolean

DefineImpressoraTermicaComoPadrao = False
Dim Impressora As Printer
Dim Contador As Byte
Contador = 0
For Each Impressora In Printers
    If UCase(Impressora.DeviceName) Like "*TM-T20*" Or UCase(Impressora.DeviceName) Like "*TM-T8*" Or UCase(Impressora.DeviceName) Like "*MP-4*" Or UCase(Impressora.DeviceName) Like "*MP-2*" Or UCase(Impressora.DeviceName) Like "*MP-1*" Then
        gIDImpressoraPadrao = Contador
        Set Printer = Printers(gIDImpressoraPadrao)
        DefineImpressoraTermicaComoPadrao = True
        Exit Function
    End If
    Contador = Contador + 1
Next
MsgBox "Nenhuma Impressora térmica foi encontrada!"
End Function

