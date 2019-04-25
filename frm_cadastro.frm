VERSION 5.00
Begin VB.Form frm_cadastro 
   Caption         =   "Chamando Cadastro..."
   ClientHeight    =   240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lNomePrograma As String
Private Sub ChamaPrograma()
    If lNomePrograma = "cadastro_aliquota" Then
        cadastro_aliquota.Show
    ElseIf lNomePrograma = "cadastro_banco" Then
        cadastro_banco.Show
    ElseIf lNomePrograma = "cadastro_bomba" Then
        cadastro_bomba.Show
    ElseIf lNomePrograma = "cadastro_cartao" Then
        cadastro_cartao.Show
    ElseIf lNomePrograma = "cadastro_cliente" Then
        cadastro_cliente.Show
    ElseIf lNomePrograma = "cadastro_cliente_conveniado" Then
        cadastro_cliente_conveniado.Show
    ElseIf lNomePrograma = "cadastro_combustivel" Then
        cadastro_combustivel.Show
    ElseIf lNomePrograma = "cadastro_composicao_caixa" Then
        cadastro_composicao_caixa.Show
    ElseIf lNomePrograma = "cadastro_configuracao" Then
        cadastro_configuracao.Show
    ElseIf lNomePrograma = "cadastro_conta_bancaria" Then
        cadastro_conta_bancaria.Show
    ElseIf lNomePrograma = "cadastro_convenio" Then
        cadastro_convenio.Show
    ElseIf lNomePrograma = "cadastro_conversao_medicao" Then
        cadastro_conversao_medicao.Show
    ElseIf lNomePrograma = "cadastro_dependente" Then
        cadastro_dependente.Show
    ElseIf lNomePrograma = "cadastro_empresa" Then
        cadastro_empresa.Show
    ElseIf lNomePrograma = "cadastro_fornecedor" Then
        cadastro_fornecedor.Show
    ElseIf lNomePrograma = "cadastro_funcionario" Then
        cadastro_funcionario.Show
    ElseIf lNomePrograma = "cadastro_grau_dependencia" Then
        cadastro_grau_dependencia.Show
    ElseIf lNomePrograma = "cadastro_grupo" Then
        cadastro_grupo.Show
    ElseIf lNomePrograma = "cadastro_historico" Then
        cadastro_historico.Show
    ElseIf lNomePrograma = "cadastro_local_cobranca" Then
        cadastro_local_cobranca.Show
    ElseIf lNomePrograma = "cadastro_menu" Then
        cadastro_menu.Show
    ElseIf lNomePrograma = "cadastro_plano_conta" Then
        cadastro_plano_conta.Show
    ElseIf lNomePrograma = "cadastro_produto" Then
        cadastro_produto.Show
    ElseIf lNomePrograma = "cadastro_programa" Then
        cadastro_programa.Show
    ElseIf lNomePrograma = "cadastro_situacao_cheque_devolvido" Then
        cadastro_situacao_cheque_devolvido.Show
    ElseIf lNomePrograma = "cadastro_sub_grupo" Then
        cadastro_sub_grupo.Show
    ElseIf lNomePrograma = "cadastro_tabela_folha" Then
        cadastro_tabela_folha.Show
    ElseIf lNomePrograma = "cadastro_tabela_premiacao" Then
        cadastro_tabela_premiacao.Show
    ElseIf lNomePrograma = "cadastro_tabela_provento_desconto" Then
        cadastro_tabela_provento_desconto.Show
    ElseIf lNomePrograma = "cadastro_tabela_vencimento" Then
        cadastro_tabela_vencimento.Show
    ElseIf lNomePrograma = "cadastro_tanque_combustivel" Then
        cadastro_tanque_combustivel.Show
    ElseIf lNomePrograma = "cadastro_tipo_documento" Then
        cadastro_tipo_documento.Show
    ElseIf lNomePrograma = "cadastro_usuario" Then
        cadastro_usuario.Show
    Else
        Unload Me
    End If
End Sub
Function ExisteSgpCadastroIni() As Boolean
    Dim xNomeArquivo As String
    ExisteSgpCadastroIni = False
    'xNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    xNomeArquivo = "C:" & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    If gArqTxt.FileExists(xNomeArquivo) Then
        ExisteSgpCadastroIni = True
    End If
End Function
Private Sub Finaliza()
    cnnSGP.Close
    End
End Sub
Private Sub LeSgpCadastroIni()
    Dim xNomeArquivo As String
    Dim xString As String
    'xNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    xNomeArquivo = "C:" & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    If gArqTxt.FileExists(xNomeArquivo) Then
        Set gArquivoTMP = gArqTxt.OpenTextFile(xNomeArquivo, ForReading)
        Do Until gArquivoTMP.AtEndOfStream
            xString = gArquivoTMP.ReadLine
            
            '[Empresa]
            If Mid(xString, 1, 8) = "Empresa=" Then
                g_empresa = Mid(xString, 9, Len(xString) - 8)
            ElseIf Mid(xString, 1, 12) = "NomeEmpresa=" Then
                g_nome_empresa = Mid(xString, 13, Len(xString) - 12)
            ElseIf Mid(xString, 1, 14) = "CidadeEmpresa=" Then
                g_cidade_empresa = Mid(xString, 15, Len(xString) - 14)
            
            '[Programa]
            ElseIf Mid(xString, 1, 9) = "Programa=" Then
                lNomePrograma = Mid(xString, 10, Len(xString) - 9)
            
            '[Usuario]
            ElseIf Mid(xString, 1, 8) = "Usuario=" Then
                g_usuario = Mid(xString, 9, Len(xString) - 8)
            ElseIf Mid(xString, 1, 12) = "NomeUsuario=" Then
                g_nome_usuario = Mid(xString, 13, Len(xString) - 12)
            ElseIf Mid(xString, 1, 19) = "NivelAcessoUsuario=" Then
                g_nivel_acesso = Mid(xString, 20, Len(xString) - 19)
            
            '[Outras]
            ElseIf Mid(xString, 1, 8) = "DataDef=" Then
                g_data_def = Mid(xString, 9, Len(xString) - 8)
            ElseIf Mid(xString, 1, 8) = "FlagLmc=" Then
                g_lmc = Mid(xString, 9, Len(xString) - 8)
            ElseIf Mid(xString, 1, 20) = "ImpressoraMatricial=" Then
                g_impressora_matricial = Mid(xString, 21, Len(xString) - 20)
            ElseIf Mid(xString, 1, 15) = "CaixaUnificado=" Then
                g_caixa_unificado = Mid(xString, 16, Len(xString) - 15)
                
            '[ContaBancaria]
            ElseIf Mid(xString, 1, 14) = "ContaBancaria=" Then
                g_conta_bancaria = Mid(xString, 15, Len(xString) - 14)
            ElseIf Mid(xString, 1, 18) = "NomeContaBancaria=" Then
                g_nome_conta = Mid(xString, 19, Len(xString) - 18)
                
            '[Liberacao]
            ElseIf Mid(xString, 1, 15) = "EmpresaInicial=" Then
                g_cfg_empresa_i = Mid(xString, 16, Len(xString) - 15)
            ElseIf Mid(xString, 1, 13) = "EmpresaFinal=" Then
                g_cfg_empresa_f = Mid(xString, 14, Len(xString) - 13)
            ElseIf Mid(xString, 1, 12) = "DataInicial=" Then
                g_cfg_data_i = Mid(xString, 13, Len(xString) - 12)
            ElseIf Mid(xString, 1, 10) = "DataFinal=" Then
                g_cfg_data_f = Mid(xString, 11, Len(xString) - 10)
            ElseIf Mid(xString, 1, 15) = "PeriodoInicial=" Then
                g_cfg_periodo_i = Mid(xString, 16, Len(xString) - 15)
            ElseIf Mid(xString, 1, 13) = "PeriodoFinal=" Then
                g_cfg_periodo_f = Mid(xString, 14, Len(xString) - 13)
                
            '[Automacao]
            ElseIf Mid(xString, 1, 10) = "Automacao=" Then
                g_automacao = Mid(xString, 11, Len(xString) - 10)
            End If
        Loop
        gArquivoTMP.Close
        gArqTxt.DeleteFile (xNomeArquivo)
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        lFlagCadastro = 1
        LeSgpCadastroIni
        lNomePrograma = "cadastro_bomba"
        ChamaPrograma
        Me.Hide
    Else
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim xNomeBancoDados As String
    Call ChamaDrive
    '''If Not ExisteSgpCadastroIni Then
    '''    End
    '''End If
    Screen.MousePointer = 11
    bdAccess = False
    bdMySql = False
    bdSqlServer = False
    bdOracle = False
    xNomeBancoDados = ReadINI("SGBD", "Gerenciador de Banco de Dados", ArqSgpIni)
    If xNomeBancoDados = "ACCESS" Then
        bdAccess = True
    ElseIf xNomeBancoDados = "MYSQL" Then
        bdMySql = True
    ElseIf xNomeBancoDados = "SQLSERVER" Then
        bdSqlServer = True
    ElseIf xNomeBancoDados = "ORACLE" Then
        bdOracle = True
    End If
    
    Set Conectar = New CConexao
    If bdAccess Then
        gConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & gDrive & gDiretorioData & gNomeBancoDados
        cnnSGP.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gDrive & gDiretorioData & gNomeBancoDados & ";Uid=Admin;Pwd=;"
    ElseIf bdSqlServer Then
        gConnectionString = Conectar.ConnectionString
        cnnSGP.Open gConnectionString
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
