VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form menu_personalizado 
   Caption         =   "Sistema Gerênciador de Postos de Combustíveis"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11355
   Icon            =   "Menupers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Menupers.frx":0442
   ScaleHeight     =   1500
   ScaleWidth      =   11355
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConexaoGic 
      Caption         =   "G&IC"
      Height          =   915
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ativa verificação de comunicação com o GIC."
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdResgate 
      Caption         =   "Resgate &NFe"
      Height          =   915
      Left            =   8460
      Picture         =   "Menupers.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Manifesto e Importação de NF-e."
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdTransfereDadosLMC 
      Caption         =   "&Transfere p/ LMC"
      Height          =   735
      Left            =   9840
      Picture         =   "Menupers.frx":14CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Transfere as entradas de combustíveis para o LMC."
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1125
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer timerVbNet 
      Enabled         =   0   'False
      Left            =   480
      Top             =   540
   End
   Begin VB.CommandButton cmd_sql 
      Height          =   555
      Left            =   6480
      Picture         =   "Menupers.frx":28BC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exporta / Importa dados."
      Top             =   360
      Width           =   795
   End
   Begin VB.CommandButton cmd_configuracao 
      Caption         =   "&Liberação"
      Height          =   915
      Left            =   7440
      Picture         =   "Menupers.frx":3B96
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Libera data de digitação."
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmd_senha 
      Caption         =   "S&enha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Muda senha."
      Top             =   0
      Width           =   795
   End
   Begin MSAdodcLib.Adodc adodc_empresa 
      Height          =   330
      Left            =   3240
      Top             =   120
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adodc_empresa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dtcbo_empresa 
      Bindings        =   "Menupers.frx":4E70
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Nome"
      BoundColumn     =   "Codigo"
      Text            =   "dtcbo_empresa"
   End
   Begin VB.Label Label7 
      Caption         =   "Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Menu mnu_cadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnu_cadastro_item 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_consulta 
      Caption         =   "Co&nsultas"
      Begin VB.Menu mnu_consulta_item 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_grafico 
      Caption         =   "&Gráficos"
      Begin VB.Menu mnu_grafico_item 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_movimentacao 
      Caption         =   "&Movimentação"
      Begin VB.Menu mnu_movimentacao_item 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_relatorio 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnu_relatorio_item 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_sobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "menu_personalizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sair As Integer
Dim i_ca As Integer
Dim i_co As Integer
Dim i_gr As Integer
Dim i_mo As Integer
Dim i_re As Integer
Dim lSQL As String
Dim lNomeArquivo As String
Dim lArquivoVb6VbNet As String
Dim lArquivoVb6VbNet2 As String
'Dim lContadorTimer As Integer
'Dim lContadorTimer2 As Integer
Dim lNotificacaoGic As Boolean
Dim lResgateAbertoPeloSGP As Boolean
Dim lForm As Form

Dim lEmpresaAtualPetromovel As Integer


Private RstMenu As ADODB.Recordset

Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private dados As New cDados
Private Empresa As New cEmpresa
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private Programa As New cPrograma



Private Sub BuscaConfiguracao()
    g_caixa_unificado = False
    BuscaConfiguracaoCaixaUnificado
    gInternetBandaLarga = False
    lNotificacaoGic = False
    DesativaVerificacaoGIC
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Internet Banda Larga") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            gInternetBandaLarga = True
            'cmdConexaoGic.Visible = True - ALEX - 01/09/2017
            If ConfiguracaoDiversa.LocalizarCodigo(1, "GIC: Notificacao Periodica") Then
                If ConfiguracaoDiversa.Verdadeiro Then
                    lNotificacaoGic = True
                    AtivaVerificacaoGIC
                End If
            End If
        End If
    End If
End Sub
Private Sub BuscaConfiguracaoCaixaUnificado()
    g_caixa_unificado = False
    gQtdPeriodo = 2
    If Configuracao.LocalizarCodigo(g_empresa) Then
        gQtdPeriodo = Configuracao.QuantidadePeriodos
        If Mid(Configuracao.OutrasConfiguracoes, 1, 1) = "S" Then
            g_caixa_unificado = True
        Else
            g_caixa_unificado = False
        End If
        If Mid(Configuracao.OutrasConfiguracoes, 5, 1) = "S" Then
            g_automacao = True
        Else
            g_automacao = False
        End If
    Else
        Call CriaConfiguracao
    End If
End Sub
Private Sub BuscaDadosGIC()
    Dim xStringConexao As String
    Dim xIpBanco As String
    Dim rsTabela As New ADODB.Recordset
    Dim xSQL As String
    Dim xCodigoFuncionario As Integer
    Dim xNomeUsuario As String
    Dim xCodigoGrupoEmpresa As Integer
    Dim i As Integer
    
    On Error GoTo FileError
        
    If ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni) = "TEIXEIRA E PINHEIRO LTDA" Then
        'xIpBanco = "192.168.1.6,4949"
        xIpBanco = "192.168.1.6" & gPortaBanco
    Else
        'xIpBanco = "tasso.myvnc.com,4949"
        xIpBanco = "tasso.myvnc.com" & gPortaBanco
    End If
    xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xIpBanco & ";INITIAL CATALOG=" & "CerradoData" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
    Set gConnGic = New ADODB.Connection
    gConnGic.ConnectionString = xStringConexao
    gConnGic.Open
    
    'Busca Funcionario
    xCodigoFuncionario = 0
    gUsuarioGlobal = 0
    xCodigoGrupoEmpresa = 0

    i = InStr(1, g_nome_usuario, " ", vbTextCompare)
    If i > 0 Then
        xNomeUsuario = UCase(Mid(g_nome_usuario, 1, i - 1))
    Else
        xNomeUsuario = UCase(g_nome_usuario)
    End If

    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, Nome"
    xSQL = xSQL & "  FROM Funcionario"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Codigo do Usuario] = " & g_usuario
    xSQL = xSQL & "   AND Situacao = " & preparaTexto("A")
    Set rsTabela = New ADODB.Recordset
    rsTabela.Open xSQL, Conectar.Conexao, adOpenForwardOnly, adLockReadOnly
    If Not rsTabela.EOF Then
        
        Do Until rsTabela.EOF
            If UCase(rsTabela("Nome").Value) Like "*" & xNomeUsuario & "*" Then
                xCodigoFuncionario = rsTabela("Codigo").Value
                Exit Do
            End If
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    
    If xCodigoFuncionario = 0 Then
        gEmpresaGlobal = 7
    End If
    
    If xCodigoFuncionario > 0 Or gEmpresaGlobal = 7 Then
        'Busca Codigo do Grupo da Empresa
        xSQL = ""
        xSQL = xSQL & "SELECT [Empresa do Grupo]"
        xSQL = xSQL & "  FROM Empresa"
        xSQL = xSQL & " WHERE [Codigo Global] = " & gEmpresaGlobal
        Set rsTabela = New ADODB.Recordset
        rsTabela.Open xSQL, gConnGic, adOpenForwardOnly, adLockReadOnly
        If Not rsTabela.EOF Then
            xCodigoGrupoEmpresa = rsTabela("Empresa do Grupo").Value
        End If
        rsTabela.Close
        Set rsTabela = Nothing
        
        'Busca Codigo Usuario Global
        xSQL = ""
        xSQL = xSQL & "SELECT [Codigo Global], Nome"
        xSQL = xSQL & "  FROM UsuarioComunicacao"
        xSQL = xSQL & " WHERE Nome Like " & preparaTexto("%" & xNomeUsuario & "%")
        xSQL = xSQL & "   AND [Grupo de Empresa] = " & xCodigoGrupoEmpresa
        xSQL = xSQL & "   AND Ativo = " & preparaBooleano(True)
        If gEmpresaGlobal = 7 Then
            xSQL = xSQL & "   AND [Empresa Global] = " & 7
        End If
        Set rsTabela = New ADODB.Recordset
        rsTabela.Open xSQL, gConnGic, adOpenForwardOnly, adLockReadOnly
        If Not rsTabela.EOF Then
            gUsuarioGlobal = rsTabela("Codigo Global").Value
        End If
        rsTabela.Close
        Set rsTabela = Nothing
        
        'Busca Comunicacao
        xSQL = ""
        xSQL = xSQL & "SELECT Count(1) AS Quantidade"
        xSQL = xSQL & "  FROM MovimentoComunicacao"
        xSQL = xSQL & " WHERE [Empresa Global Destinatario] = " & gEmpresaGlobal
        xSQL = xSQL & "   AND [Codigo do Destinatario Global] = " & gUsuarioGlobal
        xSQL = xSQL & "   AND Concluida = " & preparaBooleano(False)
        xSQL = xSQL & "   AND Cancelada = " & preparaBooleano(False)
        Set rsTabela = New ADODB.Recordset
        rsTabela.Open xSQL, gConnGic, adOpenForwardOnly, adLockReadOnly
        If Not rsTabela.EOF Then
            If rsTabela("Quantidade").Value = 1 Then
'                If (MsgBox("Voce tem " & rsTabela("Quantidade").Value & " comunicacao pendente." & vbCrLf & "Deseja ler agora?", vbYesNo + vbQuestion + vbDefaultButton1, "Comunicação Pendente!")) = vbYes Then
                gStringChamada = "1 Comunicação Pendente!"
                Call GravaSgpCadastroIni("Avisa MovimentoComunicacaoGIC")
'                End If
                gStringChamada = ""
            ElseIf rsTabela("Quantidade").Value > 1 Then
'                If (MsgBox("Voce tem " & rsTabela("Quantidade").Value & " comunicações pendentes." & vbCrLf & "Deseja ler agora?", vbYesNo + vbQuestion + vbDefaultButton1, "Comunicação Pendente!")) = vbYes Then
                gStringChamada = rsTabela("Quantidade").Value & " Comunicações Pendentes."
                Call GravaSgpCadastroIni("Avisa MovimentoComunicacaoGIC")
'                End If
                gStringChamada = ""
            End If
    '            Do Until rsTabela.EOF
    '                MsgBox "Assunto: " & rsTabela("Assunto").Value & vbCrLf & "Texto: " & rsTabela("Texto").Value
    '                rsTabela.MoveNext
    '            Loop
        End If
        rsTabela.Close
        Set rsTabela = Nothing
        gConnGic.Close
        Set gConnGic = Nothing
        End If
    Exit Sub

FileError:
    MsgBox "Erro ao abrir acessar dados do GIC!" & Chr(10) & Error, vbCritical, "Erro de Conexão.!"
End Sub
Private Sub BuscaDados()
    If dados.LocalizarCodigo(1) Then
        'If Dados.Empresa2 = 0 Then
        '    If CDate(Date) >= CDate("10/04/2000") Then
        '        Dados.Empresa2 = 1
        '        If Not Dados.Alterar(1) Then
        '            MsgBox "Registro não alterado.", vbInformation, "Erro de Integridade"
        '        End If
        '    End If
        'End If
        'If Dados.Empresa2 <> 0 Then
        '    Printer.PaperSize = 300
        'End If
        g_empresa = dados.Empresa
        g_conta_bancaria = dados.ContaBancaria
        If Empresa.LocalizarCodigo(g_empresa) Then
            g_nome_empresa = Empresa.Nome
            gCNPJEmpresa = Empresa.CGC 'ALEX - NFCE
        End If
    Else
        dados.Codigo = 1
        dados.Empresa = 1
        dados.ContaBancaria = 1
        dados.Empresa2 = 0
        g_empresa = dados.Empresa
        If Not dados.Incluir Then
            MsgBox "Registro não incluído.", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub Chama(ByVal pNomeProgramaInterno As String)
    pNomeProgramaInterno = Trim(pNomeProgramaInterno)
    If g_empresa = 0 Then
        MsgBox "Selecione uma empresa!", vbInformation, "Empresa Não Selecionada!"
        dtcbo_empresa.SetFocus
        Exit Sub
    End If
'Cadastro

    'Verifica se tem formulário específico já aberto
    Dim frm As Form

    If pNomeProgramaInterno = "Movimento_Nfce_Auto" Or pNomeProgramaInterno = "Movimento_Nfce_Conveniencia" Then
        Call GravaSgpNetCadastroIni2("VerificaTipoAmbiente")
    End If
    For Each frm In Forms
        If pNomeProgramaInterno = "emissao_cupom_complementar" Then
            If frm.name = "movimento_cupom_fiscal" Or frm.name = "movimento_cupom_fiscal_auto" Then
                Call GravaAuditoria(1, Me.name, 22, "Cupom Complementar não será aberto")
                MsgBox "Para abrir o Cupom Complementar," & vbCrLf & "Será necessário fechar o programa de Cupom Fiscal.", vbOKOnly + vbInformation, "Operação Não Permitida!"
                Exit Sub
            End If
        End If
        If pNomeProgramaInterno = "movimento_cupom_fiscal" Or pNomeProgramaInterno = "movimento_cupom_fiscal_auto" Or pNomeProgramaInterno = "Movimento_Nfce_Auto" Then
            If frm.name = "movimento_cupom_fiscal" Or frm.name = "movimento_cupom_fiscal_auto" Or frm.name = "Movimento_Nfce_Auto" Then
                Call GravaAuditoria(1, Me.name, 22, "Já existe uma tela de NFCe ou de Cupom Fiscal aberta.")
                MsgBox "Já existe uma tela de NFCe ou de Cupom Fiscal aberta." & vbCrLf & "Por isso não será possível abrir outra.", vbOKOnly + vbInformation, "Operação Não Permitida!"
                If frm.WindowState = 1 Then
                    frm.WindowState = 0
                End If
                Exit Sub
            End If
        End If
    Next frm



    Screen.MousePointer = 11

    Set lForm = Forms.Add(pNomeProgramaInterno)
    lForm.Show
    
    Exit Sub
    
    
    'If pNomeProgramaInterno = "cadastro_aliquota" Then
    '    cadastro_aliquota.Show
    'ElseIf pNomeProgramaInterno = "cadastro_banco" Then
    '    cadastro_banco.Show
    'ElseIf pNomeProgramaInterno = "cadastro_bomba" Then
    '    cadastro_bomba.Show
    'ElseIf pNomeProgramaInterno = "cadastro_cartao" Then
    '    cadastro_cartao.Show
    'ElseIf pNomeProgramaInterno = "cadastro_cliente" Then
    '    cadastro_cliente.Show
    'ElseIf pNomeProgramaInterno = "cadastro_cliente_conveniado" Then
    '    cadastro_cliente_conveniado.Show
    'ElseIf pNomeProgramaInterno = "cadastro_combustivel" Then
    '    cadastro_combustivel.Show
    'ElseIf pNomeProgramaInterno = "cadastro_composicao_caixa" Then
    '    cadastro_composicao_caixa.Show
    'ElseIf pNomeProgramaInterno = "cadastro_configuracao" Then
    '    cadastro_configuracao.Show
    'ElseIf pNomeProgramaInterno = "cadastro_conta_bancaria" Then
    '    cadastro_conta_bancaria.Show
    'ElseIf pNomeProgramaInterno = "cadastro_convenio" Then
    '    cadastro_convenio.Show
    'ElseIf pNomeProgramaInterno = "cadastro_conversao_medicao" Then
    '    cadastro_conversao_medicao.Show
    'ElseIf pNomeProgramaInterno = "cadastro_dependente" Then
    '    cadastro_dependente.Show
    'ElseIf pNomeProgramaInterno = "cadastro_empresa" Then
    '    cadastro_empresa.Show
    'ElseIf pNomeProgramaInterno = "cadastro_fornecedor" Then
    '    cadastro_fornecedor.Show
    'ElseIf pNomeProgramaInterno = "cadastro_funcionario" Then
    '    cadastro_funcionario.Show
    'ElseIf pNomeProgramaInterno = "cadastro_grau_dependencia" Then
    '    cadastro_grau_dependencia.Show
    'ElseIf pNomeProgramaInterno = "cadastro_grupo" Then
    '    cadastro_grupo.Show
    'ElseIf pNomeProgramaInterno = "cadastro_historico" Then
    '    cadastro_historico.Show
    'ElseIf pNomeProgramaInterno = "cadastro_mala_direta" Then
    '    cadastro_mala_direta.Show
    'ElseIf pNomeProgramaInterno = "cadastro_menu" Then
    '    cadastro_menu.Show
    'ElseIf pNomeProgramaInterno = "cadastro_produto" Then
    '    cadastro_produto.Show
    'ElseIf pNomeProgramaInterno = "cadastro_programa" Then
    '    cadastro_programa.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tabela_folha" Then
    '    cadastro_tabela_folha.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tabela_premiacao" Then
    '    cadastro_tabela_premiacao.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tabela_provento_desconto" Then
    '    cadastro_tabela_provento_desconto.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tabela_vencimento" Then
    '    cadastro_tabela_vencimento.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tanque_combustivel" Then
    '    cadastro_tanque_combustivel.Show
    'ElseIf pNomeProgramaInterno = "cadastro_tipo_documento" Then
    '    cadastro_tipo_documento.Show
    'ElseIf pNomeProgramaInterno = "cadastro_usuario" Then
    '    cadastro_usuario.Show

''Cadastro Conversao
''Consulta
'    If pNomeProgramaInterno = "cerrado_calendario" Then
'        cerrado_calendario.Show
'    ElseIf pNomeProgramaInterno = "consulta_lmc" Then
'        consulta_lmc.Show
'    ElseIf pNomeProgramaInterno = "consulta_nota_cliente" Then
'        consulta_nota_cliente.Show
'    ElseIf pNomeProgramaInterno = "consulta_nota_conveniado" Then
'        consulta_nota_conveniado.Show
'    ElseIf pNomeProgramaInterno = "consulta_quadro_funcionario" Then
'        consulta_quadro_funcionario.Show
'    ElseIf pNomeProgramaInterno = "consulta_movimento_cupom_fiscal" Then
'        consulta_movimento_cupom_fiscal.Show
'    ElseIf pNomeProgramaInterno = "envia_email" Then
'        envia_email.Show
'    ElseIf pNomeProgramaInterno = "super_consulta" Then
'        super_consulta.Show
'    ElseIf pNomeProgramaInterno = "visualizador_log" Then
'        visualizador_log.Show
'    ElseIf pNomeProgramaInterno = "ConsultaEstoque" Then
'        ConsultaEstoque.Show
'    ElseIf pNomeProgramaInterno = "con_cheque_predatados" Then
'        consulta_cheque.Show
'
''Gráficos
'    ElseIf pNomeProgramaInterno = "grafico_despesa_anual" Then
'        grafico_despesa_anual.Show
'    ElseIf pNomeProgramaInterno = "grafico_despesa_mensal" Then
'        grafico_despesa_mensal.Show
'    ElseIf pNomeProgramaInterno = "grafico_venda_combustivel_anual" Then
'        grafico_venda_combustivel_anual.Show
'    ElseIf pNomeProgramaInterno = "grafico_venda_combustivel_mensal" Then
'        grafico_venda_combustivel_mensal.Show
'
''Movimento Baixa
'    ElseIf pNomeProgramaInterno = "baixa_cheque" Then
'        baixa_cheque.Show
'    ElseIf pNomeProgramaInterno = "baixa_cheque_individual" Then
'        baixa_cheque_individual.Show
'    ElseIf pNomeProgramaInterno = "baixa_cartao" Then
'        baixa_cartao.Show
'    ElseIf pNomeProgramaInterno = "baixa_cheque_devolvido_descontado" Then
'        baixa_cheque_devolvido_descontado.Show
''    ElseIf pNomeProgramaInterno = "baixa_contas_pagar" Then
''        baixa_contas_pagar.Show
'    ElseIf pNomeProgramaInterno = "baixa_duplicata_receber" Then
'        baixa_duplicata_receber.Show
'    ElseIf pNomeProgramaInterno = "baixa_nota_abastecimento" Then
'        baixa_nota_abastecimento.Show
'    ElseIf pNomeProgramaInterno = "baixa_nota_abastecimento_periodo" Then
'        baixa_nota_abastecimento_periodo.Show
''Movimento
'    ElseIf pNomeProgramaInterno = "GeraArquivoSintegra" Then
'        GeraArquivoSintegra.Show
'    ElseIf pNomeProgramaInterno = "gera_disquete_deposito" Then
'        gera_disquete_deposito.Show
'    ElseIf pNomeProgramaInterno = "importa_exporta_mapa_resumo" Then
'        importa_exporta_mapa_resumo.Show
'    ElseIf pNomeProgramaInterno = "movimento_bomba" Then
'        movimento_bomba.Show
'    ElseIf pNomeProgramaInterno = "movimento_cheque" Then
'        movimento_cheque.Show
'    ElseIf pNomeProgramaInterno = "movimento_cheque_avista" Then
'        movimento_cheque_avista.Show
'    ElseIf pNomeProgramaInterno = "movimento_cheque_devolvido" Then
'        movimento_cheque_devolvido.Show
'    ElseIf pNomeProgramaInterno = "movimento_cheque_devolvido_baixado" Then
'        movimento_cheque_devolvido_baixado.Show
''    ElseIf pNomeProgramaInterno = "movimento_cheque_extraviado" Then
''        movimento_cheque_extraviado.Show
'    ElseIf pNomeProgramaInterno = "movimento_desconto_personalizado" Then
'        movimento_desconto_personalizado.Show
'    ElseIf pNomeProgramaInterno = "movimento_entrada_produto" Then
'        movimento_entrada_produto.Show
'    ElseIf pNomeProgramaInterno = "movimento_falta_caixa" Then
'        movimento_falta_caixa.Show
'    ElseIf pNomeProgramaInterno = "mov_contas_pagar" Then
'        mov_contas_pagar.Show
'    ElseIf pNomeProgramaInterno = "movimento_duplicata_receber" Then
'        movimento_duplicata_receber.Show
'    ElseIf pNomeProgramaInterno = "mov_entrada_combustiveis" Then
'        mov_entrada_combustiveis.Show
'    ElseIf pNomeProgramaInterno = "movimento_medicao_combustivel" Then
'        movimento_medicao_combustivel.Show
'    ElseIf pNomeProgramaInterno = "MovimentoMedicaoCombustivelRegua" Then
'        MovimentoMedicaoCombustivelRegua.Show
'    'ElseIf pNomeProgramaInterno = "mov_nota_abastecimento" Then
'    '    mov_nota_abastecimento.Show
'    ElseIf pNomeProgramaInterno = "movimento_advertencia_suspencao" Then
'        movimento_advertencia_suspencao.Show
'    ElseIf pNomeProgramaInterno = "movimento_afericao" Then
'        movimento_afericao.Show
'    ElseIf pNomeProgramaInterno = "movimento_caixa" Then
'        movimento_caixa.Show
'    ElseIf pNomeProgramaInterno = "movimento_cartao_credito" Then
'        movimento_cartao_credito.Show
'    ElseIf pNomeProgramaInterno = "movimento_cheque_cobranca" Then
'        movimento_cheque_cobranca.Show
'    ElseIf pNomeProgramaInterno = "movimento_composicao_caixa" Then
'        movimento_composicao_caixa.Show
'    ElseIf pNomeProgramaInterno = "movimento_cupom_fiscal" Then
'        movimento_cupom_fiscal.Show
'    ElseIf pNomeProgramaInterno = "movimento_cupom_fiscal_auto" Then
'        movimento_cupom_fiscal_auto.Show
'    ElseIf pNomeProgramaInterno = "movimento_falta_funcionario" Then
'        movimento_falta_funcionario.Show
'    ElseIf pNomeProgramaInterno = "movimento_folha" Then
'        movimento_folha.Show
'    ElseIf pNomeProgramaInterno = "movimento_historico" Then
'        movimento_historico.Show
'    ElseIf pNomeProgramaInterno = "movimento_mapa_resumo" Then
'        movimento_mapa_resumo.Show
''    ElseIf pNomeProgramaInterno = "movimento_leasing_veiculo" Then
''        movimento_leasing_veiculo.Show
'    ElseIf pNomeProgramaInterno = "movimento_oleo_diverso" Then
'        movimento_oleo_diverso.Show
'    ElseIf pNomeProgramaInterno = "movimento_pedido_combustivel" Then
'        movimento_pedido_combustivel.Show
'    ElseIf pNomeProgramaInterno = "movimento_previsao_venda_prazo" Then
'        movimento_previsao_venda_prazo.Show
'    ElseIf pNomeProgramaInterno = "movimento_saida_transferencia_produto" Then
'        movimento_saida_transferencia_produto.Show
'    ElseIf pNomeProgramaInterno = "movimento_vale_caixa" Then
'        movimento_vale_caixa.Show
'    ElseIf pNomeProgramaInterno = "movimento_venda_conveniencia" Then
'        movimento_venda_conveniencia.Show
'    ElseIf pNomeProgramaInterno = "MovimentoVendaConveniencia2" Then
'        MovimentoVendaConveniencia2.Show
'    ElseIf pNomeProgramaInterno = "pro_encerramento_inventario" Then
'        pro_encerramento_inventario.Show
'    ElseIf pNomeProgramaInterno = "processamento_custo_combustivel" Then
'        processamento_custo_combustivel.Show
'    ElseIf pNomeProgramaInterno = "processamento_custo_produto" Then
'        processamento_custo_produto.Show
'    ElseIf pNomeProgramaInterno = "processamento_estoque" Then
'        processamento_estoque.Show
'    ElseIf pNomeProgramaInterno = "processamento_estoque_combustivel" Then
'        processamento_estoque_combustivel.Show
''Relatórios
'    ElseIf pNomeProgramaInterno = "EmissaoAnaliseNotaAbast" Then
'        EmissaoAnaliseNotaAbast.Show
'    ElseIf pNomeProgramaInterno = "emissao_advertencia" Then
'        emissao_advertencia.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_geral" Then
'        emissao_analise_geral.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_giro_estoque" Then
'        emissao_analise_giro_estoque.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_inventario" Then
'        emissao_analise_inventario.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_movimentacao_postos" Then
'        emissao_analise_movimentacao_postos.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_venda_cartao" Then
'        emissao_analise_venda_cartao.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_vendas_funcionarios" Then
'        emissao_analise_vendas_funcionarios.Show
'    ElseIf pNomeProgramaInterno = "emissao_analise_venda_produto" Then
'        emissao_analise_venda_produto.Show
'    ElseIf pNomeProgramaInterno = "EmissaoAnaliseMovEstoque" Then
'        EmissaoAnaliseMovEstoque.Show
'    ElseIf pNomeProgramaInterno = "emissao_balanco" Then
'        emissao_balanco.Show
'    ElseIf pNomeProgramaInterno = "emissao_caixa_simplificado" Then
'        emissao_caixa_simplificado.Show
'    ElseIf pNomeProgramaInterno = "emissao_caixa_pista" Then
'        emissao_caixa_pista.Show
'    ElseIf pNomeProgramaInterno = "EmissaoCalculoPrecoMedio" Then
'        EmissaoCalculoPrecoMedio.Show
'    ElseIf pNomeProgramaInterno = "emissao_conta_pagar_conferencia" Then
'        emissao_conta_pagar_conferencia.Show
'    ElseIf pNomeProgramaInterno = "emissao_cupom_complementar" Then
'        emissao_cupom_complementar.Show
'    ElseIf pNomeProgramaInterno = "emissao_cupom_complementar_conv" Then
'        emissao_cupom_complementar_conv.Show
'    ElseIf pNomeProgramaInterno = "EmissaoFluxoCaixa" Then
'        EmissaoFluxoCaixa.Show
'    ElseIf pNomeProgramaInterno = "emissao_grps" Then
'        emissao_grps.Show
'    ElseIf pNomeProgramaInterno = "emissao_kit_documento_funcionario" Then
'        emissao_kit_documento_funcionario.Show
'    ElseIf pNomeProgramaInterno = "emissao_mapa_resumo" Then
'        emissao_mapa_resumo.Show
'    ElseIf pNomeProgramaInterno = "EmissaoMapaResumoCorrecao" Then
'        EmissaoMapaResumoCorrecao.Show
'    ElseIf pNomeProgramaInterno = "emissao_medida_tanque" Then
'        emissao_medida_tanque.Show
'    ElseIf pNomeProgramaInterno = "emissao_memoria_fiscal" Then
'        emissao_memoria_fiscal.Show
'    ElseIf pNomeProgramaInterno = "emissao_planilha_produtos" Then
'        emissao_planilha_produtos.Show
'    ElseIf pNomeProgramaInterno = "emissao_plano_conta" Then
'        emissao_plano_conta.Show
'    ElseIf pNomeProgramaInterno = "emissao_recibo_folha_pagamento" Then
'        emissao_recibo_folha_pagamento.Show
'    ElseIf pNomeProgramaInterno = "emissao_resumo_folha_pagamento" Then
'        emissao_resumo_folha_pagamento.Show
'    ElseIf pNomeProgramaInterno = "emissao_resumo_movimentacao_postos" Then
'        emissao_resumo_movimentacao_postos.Show
'    ElseIf pNomeProgramaInterno = "EmissaoResumoVendaCartao" Then
'        EmissaoResumoVendaCartao.Show
'    ElseIf pNomeProgramaInterno = "emissao_rpa" Then
'        emissao_rpa.Show
'    ElseIf pNomeProgramaInterno = "emissao_cesta_basica" Then
'        emissao_cesta_basica.Show
'    ElseIf pNomeProgramaInterno = "emissao_cliente_cheque" Then
'        emissao_cliente_cheque.Show
'    ElseIf pNomeProgramaInterno = "emissao_funcionario" Then
'        emissao_funcionario.Show
'    ElseIf pNomeProgramaInterno = "emissao_funcionario_ficha" Then
'        emissao_funcionario_ficha.Show
'    ElseIf pNomeProgramaInterno = "emissao_lmc_matricial" Then
'        emissao_lmc_matricial.Show
'    ElseIf pNomeProgramaInterno = "emissao_movimento_bomba" Then
'        emissao_movimento_bomba.Show
'    ElseIf pNomeProgramaInterno = "emissao_nota_cliente" Then
'        emissao_nota_cliente.Show
'    ElseIf pNomeProgramaInterno = "emissao_nota_cliente_matricial" Then
'        emissao_nota_cliente_matricial.Show
'    ElseIf pNomeProgramaInterno = "emissao_recibo" Then
'        emissao_recibo.Show
'    ElseIf pNomeProgramaInterno = "emissao_suspencao" Then
'        emissao_suspencao.Show
'    ElseIf pNomeProgramaInterno = "emissao_tabela_medida_tanque" Then
'        emissao_tabela_medida_tanque.Show
'    ElseIf pNomeProgramaInterno = "emissao_vale_transporte" Then
'        emissao_vale_transporte.Show
'    ElseIf pNomeProgramaInterno = "EmissaoLivroPrecoDiferenciado" Then
'        EmissaoLivroPrecoDiferenciado.Show
'    ElseIf pNomeProgramaInterno = "EmissaoMovimentoCaixa" Then
'        EmissaoMovimentoCaixa.Show
'    ElseIf pNomeProgramaInterno = "emissao_movimento_digitacao" Then
'        emissao_movimento_digitacao.Show
'    ElseIf pNomeProgramaInterno = "emissao_movimento_lubrificante" Then
'        emissao_movimento_lubrificante.Show
'    ElseIf pNomeProgramaInterno = "emissao_movimentacao_diaria" Then
'        emissao_movimentacao_diaria.Show
'    ElseIf pNomeProgramaInterno = "emissao_venda_cliente" Then
'        emissao_venda_cliente.Show
'    ElseIf pNomeProgramaInterno = "EmissaoCartaFrete" Then
'        EmissaoCartaFrete.Show
'    ElseIf pNomeProgramaInterno = "EmissaoResumoCupomFiscal" Then
'        EmissaoResumoCupomFiscal.Show
'
'
'    ElseIf pNomeProgramaInterno = "frm_emissao_cheques_folhas" Then
'        frm_emissao_cheques_folhas.Show
'    ElseIf pNomeProgramaInterno = "emissao_cheque_formulario" Then
'        emissao_cheque_formulario.Show
'    ElseIf pNomeProgramaInterno = "emissao_preco_combustivel" Then
'        emissao_preco_combustivel.Show
'    ElseIf pNomeProgramaInterno = "frm_emissao_lmc" Then
'        frm_emissao_lmc.Show
'    ElseIf pNomeProgramaInterno = "frm_emissao_recibo_folhas" Then
'        frm_emissao_recibo_folhas.Show
'    ElseIf pNomeProgramaInterno = "frm_emissao_recibo_formulario" Then
'        frm_emissao_recibo_formulario.Show
'    ElseIf pNomeProgramaInterno = "listagem_cheque_formulario" Then
'        listagem_cheque_formulario.Show
'    ElseIf pNomeProgramaInterno = "lstComparaCupomLubrif" Then
'        lstComparaCupomLubrif.Show
'    ElseIf pNomeProgramaInterno = "lst_auditoria" Then
'        lst_auditoria.Show
'    ElseIf pNomeProgramaInterno = "lst_baixa_cheque_devolvido" Then
'        lst_baixa_cheque_devolvido.Show
'    ElseIf pNomeProgramaInterno = "lst_baixa_cheque_devolvido_descontado" Then
'        lst_baixa_cheque_devolvido_descontado.Show
'    ElseIf pNomeProgramaInterno = "lst_baixa_contas_a_pagar_fornecedor" Then
'        lst_baixa_contas_a_pagar_fornecedor.Show
'    ElseIf pNomeProgramaInterno = "lst_baixa_pagar" Then
'        lst_baixa_pagar.Show
'    ElseIf pNomeProgramaInterno = "lst_bordero_deposito" Then
'        lst_bordero_deposito.Show
'    ElseIf pNomeProgramaInterno = "lst_bordero_deposito_avista" Then
'        lst_bordero_deposito_avista.Show
'    ElseIf pNomeProgramaInterno = "lst_cheque" Then
'        lst_cheque.Show
'    ElseIf pNomeProgramaInterno = "lst_cheque_devolvido" Then
'        lst_cheque_devolvido.Show
'    ElseIf pNomeProgramaInterno = "lst_cheque_avista" Then
'        lst_cheque_avista.Show
'    ElseIf pNomeProgramaInterno = "lst_cheque_baixados" Then
'        lst_cheque_baixados.Show
'    ElseIf pNomeProgramaInterno = "lst_cheque_deposito_bancario" Then
'        lst_cheque_deposito_bancario.Show
'    ElseIf pNomeProgramaInterno = "lst_cliente" Then
'        lst_cliente.Show
'    ElseIf pNomeProgramaInterno = "lst_cliente_conveniado" Then
'        lst_cliente_conveniado.Show
'    ElseIf pNomeProgramaInterno = "lst_contas_pagar" Then
'        lst_contas_pagar.Show
'    ElseIf pNomeProgramaInterno = "lst_contas_pagar2" Then
'        lst_contas_pagar2.Show
'    ElseIf pNomeProgramaInterno = "lst_contas_pagar_especial" Then
'        lst_contas_pagar_especial.Show
'    ElseIf pNomeProgramaInterno = "lst_contas_a_pagar_fornecedor" Then
'        lst_contas_a_pagar_fornecedor.Show
'    ElseIf pNomeProgramaInterno = "lst_demonstracao_encerrante" Then
'        lst_demonstracao_encerrante.Show
'    ElseIf pNomeProgramaInterno = "lst_duplicata_paga" Then
'        lst_duplicata_paga.Show
'    ElseIf pNomeProgramaInterno = "lst_duplicata_receber" Then
'        lst_duplicata_receber.Show
'    ElseIf pNomeProgramaInterno = "lst_entrada_combustivel" Then
'        lst_entrada_combustivel.Show
'    ElseIf pNomeProgramaInterno = "lst_entrada_produto" Then
'        lst_entrada_produto.Show
'    ElseIf pNomeProgramaInterno = "lst_entrada_produto_conferencia" Then
'        lst_entrada_produto_conferencia.Show
'    ElseIf pNomeProgramaInterno = "emissao_extrato_bancario" Then
'        lst_extrato_bancario.Show
'    ElseIf pNomeProgramaInterno = "lst_falta_caixa" Then
'        lst_falta_caixa.Show
'    ElseIf pNomeProgramaInterno = "lst_falta_funcionario" Then
'        lst_falta_funcionario.Show
'    ElseIf pNomeProgramaInterno = "lst_estoque_medio" Then
'        lst_estoque_medio.Show
''    ElseIf pNomeProgramaInterno = "lst_historico" Then
''        lst_historico.Show
'    ElseIf pNomeProgramaInterno = "lst_inventario_produto" Then
'        lst_inventario_produto.Show
'    ElseIf pNomeProgramaInterno = "lst_impressao_duplicata" Then
'        lst_impressao_duplicata.Show
'    ElseIf pNomeProgramaInterno = "lst_lmc_abertura" Then
'        lst_lmc_abertura.Show
'    ElseIf pNomeProgramaInterno = "lst_nota_abastecimento_convenio" Then
'        lst_nota_abastecimento_convenio.Show
'    ElseIf pNomeProgramaInterno = "lst_nota_cliente_emissao" Then
'        lst_nota_cliente_emissao.Show
'    ElseIf pNomeProgramaInterno = "lst_nota_cliente_geral" Then
'        lst_nota_cliente_geral.Show
'    ElseIf pNomeProgramaInterno = "lst_movimentacao_estoque" Then
'        lst_movimentacao_estoque.Show
'    ElseIf pNomeProgramaInterno = "lst_perda_sobra" Then
'        lst_perda_sobra.Show
'    ElseIf pNomeProgramaInterno = "lst_quadro_funcionario" Then
'        lst_quadro_funcionario.Show
'    ElseIf pNomeProgramaInterno = "lst_resumo_lmc" Then
'        lst_resumo_lmc.Show
'    ElseIf pNomeProgramaInterno = "lst_venda_conveniencia" Then
'        lst_venda_conveniencia.Show
'    ElseIf pNomeProgramaInterno = "lst_venda_cupom" Then
'        lst_venda_cupom.Show
'    ElseIf pNomeProgramaInterno = "lst_vendas_lmc_ecf" Then
'        lst_vendas_lmc_ecf.Show
'    ElseIf pNomeProgramaInterno = "relatorio_cheque_cobranca" Then
'        relatorio_cheque_cobranca.Show
'    ElseIf pNomeProgramaInterno = "relatorio_cheque_folha" Then
'        relatorio_cheque_folha.Show
'    Else
'        Screen.MousePointer = 1
'    End If
End Sub

Public Sub ChamaProgramaVB6ClickMenu(ByVal pTipo As String, ByVal pNomeInterno As String, ByVal pStringChamada As String)
    gStringChamada = pStringChamada
    Call ClickMenu(pTipo, pNomeInterno)
End Sub

Private Sub ClickMenu(ByVal x_tipo As String, ByVal xNomeMenu As String)
    Dim i As Integer
    Dim xNomeInterno As String
    
'    'Prepara SQL
'    lSQl = ""
'    lSQl = lSQl & "   SELECT Interno"
'    lSQl = lSQl & "     FROM Menu"
'    lSQl = lSQl & "    WHERE Usuario = " & g_usuario
'    lSQl = lSQl & "      AND Tipo = " & preparaTexto(x_tipo)
'    lSQl = lSQl & "      AND Menu = " & preparaTexto(xNomeMenu)
'
'    'Abre RecordSet
'    Set RstMenu = New adodb.Recordset
'    Set RstMenu = Conectar.RsConexao(lSQl)
'
'    xNomeInterno = Trim(RstMenu("Interno").Value)
'    RstMenu.Close
'    Set RstMenu = Nothing
    
    If Programa.LocalizarNomeMenu(x_tipo, xNomeMenu) Then
        xNomeInterno = Programa.NomeInterno
    End If
    
    ' NFCe Complementar*
    If xNomeInterno = "EmissaoNFCeComplementar" Then
        If g_automacao = True Then
            xNomeInterno = "EmissaoNFCeComplementarAuto"
        End If
    End If
    
    
    If g_empresa = 0 Then
        MsgBox "Selecione uma empresa!", vbInformation, "Empresa Não Selecionada!"
        dtcbo_empresa.SetFocus
        Exit Sub
    End If
    If xNomeInterno = "CPcadastroCliente" Then
        Call GravaCheqPostoIni(Mid(xNomeInterno, 3, Len(xNomeInterno) - 2))
    ElseIf xNomeInterno = "CPconsultaCheq" Then
        Call GravaCheqPostoIni(Mid(xNomeInterno, 3, Len(xNomeInterno) - 2))
    ElseIf xNomeInterno = "CadastroPeriodoTrocaOleo" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ConfiguracaoDiversa" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "CadastroContaTesouraria" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "CadastroTipoMovimentoTesouraria" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf x_tipo = "CA" Then
        Call GravaSgpCadastroIni(xNomeInterno)
        
        
'*******************************
'*** CONSULTA VB.NET         ***
'*******************************
    
        
    ElseIf xNomeInterno = "ConfereNotaAbastecimento" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
        
        
'*******************************
'*** MOVIMENTAÇÃO VB.NET     ***
'*******************************
    
    ElseIf xNomeInterno = "ArquivaBancoDados" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "AtualizaCadastro" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ConverteBancoDados" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EnviaNFeEmail" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ExportaImportaCadastro" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ExportaDadosFortes" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ExportaDadosInforlub" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "GeraRemessaDuplicata" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovAlteraPrecoAutomacao" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoAbastecimento" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoAcertoVendaECF" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBaixaCartao" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBaixaFaltaCaixa" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBancario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBomba" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBombaMec" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoBombaPorComb" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoCartaFrete" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoCupomFiscalNovo" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoDiario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "mov_nota_abastecimento" And ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoAberturaCaixa" Then
        gStringChamada = ""
        'Call GravaSgpCadastroIni(xNomeInterno)
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoCaixaPista" Then
'        Call GravaSgpCadastroIni(xNomeInterno)
'        Exit Sub
        gStringChamada = ""
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoCustodiaCheque" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoDespesa" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoEncerranteProduto" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoImportaNFeEntrada" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoEnviaMensagemEmail" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "movimento_entrada_produto" Then
        If gArqTxt.FolderExists("C:\Cerrado.Net\SgpNet") Then
            Call GravaSgpNetCadastroIni(xNomeInterno)
        Else
            Call Chama(xNomeInterno)
        End If
    ElseIf xNomeInterno = "MovimentoFinanceiro" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoHorarioVerao" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "movimento_cheque_devolvido" And ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoNotaFiscalSaida" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoNFeSaida" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoNFeSaidaCancelada" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoNFeSaidaInutilizada" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "movimento_vale_abastecimento_emitido" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "movimento_vale_abastecimento_recebido" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoOrdemCompra" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoPonto" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoLivroLMC" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoLocalizacaoProduto" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoObservacao" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EnviaDadosEmail" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ImportaDadosPista" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "RecebeDadosEmail" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ImportaAbastecimento" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoNFeCartaCorrecao" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "MovimentoTesouraria" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoLivroCaixa" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoHistoricoPrecoCombustivel" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
        

'*******************************
'*** RELATORIOS VB.NET       ***
'*******************************
    
    
    ElseIf xNomeInterno = "CupomComplementarConveniencia" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "CupomComplementarEstoque" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoAbastecimentoAutomacao" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoAbastecimentoAutomacaoConf" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoAbastecimentoAutomacaoFuncionario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoCartaAniversariante" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoCartaoCredito" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoCartaoCreditoBaixado" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoCartaoCreditoRecebido" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoComissaoVenda" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoConciliaEstoqueComb" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoConciliaEstoqueCombEmp" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoCustodiaCheque" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoDescontoPersonalizado" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoDespesaFornecedor" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoDiferencaInventario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoEspelhoNfProduto" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoFaltaCaixaGeral" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoFaturamentoBrutoMensal" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoFuncionario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoGrupo" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoInventarioContabil" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoLivroPreco" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoLucroBrutoVenda" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoNotaAbastecimentoBaixada" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoNotaAbastecimentoFrota" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoPlanilhaCaixa" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoValeAbastecimento" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoValeAbastRecebido" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoVendaEcfResumida" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoNotaAbastecimentoExcel" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "TefEmissaoExtratoCliente" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "ConsultaNFCeEmitidaXML" Then
        Call GravaSgpNetCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoAnaliseVendaCombustivelAuto" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "EmissaoRankingVendasPorFuncionario" Then
        Call GravaSgpCadastroIni(xNomeInterno)
    ElseIf xNomeInterno = "AnaliseVendaCombustivelNivelPrecoAuto" Then
        Call GravaSgpCadastroIni(xNomeInterno)
        
        
'**************************************
'*** RELATORIOS PROGRAMA EXTERNO C# ***
'**************************************
    ElseIf xNomeInterno = "EmissaoNFCeAutorizada" Then
      Dim xRetorno As String
      Dim xCaminho As String
        xCaminho = "C:\Cerrado Tecnologia\Petromovel\Petromovel.exe"
        If gArqTxt.FileExists(xCaminho) Then
            xRetorno = Shell(xCaminho & " " & g_empresa & " " & Replace(g_nome_empresa, " ", "_") & " " & g_usuario & " " & Replace(g_nome_usuario, " ", "_") & " " & g_nivel_acesso & " " & Replace(gDrive, ":", "") & " " & gIpBanco & " " & Replace(gPortaBanco, ",", "") & " " & gNomeInternoBD & " " & gSenhaBD & " " & gCNPJEmpresa & " " & Empresa.Cidade & " " & "frmRelNFCeAutorizada", vbNormalFocus)
        Else
            MsgBox "Programa não encontrado!"
        End If
    ElseIf xNomeInterno = "EmissaoNFeAutorizada" Then
      Dim xRetorno2 As String
      Dim xCaminho2 As String
         
        xCaminho2 = "C:\Cerrado Tecnologia\Petromovel\Petromovel.exe"
        If gArqTxt.FileExists(xCaminho2) Then
            xRetorno2 = Shell(xCaminho2 & " " & g_empresa & " " & Replace(g_nome_empresa, " ", "_") & " " & g_usuario & " " & Replace(g_nome_usuario, " ", "_") & " " & g_nivel_acesso & " " & Replace(gDrive, ":", "") & " " & gIpBanco & " " & Replace(gPortaBanco, ",", "") & " " & gNomeInternoBD & " " & gSenhaBD & " " & gCNPJEmpresa & " " & Empresa.Cidade & " " & "frmRelNFeAutorizada", vbNormalFocus)
        Else
            MsgBox "Programa não encontrado!"
        End If
    ElseIf xNomeInterno = "ManutencaoAbastecimento" Then
      Dim xRetorno3 As String
      Dim xCaminho3 As String
         
        xCaminho3 = "C:\Cerrado Tecnologia\Petromovel\Petromovel.exe"
        If gArqTxt.FileExists(xCaminho3) Then
            xRetorno2 = Shell(xCaminho3 & " " & g_empresa & " " & Replace(g_nome_empresa, " ", "_") & " " & g_usuario & " " & Replace(g_nome_usuario, " ", "_") & " " & g_nivel_acesso & " " & Replace(gDrive, ":", "") & " " & gIpBanco & " " & Replace(gPortaBanco, ",", "") & " " & gNomeInternoBD & " " & gSenhaBD & " " & gCNPJEmpresa & " " & Empresa.Cidade & " " & "frmManutencaoAbastecimento", vbNormalFocus)
        Else
            MsgBox "Programa não encontrado!"
        End If
    Else
        Call Chama(xNomeInterno)
    End If
End Sub
Private Sub CriaConfiguracao()
    Configuracao.Empresa = g_empresa
    Configuracao.QuantidadePeriodos = 2
    Configuracao.QuantidadeBico = 6
    Configuracao.PularBomba = 0
    Configuracao.CustoDuplicata = 0
    
    Configuracao.ValorSuperior = 0.6
    Configuracao.ValorEsquerda = 13.5
    Configuracao.Extenso1Superior = 1.6
    Configuracao.Extenso1Esquerda = 1.7
    Configuracao.Extenso2Superior = 2.3
    Configuracao.Extenso2Esquerda = 0.5
    Configuracao.FavorecidoSuperior = 2.9
    Configuracao.FavorecidoEsquerda = 0.5
    Configuracao.CidadeSuperior = 3.5
    Configuracao.CidadeEsquerda = 8.5
    Configuracao.DiaSuperior = 3.5
    Configuracao.DiaEsquerda = 10.5
    Configuracao.MesSuperior = 3.5
    Configuracao.MesEsquerda = 11.7
    Configuracao.AnoSuperior = 3.5
    Configuracao.AnoEsquerda = 15.9
    
    Configuracao.OutrasConfiguracoes = "NNNNN00NN00000000   "
    Configuracao.MensagemCobranca = "."
    Configuracao.QuantidadeIlha = 1
    Configuracao.ProgramacaoAntiga = False
    Configuracao.HoraFechamento1 = CDate("00:00:00")
    Configuracao.HoraFechamento2 = CDate("00:00:00")
    Configuracao.HoraFechamento3 = CDate("00:00:00")
    Configuracao.HoraFechamento4 = CDate("00:00:00")
    Configuracao.HoraFechamento5 = CDate("00:00:00")
    Configuracao.HoraFechamento6 = CDate("00:00:00")
    Configuracao.HoraFechamento7 = CDate("00:00:00")
    Configuracao.HoraFechamento8 = CDate("00:00:00")
    
    Configuracao.ImprimirReducaoZ = False
    Configuracao.QuantidadeViasTEF = 1
    Configuracao.ControleSolicitacaoTEF = 100
    Configuracao.IntegraMovimentoBombaCaixa = True
    Configuracao.AlteraAberturaBomba = False
    Configuracao.AlteraPrecoMovimentoBomba = False
    Configuracao.ECFBaixaEstoque = True
    Configuracao.NomeclaturaCaixa = "PERIODO"
    Configuracao.AlteracaoCaixaPeloResponsavel = True
    Configuracao.AlteraPrecoProdutoPelaVenda = True
    Configuracao.InverteEncerrantenaPlanilha = False
    Configuracao.IdentificaFuncionarioaCadaCupom = True
    Configuracao.RelacaoNotasnoCaixa = True
    Configuracao.BloqueiaVendaPeloEstoque = False
    Configuracao.BloqueiaVendaPeloSubEstoque = False
    Configuracao.NumeroDuplicata = 0
    
    If Not Configuracao.Incluir Then
        MsgBox "Erro ao incluir o registro de configuração do sistema!", vbInformation, "Erro de Integridade!"
    End If
End Sub
Public Sub AtivaVerificacaoGIC()
    lNotificacaoGic = True
    cmdConexaoGic.ToolTipText = "Desativa verificação de comunicação com o GIC."
    If gArqTxt.FileExists("c:\vb5\sgp\icons\PEOPLE 061.ICO") Then
        cmdConexaoGic.Picture = LoadPicture("c:\vb5\sgp\icons\PEOPLE 061.ICO")
    End If
    StatusBar1.Panels(3).Text = "Verificando GIC"
    StatusBar1.Panels(3).ToolTipText = "Verificando comunicação no GIC"
    StatusBar1.Panels(3).AutoSize = sbrContents
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "Empresa Global") Then
        gEmpresaGlobal = ConfiguracaoDiversa.Codigo
    End If
End Sub
Public Sub DesativaVerificacaoGIC()
    gEmpresaGlobal = 0
    lNotificacaoGic = False
    cmdConexaoGic.ToolTipText = "Ativa verificação de comunicação com o GIC."
    If gArqTxt.FileExists("c:\vb5\sgp\icons\message.ICO") Then
        cmdConexaoGic.Picture = LoadPicture("c:\vb5\sgp\icons\message.ICO")
    End If
    StatusBar1.Panels(3).Text = "Não verificando GIC"
    StatusBar1.Panels(3).ToolTipText = "Não está verificando comunicação no GIC"
    StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "S.G.P.")
    dados.Empresa = g_empresa
    If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", lNomeArquivo)
    End If
    If gArqTxt.FolderExists("C:\Cerrado.Net\SgpNet") Then
        Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini")
    End If
    If g_conta_bancaria <> "" Then
        dados.ContaBancaria = g_conta_bancaria
        dados.ContaBancaria = " "
    End If
     
    If lResgateAbertoPeloSGP Then
        FinalizaPrograma ("NFeResgaste.exe")
    End If
    
    If PetromovelEstaAtivo Then
        Call GravaAtividadePetromovelFinalizado
        FinalizaPrograma ("PetromovelAuto.exe")
    End If
     
    If Not dados.Alterar(1) Then
        MsgBox "Registro não alterado.", vbInformation, "Erro de Integridade"
    End If
    
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set dados = Nothing
    Set Empresa = Nothing
    Set LiberacaoDigitacao = Nothing
    Set Programa = Nothing
    
    If bdAccess Then
        bd_sgp.Close
        'bd_sgp_b.Close
        'bd_sgp_m.Close
        'cnnSGPb.Close
        'cnnSGPm.Close
        cnnSGP.Close
    End If
    End
End Sub
Private Sub FinalizaPrograma(ByVal pNomePrograma As String)

    On Error GoTo FileError

    Dim Comando As String
    Comando = "TASKKILL -F -IM " & pNomePrograma
    Shell Comando
    
FileError:
    Exit Sub
End Sub
Private Function FormulariosFechados() As Boolean
    Dim frm As Form
    Dim i As Integer

    FormulariosFechados = False
    i = 0
    For Each frm In Forms
        If frm.name <> "menu_personalizado" Then
            i = i + 1
        End If
    Next frm
    If i > 0 Then
        If (MsgBox("Tem " & i & " tela(s) do sistema aberto(s)." & Chr(10) & "Deseja fecha-lo(s) automaticamente?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!") = vbYes) Then
            For Each frm In Forms
                If frm.name <> "menu_personalizado" Then
                    Unload frm
                End If
            Next frm
            i = 0
            For Each frm In Forms
                If frm.name <> "menu_personalizado" Then
                    i = i + 1
                End If
            Next frm
            If i = 0 Then
                FormulariosFechados = True
            End If
        End If
    Else
        FormulariosFechados = True
    End If
End Function
Public Function GravaSgpCadastroIni(ByRef xNomePrograma As String) As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim retval As Long
    
    On Error GoTo FileError
    
    GravaSgpCadastroIni = False
    'lNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    retval = Len(lNomeArquivo)
    xArquivoTmp = Mid(lNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(lNomeArquivo) Then
        Call gArqTxt.DeleteFile(lNomeArquivo, True)
    End If
    
    Call GravaAuditoria(1, Me.name, 21, g_empresa & "-" & g_nome_empresa)
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    gArquivoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("EmpresaGlobalAzure=" & Format(gEmpresaGlobal, "000"))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=" & xNomePrograma)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    gArquivoTMP.WriteLine ("UsuarioGlobal=" & CStr(gUsuarioGlobal))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    gArquivoTMP.WriteLine ("InternetBandaLarga=" & CStr(gInternetBandaLarga))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    gArquivoTMP.Close
    
    
    'Call WriteINI("Empresa", "Empresa", Format(g_empresa, "000"), xArquivoTmp)
    'Call WriteINI("Empresa", "NomeEmpresa", g_nome_empresa, xArquivoTmp)
    'Call WriteINI("Empresa", "CidadeEmpresa", g_cidade_empresa, xArquivoTmp)
    
    'Call WriteINI("Programa", "Programa", xNomePrograma, xArquivoTmp)
    
    'Call WriteINI("Usuario", "Usuario", CStr(g_usuario), xArquivoTmp)
    'Call WriteINI("Usuario", "NomeUsuario", g_nome_usuario, xArquivoTmp)
    'Call WriteINI("Usuario", "NivelAcessoUsuario", CStr(g_nivel_acesso), xArquivoTmp)
    
    'Call WriteINI("Outras", "DataDef", CStr(g_data_def), xArquivoTmp)
    'Call WriteINI("Outras", "FlagLmc", CStr(g_lmc), xArquivoTmp)
    'Call WriteINI("Outras", "ImpressoraMatricial", CStr(g_impressora_matricial), xArquivoTmp)
    'Call WriteINI("Outras", "CaixaUnificado", CStr(g_caixa_unificado), xArquivoTmp)
    
    'Call WriteINI("ContaBancaria", "ContaBancaria", CStr(g_conta_bancaria), xArquivoTmp)
    'Call WriteINI("ContaBancaria", "NomeContaBancaria", CStr(g_nome_conta), xArquivoTmp)
    
    'Call WriteINI("Liberacao", "EmpresaInicial", CStr(g_cfg_empresa_i), xArquivoTmp)
    'Call WriteINI("Liberacao", "EmpresaFinal", CStr(g_cfg_empresa_f), xArquivoTmp)
    'Call WriteINI("Liberacao", "DataInicial", CStr(g_cfg_data_i), xArquivoTmp)
    'Call WriteINI("Liberacao", "DataFinal", CStr(g_cfg_data_f), xArquivoTmp)
    'Call WriteINI("Liberacao", "PeriodoInicial", CStr(g_cfg_periodo_i), xArquivoTmp)
    'Call WriteINI("Liberacao", "PeriodoFinal", CStr(g_cfg_periodo_f), xArquivoTmp)
    
    'Call WriteINI("Automacao", "Automacao", CStr(g_automacao), xArquivoTmp)
    
    Call gArqTxt.MoveFile(xArquivoTmp, lNomeArquivo)

    If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "NAO" Then
        'Cadastro VB6
        retval = Shell("C:" & gDiretorioAplicativo & "SGP_CADASTRO.exe", vbMinimizedNoFocus)
    End If
    GravaSgpCadastroIni = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar Sgp_Cadastro.ini!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Public Function GravaSgpNetCadastroIni(ByRef pNomePrograma As String) As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim retval As Long
    Dim xNomeArquivo As String
    Dim xProcessaNFCe As Boolean
    Dim xFaseErro As Integer
        
    On Error GoTo FileError
    
    xFaseErro = 0
    If pNomePrograma = "ProcessaNFCe" Then
        xNomeArquivo = "C:\Cerrado.Net\SgpNet\SgpNetTemporarioNFCE.ini"
        xProcessaNFCe = True
    Else
        xNomeArquivo = "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini"
        xProcessaNFCe = False
    End If

    GravaSgpNetCadastroIni = False
    'lNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    retval = Len(xNomeArquivo)
    xArquivoTmp = Mid(xNomeArquivo, 1, retval - 3) & "tmp"
    xFaseErro = 10
    If gArqTxt.FileExists(xArquivoTmp) Then
        xFaseErro = 20
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    
    xFaseErro = 30
    If gArqTxt.FileExists(xNomeArquivo) And xProcessaNFCe = True Then
        CriaLogSGP "GravaSgpNetCadastroIni: Arquivo será substituido:" & xNomeArquivo, "", "pNomePrograma=" & pNomePrograma
        xFaseErro = 40
        Call CopiaArquivoDataHora(xNomeArquivo) 'Renomeia para Colocar em fila e posteriormente ser processado
        xFaseErro = 50
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    
    
    xFaseErro = 60
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    xFaseErro = 70
    gArquivoTMP.WriteLine ("[Empresa]")
    xFaseErro = 80
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    gArquivoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("EmpresaGlobalAzure=" & Format(gEmpresaGlobal, "000"))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=" & pNomePrograma)
    
    xFaseErro = 200
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    
    xFaseErro = 300
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    xFaseErro = 400
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    xFaseErro = 500
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    xFaseErro = 600
    gArquivoTMP.Close
    xFaseErro = 700
    
    'Coloquei esse teste, pois se clicar mais de uma fez rápido em caixa de pista dá
    'a mensagem de erro.
    If pNomePrograma <> "ProcessaNFCe" Then
        xFaseErro = 710
        If gArqTxt.FileExists(xNomeArquivo) Then
            xFaseErro = 720
            Call gArqTxt.DeleteFile(xNomeArquivo, True)
            xFaseErro = 730
        End If
    End If
    
    Call gArqTxt.MoveFile(xArquivoTmp, xNomeArquivo)
    xFaseErro = 750
    GravaSgpNetCadastroIni = True
    Exit Function

FileError:
    Call CriaLogSGP("[GravaSgpNetCadastroIni]", " - ERRO - Chamada de Programa .SGPNET: gStringChamada=" & gStringChamada & " - pNomePrograma=" & pNomePrograma & " - xFaseErro=" & xFaseErro & " - xNomeArquivo=" & xNomeArquivo & " - Error=" & Err.Description, "")
    If pNomePrograma = "ProcessaNFCe" Then
        Call CriaLogCupom(Time & " - ERRO - Chamada ProcessaNFCe: gStringChamada=" & gStringChamada & " - pNomePrograma=" & pNomePrograma & " - Error=" & Err.Description)
    End If
    MsgBox "Erro ao gravar SgpNetCadastro.ini!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Public Function GravaSgpNetCadastroIni2(ByRef pNomePrograma As String) As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim retval As Long
    Dim xNomeArquivo As String
        
    On Error GoTo FileError
    
    xNomeArquivo = "C:\Cerrado.Net\SgpNet\SgpNetTemporario2.ini"
    GravaSgpNetCadastroIni2 = False
    'lNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    retval = Len(xNomeArquivo)
    xArquivoTmp = Mid(xNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    gArquivoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("EmpresaGlobalAzure=" & Format(gEmpresaGlobal, "000"))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=" & pNomePrograma)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    gArquivoTMP.Close
    
    
    
    Call gArqTxt.MoveFile(xArquivoTmp, xNomeArquivo)

    GravaSgpNetCadastroIni2 = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar SgpNetCadastro.ini2!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Public Function GravaCheqPostoIni(ByRef xNomePrograma As String) As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim xNomeArquivo As String
    Dim retval As Long
    
    On Error GoTo FileError
    
    GravaCheqPostoIni = False
    xNomeArquivo = "C:\Cerrado.Net\CheqPosto\CheqPosto_cadastro.ini"
    xArquivoTmp = Mid(xNomeArquivo, 1, Len(xNomeArquivo) - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=" & xNomePrograma)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    gArquivoTMP.Close
    
    
    Call gArqTxt.MoveFile(xArquivoTmp, xNomeArquivo)

    If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        retval = Shell("C:\Cerrado.Net\CheqPosto\Bin\CheqPosto.exe", vbMinimizedNoFocus)
        GravaCheqPostoIni = True
    Else
        MsgBox "Este programa não poderá ser executado!", vbInformation, "Erro de Configuração!"
    End If
    Exit Function

FileError:
    MsgBox "Erro ao gravar CheqPosto.ini!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function

Private Sub GravaAtividadePetromovelFinalizado()
    Dim xMovAtividade As New CadastroDLL.cMovAtividadeProg
    Const TIPO_ATIVIDADE_FINALIZADO As String = "FINALIZADO"


    xMovAtividade.DataHora_MovAtividadeProgramaExterno = Now
    xMovAtividade.IdEstabelecimento_MovAtividadeProgramaExterno = g_empresa
    xMovAtividade.IpComputadorAC_MovAtividadeProgramaExterno = GetIPAddress
    xMovAtividade.NomePrograma_MovAtividadeProgramaExterno = "PETROMOVEL_AUTO"
    xMovAtividade.Observacao_MovAtividadeProgramaExterno = "FINALIZADO PELO SGP"
    xMovAtividade.Tipo_MovAtividadeProgramaExterno = TIPO_ATIVIDADE_FINALIZADO
    xMovAtividade.VersaoHost_MovAtividadeProgramaExterno = gVersaoSGP

    xMovAtividade.Incluir

End Sub


Private Sub LimpaMenu()
    Do Until i_ca = 0
        Unload mnu_cadastro_item(i_ca)
        i_ca = i_ca - 1
    Loop
    mnu_cadastro_item(i_ca).Caption = ""
    Do Until i_co = 0
        Unload mnu_consulta_item(i_co)
        i_co = i_co - 1
    Loop
    mnu_consulta_item(i_co).Caption = ""
    Do Until i_gr = 0
        Unload mnu_grafico_item(i_gr)
        i_gr = i_gr - 1
    Loop
    mnu_grafico_item(i_gr).Caption = ""
    Do Until i_mo = 0
        Unload mnu_movimentacao_item(i_mo)
        i_mo = i_mo - 1
    Loop
    mnu_movimentacao_item(i_mo).Caption = ""
    Do Until i_re = 0
        Unload mnu_relatorio_item(i_re)
        i_re = i_re - 1
    Loop
    mnu_relatorio_item(i_re).Caption = ""
End Sub
Private Sub MontaMenu()
    Dim i As Integer
    Dim x_tipo As String
    l_sair = 0
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Tipo, Menu"
    lSQL = lSQL & "     FROM Menu"
    lSQL = lSQL & "    WHERE Usuario = " & g_usuario
    lSQL = lSQL & " ORDER BY Tipo, Menu"
    
    'Abre RecordSet
    Set RstMenu = New ADODB.Recordset
    Set RstMenu = Conectar.RsConexao(lSQL)
    
    Do Until RstMenu.EOF
        If RstMenu("Tipo").Value <> x_tipo Then
            x_tipo = RstMenu("Tipo").Value
            i = -1
        End If
        i = i + 1
        If RstMenu("Tipo").Value = "CA" Then
            l_sair = i
            If i > 0 Then
                Load mnu_cadastro_item(i)
            End If
            mnu_cadastro_item(i).Caption = "&" & RstMenu("Menu").Value
        ElseIf RstMenu("Tipo").Value = "CO" Then
            i_co = i
            If i > 0 Then
                Load mnu_consulta_item(i)
            End If
            mnu_consulta_item(i).Caption = "&" & RstMenu("Menu").Value
        ElseIf RstMenu("Tipo").Value = "GR" Then
            i_gr = i
            If i > 0 Then
                Load mnu_grafico_item(i)
            End If
            mnu_grafico_item(i).Caption = "&" & RstMenu("Menu").Value
        ElseIf RstMenu("Tipo").Value = "MO" Then
            i_mo = i
            If i > 0 Then
                Load mnu_movimentacao_item(i)
            End If
            mnu_movimentacao_item(i).Caption = "&" & RstMenu("Menu").Value
        ElseIf RstMenu("Tipo").Value = "RE" Then
            i_re = i
            If i > 0 Then
                Load mnu_relatorio_item(i)
            End If
            mnu_relatorio_item(i).Caption = "&" & RstMenu("Menu").Value
        End If
        RstMenu.MoveNext
    Loop
    RstMenu.Close
    Set RstMenu = Nothing
    l_sair = l_sair + 1
    If l_sair > 0 Then
        Load mnu_cadastro_item(l_sair)
    End If
    mnu_cadastro_item(l_sair).Caption = "-"
    l_sair = l_sair + 1
    i_ca = l_sair
    Load mnu_cadastro_item(l_sair)
    mnu_cadastro_item(l_sair).Caption = "Sai&r"
End Sub
Public Function MudaEmpresaVbNet2003() As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim retval As Long
    
    On Error GoTo FileError
    
    MudaEmpresaVbNet2003 = False
    retval = Len(lNomeArquivo)
    xArquivoTmp = Mid(lNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(lNomeArquivo) Then
        Call gArqTxt.DeleteFile(lNomeArquivo, True)
    End If
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    gArquivoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("EmpresaGlobalAzure=" & Format(gEmpresaGlobal, "000"))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=")
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    gArquivoTMP.WriteLine ("UsuarioGlobal=" & CStr(gUsuarioGlobal))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    gArquivoTMP.WriteLine ("InternetBandaLarga=" & CStr(gInternetBandaLarga))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    gArquivoTMP.Close
    
    Call gArqTxt.MoveFile(xArquivoTmp, lNomeArquivo)

    MudaEmpresaVbNet2003 = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar Mudar empresa .net2003 !" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Public Function MudaEmpresaVbNet2008() As Boolean
    Dim xString As String
    Dim xNomeArquivo As String
    Dim xNomeArquivoTmp As String
    Dim retval As Long
    
    On Error GoTo FileError
    
    MudaEmpresaVbNet2008 = False
    xNomeArquivo = "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini"
    retval = Len(xNomeArquivo)
    xNomeArquivoTmp = Mid(xNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xNomeArquivoTmp) Then
        Call gArqTxt.DeleteFile(xNomeArquivoTmp, True)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xNomeArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    gArquivoTMP.WriteLine ("EmpresaGlobal=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("EmpresaGlobalAzure=" & Format(gEmpresaGlobal, "000"))
    gArquivoTMP.WriteLine ("CNPJEmpresa=" & gCNPJEmpresa) 'CNPJ Empresa
    
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=")
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    gArquivoTMP.Close
    
    Call gArqTxt.MoveFile(xNomeArquivoTmp, xNomeArquivo)
    

    MudaEmpresaVbNet2008 = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar Mudar empresa .net2008 !" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function
Public Function MudaEmpresaPetromovelAuto() As Boolean
    MudaEmpresaPetromovelAuto = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
       If ConfiguracaoDiversa.Verdadeiro Then
          If AtualizaVariaveisGlobaisPetromovelAuto Then
            lEmpresaAtualPetromovel = g_empresa
            MudaEmpresaPetromovelAuto = True
          End If
       End If
    End If
End Function
Private Sub cmdConexaoGic_Click()
'    If cmdConexaoGic.ToolTipText = "Desativa verificação de comunicação com o GIC." Then
'        DesativaVerificacaoGIC
'        lContadorTimer2 = 0
'    Else
'        AtivaVerificacaoGIC
'        lContadorTimer = 600
'    End If

End Sub

Private Sub cmdResgate_Click()

 Dim xRetorno As String
 Dim xCaminho As String
 Dim xCaminho2 As String
 
 xCaminho = "C:\Cerrado Tecnologia\NFeResgaste.exe"
 xCaminho2 = "C:\Cerrado Tecnologia\NFeResgate\NFeResgaste.exe"
If gArqTxt.FileExists(xCaminho) Then
    lResgateAbertoPeloSGP = True
    xRetorno = Shell(xCaminho & " " & g_usuario & " " & Replace(g_nome_usuario, " ", "_") & " " & g_nivel_acesso, vbNormalFocus)
ElseIf gArqTxt.FileExists(xCaminho2) Then
    lResgateAbertoPeloSGP = True
    xRetorno = Shell(xCaminho2 & " " & g_usuario & " " & Replace(g_nome_usuario, " ", "_") & " " & g_nivel_acesso, vbNormalFocus)
Else
    MsgBox "Programa não encontrado!"
End If
    
End Sub

Private Sub cmdTransfereDadosLMC_Click()
    Dim xTransferiu As Boolean
    Dim xUltimaData As Date
    Dim EntradaCombustivel As New cEntradaCombustivel
    Dim MedicaoCombustivel As New cMedicaoCombustivel
    Dim MovAfericao As New cMovimentoAfericao
    Dim MovimentoBomba As New cMovimentoBomba
        
    xTransferiu = False
    EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
    MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
    MovAfericao.NomeTabela = "Movimento_Afericao_LMC"
    MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    
    Call GravaAuditoria(1, Me.name, 23, "Transferencia Geral Para LMC")
    
    If EntradaCombustivel.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & EntradaCombustivel.UltimaData(g_empresa))
        If EntradaCombustivel.TransfereDadosLMC(g_empresa, False) Then
            xTransferiu = True
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem entrada de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    End If
    
    If MedicaoCombustivel.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & MedicaoCombustivel.UltimaData(g_empresa))
        If MedicaoCombustivel.TransfereDadosLMC(g_empresa, False) Then
            xTransferiu = True
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem medição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    End If
    
    If MovAfericao.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & MovAfericao.UltimaData(g_empresa))
        If MovAfericao.TransfereDadosLMC(g_empresa, False) Then
            xTransferiu = True
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem aferição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    End If
    
    If MovimentoBomba.TransfereDadosLMC(g_empresa, True) Then
        xUltimaData = MovimentoBomba.UltimaData(g_empresa)
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & xUltimaData)
        'Exclui Movimento da Última Data
        If Not MovimentoBomba.ExcluirData(g_empresa, xUltimaData) Then
            MsgBox "Não foi possível excluir registros do movimento de bomba.", vbInformation, "Erro de Verificação"
        End If
        'Transfere Dados para o LMC
        If MovimentoBomba.TransfereDadosLMC(g_empresa, False) Then
            xTransferiu = True
            'Recalcula Encerrantes
            If Not MovimentoBomba.RecalculaEncerrante(g_empresa, xUltimaData, 0) Then
                MsgBox "Erro ao recalcular encerrantes.", vbInformation, "Erro de Integridade"
            End If
            'Acumula Períodos
            xUltimaData = MovimentoBomba.LocalizarPDPM1(g_empresa)
            Do Until xUltimaData = "00:00:00"
                If Not MovimentoBomba.AgrupaPeriodoData(g_empresa, xUltimaData) Then
                    MsgBox "Erro ao acumular períodos!", vbInformation, "Erro de Integridade!"
                    Exit Do
                End If
                xUltimaData = MovimentoBomba.LocalizarPDPM1(g_empresa)
            Loop
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem movimento de bomba à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    End If
    
    If xTransferiu Then
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Teve dados novos transferidos para o L.M.C.", vbInformation, "Transferência Concluida!"
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem dados novos à serem transferidos para o L.M.C.", vbInformation, "Transferência Não Efetuada!"
    End If
    
    Set EntradaCombustivel = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovAfericao = Nothing
    Set MovimentoBomba = Nothing
End Sub
'Private Sub cmdWeb_Click()
'    Dim retval As Long
'
'    Screen.MousePointer = 11
'
'    If (MsgBox("Para acessar o site da Cerrado, escolha (Sim)." & vbCrLf & "Para acessar o GIC, escolha (Não)", vbQuestion + vbYesNo + vbDefaultButton1, "Escolha o site a acessar!")) = vbYes Then
'        g_string = "http://www.cerradoinformatica.com/cerradoinformatica"
'    Else
'        If ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni) = "TEIXEIRA E PINHEIRO LTDA" Then
'            g_string = "http://192.168.1.6:8080/GIC"
'        Else
'            g_string = "http://tasso.myvnc.com:8080/GIC"
'        End If
'    End If
'    'cerradoBrowser.Show 0
'    MsgBox "formulário cerradoBrowser removido do projeto"
'    Exit Sub
'
'    On Error GoTo FileError
'    If (MsgBox("Para acessar o site da Cerrado, escolha (Sim)." & "Para acessar o GIC, escolha (Não)", vbQuestion + vbYesNo + vbDefaultButton1, "Escolha o site a acessar!")) = vbYes Then
'        retval = Shell("C:\Arquivos de programas\Internet Explorer\IEXPLORE.EXE http://www.cerradoinformatica.com", vbNormalFocus)
'    Else
'        If ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni) = "TEIXEIRA E PINHEIRO LTDA" Then
'            retval = Shell("C:\Arquivos de programas\Internet Explorer\IEXPLORE.EXE http://192.168.1.6:8080/GIC", vbNormalFocus)
'        Else
'            retval = Shell("C:\Arquivos de programas\Internet Explorer\IEXPLORE.EXE http://tasso.myvnc.com:8080/GIC", vbNormalFocus)
'        End If
'    End If
'    Exit Sub
'
'FileError:
'    MsgBox "Aplicativo inexistente ou localidade desconhecida!", vbInformation, "Erro ao Executar Aplicataivo!"
'End Sub
Private Sub dtcbo_empresa_Click(Area As Integer)
    If dtcbo_empresa.BoundText <> "" Then
        g_empresa = Val(dtcbo_empresa.BoundText)
        g_nome_empresa = dtcbo_empresa.Text
    Else
        dtcbo_empresa.SetFocus
    End If
End Sub
Private Sub dtcbo_empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 7 Then
        KeyAscii = 0
        gera_string_insert.Show
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_senha.SetFocus
    End If
End Sub
Private Sub dtcbo_empresa_LostFocus()
    If dtcbo_empresa.BoundText <> "" Then
        g_empresa = Val(dtcbo_empresa.BoundText)
        g_nome_empresa = dtcbo_empresa.Text
        Call GravaAuditoria(1, Me.name, 21, g_empresa)
    End If
    If g_empresa > 0 Then
        BuscaConfiguracao
        BuscaConfiguracaoCaixaUnificado
        BuscaRegistroLiberacaoDigitacao
        If Empresa.LocalizarCodigo(g_empresa) Then
            g_cidade_empresa = Trim(Empresa.Cidade)
            gUfEmpresa = Trim(Empresa.Estado)
            gEmpresaGlobal = Empresa.EmpresaGlobal
            gCNPJEmpresa = Empresa.CGC 'ALEX - NFCE
        Else
            g_cidade_empresa = "Goiânia"
            gUfEmpresa = "GO"
        End If
        MudaEmpresaVbNet2003
        MudaEmpresaVbNet2008
        If lEmpresaAtualPetromovel <> g_empresa Then
            MudaEmpresaPetromovelAuto
        End If
    End If
End Sub
'Private Sub cmd_calc_Click()
''    Screen.MousePointer = 11
''    zzTeste.Show 0
''
''    Exit Sub
''
''
'    Dim retval As Long
'    retval = Shell("calc", vbNormalFocus)
'End Sub
'Private Sub cmd_calendario_Click()
'    Screen.MousePointer = 11
'    cerrado_calendario.Show 1
'    g_string = ""
'End Sub
Private Sub cmd_configuracao_Click()
    Screen.MousePointer = 11
    config_liberacao_caixa.Show 1
End Sub
Private Sub cmd_senha_Click()
    If FormulariosFechados = False Then
        MsgBox "O usuário não pode ser trocado com alguma tela aberta.", vbCritical, "Outra tela aberta!"
        Exit Sub
    End If
    LimpaMenu
    g_nivel_acesso = 0
    cmdTransfereDadosLMC.Visible = False
    DesativaVerificacaoGIC
    gInternetBandaLarga = False
    
    If lResgateAbertoPeloSGP Then
       FinalizaPrograma ("NFeResgaste.exe")
    End If
    
    frm_identificacao.Show 1
    BuscaConfiguracao
    If g_nome_usuario = "L.M.C." Then
        cmdTransfereDadosLMC.Visible = True
    End If
    'Me.Caption = "Sistema Gerenciador de Posto - " & UCase(Mid(g_nome_usuario, 1, 1)) & LCase(Mid(g_nome_usuario, 2, Len(g_nome_usuario) - 1))
    MontaMenu
    MudaEmpresaVbNet2003
    MudaEmpresaVbNet2008
    StatusBar1.Panels(2).Text = g_nome_usuario
    'BuscaDados
End Sub
Private Sub cmd_sql_Click()
    'Screen.MousePointer = 11
    Call GravaSgpCadastroIni("ExportaImportaDados")
    'exporta_importa.Show
End Sub
Private Sub BuscaRegistroLiberacaoDigitacao()
    Dim xTipoLiberacao As Integer
    Dim xTipoVenda As String
    
    xTipoLiberacao = 1
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CONVENIENCIA" Or xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Then
        xTipoLiberacao = 3
    Else
        If ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni) = "SIM" Then
            xTipoLiberacao = 2
        End If
    End If
    
    
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, xTipoLiberacao) Then
        g_cfg_empresa_i = LiberacaoDigitacao.Empresa
        g_cfg_empresa_f = LiberacaoDigitacao.Empresa
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
    Else
        LiberacaoDigitacao.Empresa = g_empresa
        LiberacaoDigitacao.DataInicial = Date
        LiberacaoDigitacao.DataFinal = Date
        LiberacaoDigitacao.PeriodoInicial = 1
        LiberacaoDigitacao.PeriodoFinal = 1
        LiberacaoDigitacao.TipoMovimento = xTipoLiberacao
        If Not LiberacaoDigitacao.Incluir Then
            MsgBox "Não foi possível incluir o registro de liberação de digitação!", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub Form_Activate()
    Dim xUsuario As Integer
    If g_nivel_acesso <= 4 Then
        cmd_configuracao.Enabled = True
    Else
        cmd_configuracao.Enabled = False
    End If
    If g_nivel_acesso = 0 Then
        xUsuario = g_usuario
        'a linha de comando abaixo he uma tentativa de passar
        'o foco para este programa
        'e nao para o programa .net
        Me.Show
        frm_identificacao.Show 1
        If g_nome_usuario = "L.M.C." Then
            cmdTransfereDadosLMC.Visible = True
        End If
        If g_nivel_acesso <> 0 Then
            'Me.Caption = "Sistema Gerenciador de Posto - " & UCase(Mid(g_nome_usuario, 1, 1)) & LCase(Mid(g_nome_usuario, 2, Len(g_nome_usuario) - 1))
            MontaMenu
            BuscaDados
            BuscaRegistroLiberacaoDigitacao
            BuscaConfiguracao
            timerVbNet.Enabled = True
            timerVbNet.Interval = 500
            If xUsuario = 0 Then
                Call ClickMenu("CA", "Teste")
            End If
        Else
            Finaliza
        End If
        StatusBar1.Panels(2).Text = g_nome_usuario
        StatusBar1.Panels(2).AutoSize = sbrContents
    End If
    
    'Teste quando perdia conexão ao passar cartao pela segunda vez
    'If adodc_empresa.Recordset.State = 0 Then
    '    PreencheCboEmpresa
    '    dtcbo_empresa.BoundText = ""
    'End If
    
    If dtcbo_empresa.BoundText = "" Then
        dtcbo_empresa.BoundText = g_empresa
        g_nome_empresa = dtcbo_empresa.Text
        If Empresa.LocalizarCodigo(g_empresa) Then
            g_cidade_empresa = Trim(Empresa.Cidade)
            gUfEmpresa = Trim(Empresa.Estado)
            gEmpresaGlobal = Empresa.EmpresaGlobal
            gCNPJEmpresa = Empresa.CGC 'ALEX - NFCE
        Else
            g_cidade_empresa = "Goiânia"
            gUfEmpresa = "GO"
        End If
    End If
    VerificaEcfComEmpresa
    If dtcbo_empresa.Enabled Then
        dtcbo_empresa.SetFocus
    End If
   
    
End Sub
Private Sub Form_Load()
    Dim retval As Long
    
    
    'Versão para SGP  = 7
    'Sub-Versao (Mes) = 8
    'Sub-versao (Dia) = 9
    'Correcao do dia  = 2
    'gVersaoSGP = "12.01.09a"
    
    'Ano = 19
    'Sub-Versao Mes   = 01
    'Sub-versao Dia   = 22
    'Correcao do dia  = a
    gVersaoSGP = "19.08.02a"

    gTipoAmbienteNFCe = ""
    Me.Caption = "Sistema Gerênciador de Postos de Combustíveis - Versão " & gVersaoSGP
    gArquivoIni = "C:\Cerrado.Net\Sgp.INI"
    VerificaPastaCerradoNet
    If App.PrevInstance Then
        MsgBox "Atenção! Este Programa já encontra-se aberto!", vbExclamation, "SGP está aberto!"
        End
    End If
    
    Me.Top = 0
    Me.Left = 0
    gEmpresaGlobal = 0
    gUsuarioGlobal = 0
    lEmpresaAtualPetromovel = 0
'    lContadorTimer = 600
'    lContadorTimer2 = 0
'    Me.Width = Screen.Width
'    Me.Height = Screen.Height - 300
    
    
    'CentraForm Me
    'ChDrive "C"
    'If CDate(Date) >= CDate("30/04/1999") Then
    '    MsgBox Error(71), vbCritical, "Erro Grave"
    '    End
    'End If
    
    ChamaDrive
    'lNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    lNomeArquivo = "C:" & gDiretorioAplicativo & "sgp_cadastro.ini"
    If gArqTxt.FileExists(lNomeArquivo) Then
        Call gArqTxt.DeleteFile(lNomeArquivo, True)
    End If
    If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", lNomeArquivo)
    End If
    If gArqTxt.FolderExists("C:\Cerrado.Net\SgpNet") Then
        Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini")
    End If
    splash.Show 1
    If Not ChamaDrive Then
        If Not VerificaCriaConexao Then
            End
        Else
            If Not ChamaDrive() Then
                End
            End If
        End If
    End If
    If gArqTxt.FileExists(lNomeArquivo) Then
        Call gArqTxt.DeleteFile(lNomeArquivo, True)
    End If
    If gArqTxt.FileExists("C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini") Then
        Call gArqTxt.DeleteFile("C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini", True)
    End If
    If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
        retval = Shell("C:\Cerrado.Net\sgp\bin\SGP_CADASTRO.exe", vbMinimizedNoFocus)
    End If
    
    If gArqTxt.FileExists("C:\Cerrado.Net\SgpNet\bin\Release\SgpNet.exe") Then
        retval = Shell("C:\Cerrado.Net\SgpNet\bin\Release\SgpNet.exe", vbMinimizedNoFocus)
    End If
    
    Me.Show
    
    If bdSqlServer Then
        StatusBar1.Panels(1).Text = gIpBanco
        If gNomeInternoBD <> "sgp_data" Then
            StatusBar1.Panels(1).Text = gIpBanco & gNomeInternoBD
        End If
        StatusBar1.Panels(1).AutoSize = sbrContents
        
        
'        'Teste de criacao de dsn automatico
'        Dim xDSN As New cDSN
'        MsgBox xDSN.DSNDelete("sgp_data", "SQL Server", False)
'        MsgBox xDSN.DNSCria(1, gIpBanco & ",4949", "sgp_data", "sgp_data", "SQL Server", "sgp_data", "sa", gSenhaBD , False, True)
'        Set xDSN = Nothing
''        If CriaDSN() Then
''            MsgBox "ok"
''        Else
''            MsgBox "erro"
''        End If
    
    
    Else
        StatusBar1.Panels(1).Text = gDrive
        StatusBar1.Panels(1).AutoSize = sbrContents
    End If

'    Call ChamaDrive
'    ChDir "\VB5\SGP\DATA"
'    Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
'    Set bd_sgp_b = OpenDatabase("SGP_DATA_BAIXA.MDB")
'    Set bd_sgp_m = OpenDatabase("SGP_DATA_MOVIMENTO.MDB")
    
    
    Call GravaAuditoria(1, Me.name, 32, "Sgp.exe " & gVersaoSGP)
    
    
    'adodc_empresa.ConnectionString = Conectar.ConnectionString
    'Set adodc_empresa.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Empresas WHERE Inativo = " & 0 & " ORDER BY Codigo")
    PreencheCboEmpresa
    
'    adodc_empresa.ConnectionString = gConnectionString
'    adodc_empresa.RecordSource = "SELECT Codigo, Nome FROM Empresas WHERE Inativo = FALSE ORDER BY Codigo"
'    adodc_empresa.Refresh
    
    g_data_def = Date
    g_lmc = 0
    g_impressora_matricial = False
    lArquivoVb6VbNet = "C:" & gDiretorioAplicativo & "VB.NET_VB6.INI"
    lArquivoVb6VbNet2 = "C:" & gDiretorioAplicativo & "VB.NET_VB62.INI"
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub mnu_cadastro_item_Click(Index As Integer)
    Dim i As Integer
    Dim xNomeMenu As String
    If Index = l_sair Then
        If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 6 Then
            Finaliza
        Else
            Exit Sub
        End If
    End If
    i = Len(Trim(mnu_cadastro_item.Item(Index).Caption))
    If mnu_cadastro_item.Item(Index).Caption <> "" Then
        xNomeMenu = Mid(mnu_cadastro_item.Item(Index).Caption, 2, i - 1)
        Call ClickMenu("CA", xNomeMenu)
    End If
End Sub
Private Sub mnu_consulta_item_Click(Index As Integer)
    Dim i As Integer
    Dim xNomeMenu As String
    i = Len(Trim(mnu_consulta_item.Item(Index).Caption))
    If mnu_consulta_item.Item(Index).Caption <> "" Then
        xNomeMenu = Mid(mnu_consulta_item.Item(Index).Caption, 2, i - 1)
        Call ClickMenu("CO", xNomeMenu)
    End If
End Sub
Private Sub mnu_grafico_item_Click(Index As Integer)
    Dim i As Integer
    Dim xNomeMenu As String
    i = Len(Trim(mnu_grafico_item.Item(Index).Caption))
    If mnu_grafico_item.Item(Index).Caption <> "" Then
        xNomeMenu = Mid(mnu_grafico_item.Item(Index).Caption, 2, i - 1)
        Call ClickMenu("GR", xNomeMenu)
    End If
End Sub

Private Sub mnu_movimentacao_item_Click(Index As Integer)
    Dim i As Integer
    Dim xNomeMenu As String
    i = Len(Trim(mnu_movimentacao_item.Item(Index).Caption))
    If mnu_movimentacao_item.Item(Index).Caption <> "" Then
        xNomeMenu = Mid(mnu_movimentacao_item.Item(Index).Caption, 2, i - 1)
        Call ClickMenu("MO", xNomeMenu)
    End If
End Sub

Private Sub mnu_relatorio_item_Click(Index As Integer)
    Dim i As Integer
    Dim xNomeMenu As String
    i = Len(Trim(mnu_relatorio_item.Item(Index).Caption))
    If mnu_relatorio_item.Item(Index).Caption <> "" Then
        xNomeMenu = Mid(mnu_relatorio_item.Item(Index).Caption, 2, i - 1)
        Call ClickMenu("RE", xNomeMenu)
    End If
End Sub
Private Sub mnu_sobre_Click()
    If g_lmc = 1 Then
        g_lmc = 3
    Else
        g_lmc = g_lmc + 1
    End If
    Screen.MousePointer = 11
    frm_sobre.Show 1
End Sub
Private Sub timerVbNet_Timer()
'    lContadorTimer = lContadorTimer + 1
'    lContadorTimer2 = lContadorTimer2 + 1
    If gArqTxt.FileExists(lArquivoVb6VbNet) Then
        Dim xNomePrograma As String
        xNomePrograma = ReadINI("CAIXA", "Nome do Programa", lArquivoVb6VbNet)
        g_string = ReadINI("CAIXA", "dados", lArquivoVb6VbNet)
        Call gArqTxt.DeleteFile(lArquivoVb6VbNet, True)
        If Programa.LocalizarNomeInterno(xNomePrograma) Then
            Call gArqTxt.CreateTextFile("C:" & gDiretorioAplicativo & "Retorno_VB6.TMP", True)
            Call gArqTxt.MoveFile("C:" & gDiretorioAplicativo & "Retorno_VB6.TMP", "C:" & gDiretorioAplicativo & "Retorno_VB6.INI")
            If RetiraGString(2) = "Imprimir" Or RetiraGString(2) = "Visualizar" Then
                Call ClickMenu("RE", Programa.NomeparaMenu)
            ElseIf RetiraGString(2) = "Baixar" Then
                Call ClickMenu("MO", Programa.NomeparaMenu)
            Else
                Call ClickMenu("MO", Programa.NomeparaMenu)
            End If
        Else
            If xNomePrograma = "TRAVA_SISTEMA" Then
                Call GravaAuditoria(1, Me.name, 27, "Locação Vencida")
                Dim xMensagem As String
                xMensagem = "A versão do executável do SGP está incompatível com o banco de dados atual." & Chr(13) & Chr(13)
                xMensagem = xMensagem & "O sistema não irá funcionar até que todos os computadores estejam com a versão atual do SGP." & Chr(13) & Chr(13)
                xMensagem = xMensagem & "Este procedimento irá impedir que seja danificada a integridade relacional do banco de dados." & Chr(13)
                xMensagem = xMensagem & "Entre em contato URGENTEMENTE com o suporte técnico para que seja atualizado o sistema." & Chr(13)
                xMensagem = xMensagem & "Sn. 2000-0037-12-01-" & gNumeroHd & Chr(13) & Chr(13)
                xMensagem = xMensagem & "Telefone do Suporte: (62) 3277-1017" & Chr(13) & Chr(13)
                xMensagem = xMensagem & "Cerrado Tecnologia - Soluções Inteligentes."
                MsgBox xMensagem, vbCritical, "Versão Incompatível com o Banco de Dados."
                Finaliza
            ElseIf xNomePrograma = "SolicitacaoSangria" Then
                gStringChamadaSangria = g_string
                g_string = ""
           ElseIf xNomePrograma = "ImprimeTransacaoLio" Then
                frm_preview.Show 1

            'ElseIf xNomePrograma = "VerificaTipoAmbiente" Then
            '    'gStringChamadaSangria = g_string
            '    MsgBox "Tipo Ambiente nfce = " & g_string
            '    g_string = ""
            Else
                MsgBox "Erro na ligação VB.Net para VB-6!" & "[" & xNomePrograma & "]", vbOKOnly + vbExclamation, "Programa Inexistente!"
            End If
        End If
    End If
    If gArqTxt.FileExists(lArquivoVb6VbNet2) Then
        Dim xNomePrograma2 As String
        xNomePrograma2 = ReadINI("CAIXA", "Nome do Programa", lArquivoVb6VbNet2)
        g_string = ReadINI("CAIXA", "dados", lArquivoVb6VbNet2)
        Call gArqTxt.DeleteFile(lArquivoVb6VbNet2, True)
        If xNomePrograma2 = "VerificaTipoAmbiente" Then
            
            gTipoAmbienteNFCe = g_string
            g_string = ""
            If gTipoAmbienteNFCe <> "1" Then
                If (MsgBox("Você esta no ambiente de Homologação, Deseja sair do sistema ?", vbYesNo + vbQuestion + vbDefaultButton1, "Ambiente de HOMOLOGAÇÃO!")) = vbYes Then
                    End
                End If
            End If
        ElseIf xNomePrograma2 = "ImprimeNotaAbastecimento" Then
           Call DefineImpressoraTermicaComoPadrao
           frm_preview.Show 1
        Else
            MsgBox "Erro na ligação VB.Net para VB-6 2!", vbOKOnly + vbExclamation, "Programa Inexistente!"
        End If
    End If
    'Cada 120 equivale a 1 Minuto
    '600 = 5 Minutos
'    If lContadorTimer >= 600 Then
'        lContadorTimer = 0
'        If lNotificacaoGic Then
'            If VerificaConexaoInternet Then
'                If gInternetBandaLarga And gEmpresaGlobal > 0 Then
'                    BuscaDadosGIC
'                End If
'            End If
'        End If
'    End If
    
    'Cada 120 equivale a 1 Minuto
    '3600 = 30 Minutos
'    If lContadorTimer2 = 3600 Then
'        lContadorTimer2 = 0
'        If lNotificacaoGic Then
'            If gInternetBandaLarga And gEmpresaGlobal = 0 Then
'                lContadorTimer = 0
'                AtivaVerificacaoGIC
'            End If
'        End If
'    End If
End Sub
Private Sub PreencheCboEmpresa()
    Set adodc_empresa.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Empresas WHERE Inativo = " & 0 & " ORDER BY Codigo")
End Sub

Private Function VerificaConexaoInternet() As Boolean
    Dim xHoraInicial As Date
    
    On Error GoTo FileError
    
    VerificaConexaoInternet = False
    Winsock1.Close
    If ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni) = "TEIXEIRA E PINHEIRO LTDA" Then
        Winsock1.RemoteHost = "192.168.1.6"
    Else
        Winsock1.RemoteHost = "tasso.myvnc.com"
    End If
    Winsock1.RemotePort = 80
    'Winsock1.RemotePort = 4949
    Winsock1.RemotePort = Mid(gPortaBanco, 2, 4)
    Winsock1.Connect
    
    'Aguarda 2 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 2
        DoEvents
    Loop
    
    If Winsock1.State = 6 Or Winsock1.State = 7 Then
        VerificaConexaoInternet = True
    Else
        CriaLogSGP "VerificaConexaoInternet: Não Conectado:" & Winsock1.State, "", ""
    End If
    Winsock1.Close
    Exit Function

FileError:
    CriaLogSGP "VerificaConexaoInternet: Erro ao verificar conexão com internet", Error, ""
End Function
Private Function VerificaCriaConexao() As Boolean
    Dim xMapeamentoRede As New cMapeamentoRede
    Dim i As Integer
    Dim xIpServidor As String
    
    
    VerificaCriaConexao = False
    gDrive = ReadINI("LOCAL", "Drive", gArquivoIni)
    gNomeBancoDados = ReadINI("LOCAL", "Nome do Banco de Dados", gArquivoIni)
    
    If Not IsNumeric(Mid(gNomeBancoDados, 1, 1)) Then
        VerificaCriaConexao = True
        Exit Function
    End If
    If UCase(gDrive) = "C:" Or Mid(gNomeBancoDados, 1, 9) = "127.0.0.1" Then
        VerificaCriaConexao = True
        Exit Function
    End If
    
    xIpServidor = "\\"
    For i = 1 To Len(gNomeBancoDados)
        If Mid(gNomeBancoDados, i, 1) = "," Then
            Exit For
        End If
        xIpServidor = xIpServidor & Mid(gNomeBancoDados, i, 1)
    Next
    xIpServidor = xIpServidor & "\C"
    If xMapeamentoRede.ConsultaConexao(gDrive) Then
        VerificaCriaConexao = True
    Else
        If xMapeamentoRede.CriaConexao(xIpServidor, "", gDrive) Then
            VerificaCriaConexao = True
        Else
            MsgBox "Não foi possível estabelecer uma conexão com o servidor do banco de dados." & vbCrLf & "Por esse motivo não será possível entrar no sistema." & vbCrLf & "Verifique as conexões de rede e/ou se o servidor esteja ligado." & vbCrLf & "Caso já tenha verificado, entre em contato com o suporte técnico.", vbInformation, "Erro de Comunicacao com o Servidor!"
        End If
    End If
'    If Not xMapeamentoRede.Desconecta(gDrive) Then
'        MsgBox "Erro ao desfazer a conexão!", vbInformation, "Erro de Conexao!"
'    End If
End Function
Private Sub VerificaEcfComEmpresa()
    Dim xNomeEmpresa As String
    Dim xCupomDemonstracao As Boolean
    Dim i As Integer
    
    xCupomDemonstracao = False
    If ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni) = "NAO" Then
        Exit Sub
    End If
    If ReadINI("CUPOM FISCAL", "Cupom Demonstracao", gArquivoIni) = "SIM" Then
        xCupomDemonstracao = True
    End If

    xNomeEmpresa = ReadINI("CUPOM FISCAL", "Nome da Empresa", gArquivoIni)
    If dtcbo_empresa.Text <> xNomeEmpresa Then
        adodc_empresa.Recordset.MoveFirst
        Do Until adodc_empresa.Recordset.EOF
            dtcbo_empresa.BoundText = adodc_empresa.Recordset!Codigo
            If dtcbo_empresa.Text = xNomeEmpresa Then
                Exit Do
            End If
            adodc_empresa.Recordset.MoveNext
        Loop
    End If
    
    If dtcbo_empresa.Text <> xNomeEmpresa Then
        If xCupomDemonstracao = False Then
            MsgBox "Empresa não configurada para usar cupom fiscal." & vbCrLf & "Entre em contato com o suporte técnico.", vbCritical, "Empresa Não Configurada!"
            Finaliza
        End If
    End If
End Sub
Private Sub VerificaPastaCerradoNet()
'    gArquivoIni = "C:\Cerrado.Net\Sgp.INI"
    If Not gArqTxt.FileExists(gArquivoIni) Then
        Call gArqTxt.MoveFile("C:\Sgp.Ini", gArquivoIni)
        Call gArqTxt.MoveFile("C:\Sgp_data.mdb", "C:\Cerrado.Net\")
        If gArqTxt.FileExists("C:\AtualizaSgpInternet.bat") Then
            Call gArqTxt.MoveFile("C:\AtualizaSgpInternet.bat", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\AtualizaSgpRede.bat") Then
            Call gArqTxt.MoveFile("C:\AtualizaSgpRede.bat", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\AtualizaSpedFiscal.bat") Then
            Call gArqTxt.MoveFile("C:\AtualizaSpedFiscal.bat", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\AtualizaNFeInternet.bat") Then
            Call gArqTxt.MoveFile("C:\AtualizaNFeInternet.bat", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\CheqPosto.ini") Then
            Call gArqTxt.MoveFile("C:\CheqPosto.ini", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\EnviaEmail.ini") Then
            Call gArqTxt.MoveFile("C:\EnviaEmail.ini", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\Limpatef.bat") Then
            Call gArqTxt.MoveFile("C:\Limpatef.bat", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\SuporteCerrado.exe") Then
            Call gArqTxt.MoveFile("C:\SuporteCerrado.exe", "C:\Cerrado.Net\")
        End If
        If gArqTxt.FileExists("C:\SuporteRemotoCerrado.exe") Then
            Call gArqTxt.MoveFile("C:\SuporteRemotoCerrado.exe", "C:\Cerrado.Net\")
        End If
    End If
End Sub

