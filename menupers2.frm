VERSION 5.00
Begin VB.Form menu_personalizado2 
   Caption         =   "Cupom Fiscal Cerrado"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9120
   Icon            =   "menupers2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "menupers2.frx":0442
   ScaleHeight     =   960
   ScaleWidth      =   9120
   Begin VB.CommandButton cmd_sql 
      Enabled         =   0   'False
      Height          =   555
      Left            =   6360
      Picture         =   "menupers2.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pesquisa geral."
      Top             =   360
      Width           =   795
   End
   Begin VB.CommandButton cmd_configuracao 
      Caption         =   "&Liberação"
      Height          =   915
      Left            =   7260
      Picture         =   "menupers2.frx":1B62
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Muda senha."
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmd_calc 
      Height          =   375
      Left            =   8220
      Picture         =   "menupers2.frx":2E3C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Calculadora."
      Top             =   540
      Width           =   795
   End
   Begin VB.CommandButton cmd_calendario 
      Height          =   555
      Left            =   8220
      Picture         =   "menupers2.frx":4116
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Calendário."
      Top             =   0
      Width           =   795
   End
   Begin VB.Data dta_empresa 
      Caption         =   "dta_empresa"
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Empresas"
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cbo_conta_bancaria 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.ComboBox cbo_empresa 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Conta Bancária"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label7 
      Caption         =   "Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
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
Attribute VB_Name = "menu_personalizado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl_dados As Table
Dim tbl_empresa As Table
Dim tbl_configuracao As Table
Dim tbl_conta_bancaria As Table
Dim tbl_liberacao_digitacao As Table
Dim tbl_menu As Table
Dim l_sair As Integer
Dim i_ca As Integer
Dim i_co As Integer
Dim i_gr As Integer
Dim i_mo As Integer
Dim i_re As Integer
Private Sub BuscaConfiguracao()
    g_caixa_unificado = False
    If tbl_configuracao.RecordCount > 0 Then
        tbl_configuracao.Seek "=", g_empresa
        If Not tbl_configuracao.NoMatch Then
            If Mid(tbl_configuracao![Outras Configuracoes], 1, 1) = "S" Then
                g_caixa_unificado = True
            Else
                g_caixa_unificado = False
            End If
        Else
            MsgBox "Não existe configuração para esta empresa.", 64, "Ajustar Configuração do Sistema!"
        End If
    End If
End Sub
Private Sub BuscaDados()
    If tbl_dados.RecordCount = 0 Then
        tbl_dados.AddNew
        tbl_dados!Codigo = 1
        tbl_dados!Empresa = 1
        tbl_dados![Conta Bancaria] = 1
        tbl_dados![Empresa 2] = 0
        g_empresa = tbl_dados!Empresa
        tbl_dados.Update
    End If
    tbl_dados.MoveFirst
    If tbl_dados![Empresa 2] = 0 Then
        If CDate(Date) >= CDate("10/04/2000") Then
            tbl_dados.Edit
            tbl_dados![Empresa 2] = 1
            tbl_dados.Update
            tbl_dados.MoveFirst
        End If
    End If
    If tbl_dados![Empresa 2] <> 0 Then
        Printer.PaperSize = 300
    End If
    g_empresa = tbl_dados!Empresa
    g_conta_bancaria = tbl_dados![Conta Bancaria]
End Sub
Private Sub Chama(x_interno As String)
    'BuscaNumeroDeSerie
    x_interno = Trim(x_interno)
'Cadastro
    Screen.MousePointer = 11
    If x_interno = "cadastro_aliquota" Then
        cadastro_aliquota.Show
    'ElseIf x_interno = "cadastro_banco" Then
    '    cadastro_banco.Show
    ElseIf x_interno = "cadastro_bomba" Then
        cadastro_bomba.Show
    'ElseIf x_interno = "cadastro_cartao" Then
    '    cadastro_cartao.Show
    ElseIf x_interno = "cadastro_cliente" Then
        cadastro_cliente.Show
    ElseIf x_interno = "cadastro_cliente_conveniado" Then
        cadastro_cliente_conveniado.Show
    ElseIf x_interno = "cadastro_combustivel" Then
        cadastro_combustivel.Show
    ElseIf x_interno = "cadastro_configuracao" Then
        cadastro_configuracao.Show
    'ElseIf x_interno = "cadastro_conta_bancaria" Then
    '    cadastro_conta_bancaria.Show
    ElseIf x_interno = "cadastro_convenio" Then
        cadastro_convenio.Show
    'ElseIf x_interno = "cadastro_dependente" Then
    '    cadastro_dependente.Show
    ElseIf x_interno = "cadastro_empresa" Then
        cadastro_empresa.Show
    'ElseIf x_interno = "cadastro_fornecedor" Then
    '    cadastro_fornecedor.Show
    ElseIf x_interno = "cadastro_funcionario" Then
        cadastro_funcionario.Show
    'ElseIf x_interno = "cadastro_grau_dependencia" Then
    '    cadastro_grau_dependencia.Show
    ElseIf x_interno = "cadastro_grupo" Then
        cadastro_grupo.Show
    'ElseIf x_interno = "cadastro_historico" Then
    '    cadastro_historico.Show
    ElseIf x_interno = "cadastro_menu" Then
        cadastro_menu.Show
    ElseIf x_interno = "cadastro_produto" Then
        cadastro_produto.Show
    ElseIf x_interno = "cadastro_programa" Then
        cadastro_programa.Show
    'ElseIf x_interno = "cadastro_tabela_folha" Then
    '    cadastro_tabela_folha.Show
    'ElseIf x_interno = "cadastro_tabela_premiacao" Then
    '    cadastro_tabela_premiacao.Show
    'ElseIf x_interno = "cadastro_tabela_provento_desconto" Then
    '    cadastro_tabela_provento_desconto.Show
    ElseIf x_interno = "cadastro_tabela_vencimento" Then
        cadastro_tabela_vencimento.Show
    'ElseIf x_interno = "cadastro_tipo_documento" Then
    '    cadastro_tipo_documento.Show
    ElseIf x_interno = "cadastro_usuario" Then
        cadastro_usuario.Show
'Cadastro Conversao
'Consulta
    ElseIf x_interno = "cerrado_calendario" Then
        cerrado_calendario.Show
    'ElseIf x_interno = "consulta_lmc" Then
    '    consulta_lmc.Show
    ElseIf x_interno = "consulta_nota_cliente" Then
        consulta_nota_cliente.Show
    ElseIf x_interno = "consulta_movimento_cupom_fiscal" Then
        consulta_movimento_cupom_fiscal.Show
    'ElseIf x_interno = "consulta_nota_conveniado" Then
    '    consulta_nota_conveniado.Show
    'ElseIf x_interno = "consulta_quadro_funcionario" Then
    '    consulta_quadro_funcionario.Show
    'ElseIf x_interno = "super_consulta" Then
    '    super_consulta.Show
'Gráficos
    'ElseIf x_interno = "grafico_despesa_anual" Then
    '    grafico_despesa_anual.Show
    'ElseIf x_interno = "grafico_despesa_mensal" Then
    '    grafico_despesa_mensal.Show
    'ElseIf x_interno = "grafico_venda_combustivel_anual" Then
    '    grafico_venda_combustivel_anual.Show
    'ElseIf x_interno = "grafico_venda_combustivel_mensal" Then
    '    grafico_venda_combustivel_mensal.Show
'Movimento Baixa
    'ElseIf x_interno = "baixa_cheque" Then
    '    baixa_cheque.Show
    'ElseIf x_interno = "baixa_cheque_devolvido_descontado" Then
    '    baixa_cheque_devolvido_descontado.Show
    'ElseIf x_interno = "baixa_contas_pagar" Then
    '    baixa_contas_pagar.Show
    'ElseIf x_interno = "baixa_duplicata_receber" Then
    '    baixa_duplicata_receber.Show
    ElseIf x_interno = "baixa_nota_abastecimento" Then
        baixa_nota_abastecimento.Show
    'ElseIf x_interno = "baixa_nota_abastecimento_periodo" Then
    '    baixa_nota_abastecimento_periodo.Show
'Movimento
    'ElseIf x_interno = "gera_disquete_deposito" Then
    '    gera_disquete_deposito.Show
    'ElseIf x_interno = "movimento_bancario" Then
    '    movimento_bancario.Show
    ElseIf x_interno = "movimento_bomba" Then
        movimento_bomba.Show
    'ElseIf x_interno = "movimento_cheque" Then
    '    movimento_cheque.Show
    'ElseIf x_interno = "movimento_cheque_avista" Then
    '    movimento_cheque_avista.Show
    'ElseIf x_interno = "movimento_cheque_devolvido" Then
    '    movimento_cheque_devolvido.Show
    'ElseIf x_interno = "movimento_cheque_devolvido_baixado" Then
    '    movimento_cheque_devolvido_baixado.Show
    ' ElseIf x_interno = "movimento_cheque_extraviado" Then
    '     movimento_cheque_extraviado.Show
    ElseIf x_interno = "movimento_entrada_produto" Then
        movimento_entrada_produto.Show
    'ElseIf x_interno = "movimento_falta_caixa" Then
    '    movimento_falta_caixa.Show
    'ElseIf x_interno = "mov_contas_pagar" Then
    '    mov_contas_pagar.Show
    'ElseIf x_interno = "movimento_duplicata_receber" Then
    '    movimento_duplicata_receber.Show
    'ElseIf x_interno = "mov_entrada_combustiveis" Then
    '    mov_entrada_combustiveis.Show
    'ElseIf x_interno = "mov_medicao_combustiveis" Then
    '    mov_medicao_combustiveis.Show
    ElseIf x_interno = "mov_nota_abastecimento" Then
        mov_nota_abastecimento.Show
    'ElseIf x_interno = "movimento_advertencia_suspencao" Then
    '    movimento_advertencia_suspencao.Show
    'ElseIf x_interno = "movimento_afericao" Then
    '    movimento_afericao.Show
    'ElseIf x_interno = "movimento_caixa" Then
    '    movimento_caixa.Show
    'ElseIf x_interno = "movimento_cartao_credito" Then
    '    movimento_cartao_credito.Show
    'ElseIf x_interno = "movimento_cheque_cobranca" Then
    '    movimento_cheque_cobranca.Show
    'ElseIf x_interno = "movimento_cupom_fiscal" Then
    '    movimento_cupom_fiscal.Show
    ElseIf x_interno = "movimento_cupom_fiscal2" Then
        movimento_cupom_fiscal2.Show
    'ElseIf x_interno = "movimento_falta_funcionario" Then
    '    movimento_falta_funcionario.Show
    'ElseIf x_interno = "movimento_folha" Then
    '    movimento_folha.Show
    ElseIf x_interno = "movimento_historico" Then
        movimento_historico.Show
    ' ElseIf x_interno = "movimento_leasing_veiculo" Then
    '     movimento_leasing_veiculo.Show
    ElseIf x_interno = "movimento_oleo_diverso" Then
        movimento_oleo_diverso.Show
    'ElseIf x_interno = "movimento_pedido_combustivel" Then
    '    movimento_pedido_combustivel.Show
    'ElseIf x_interno = "movimento_saida_transferencia_produto" Then
    '    movimento_saida_transferencia_produto.Show
    'ElseIf x_interno = "processamento_estoque" Then
    '    processamento_estoque.Show
'Relatórios
    'ElseIf x_interno = "emissao_advertencia" Then
    '    emissao_advertencia.Show
    'ElseIf x_interno = "emissao_analise_geral" Then
    '    emissao_analise_geral.Show
    'ElseIf x_interno = "emissao_analise_giro_estoque" Then
    '    emissao_analise_giro_estoque.Show
    'ElseIf x_interno = "emissao_analise_inventario" Then
    '    emissao_analise_inventario.Show
    'ElseIf x_interno = "emissao_analise_movimentacao_postos" Then
    '    emissao_analise_movimentacao_postos.Show
    'ElseIf x_interno = "emissao_analise_venda_cartao" Then
    '    emissao_analise_venda_cartao.Show
    'ElseIf x_interno = "emissao_analise_vendas_funcionarios" Then
    '    emissao_analise_vendas_funcionarios.Show
    'ElseIf x_interno = "emissao_analise_venda_produto" Then
    '    emissao_analise_venda_produto.Show
    'ElseIf x_interno = "emissao_balanco" Then
    '    emissao_balanco.Show
    'ElseIf x_interno = "emissao_conta_pagar_conferencia" Then
    '    emissao_conta_pagar_conferencia.Show
    ElseIf x_interno = "emissao_cupom_complementar" Then
        emissao_cupom_complementar.Show
    ElseIf x_interno = "emissao_memoria_fiscal" Then
        emissao_memoria_fiscal.Show
    'ElseIf x_interno = "emissao_grps" Then
    '    emissao_grps.Show
    'ElseIf x_interno = "emissao_kit_documento_funcionario" Then
    '    emissao_kit_documento_funcionario.Show
    'ElseIf x_interno = "emissao_recibo_folha_pagamento" Then
    '    emissao_recibo_folha_pagamento.Show
    'ElseIf x_interno = "emissao_resumo_folha_pagamento" Then
    '    emissao_resumo_folha_pagamento.Show
    'ElseIf x_interno = "emissao_resumo_movimentacao_postos" Then
    '    emissao_resumo_movimentacao_postos.Show
    'ElseIf x_interno = "emissao_rpa" Then
    '    emissao_rpa.Show
    'ElseIf x_interno = "emissao_cesta_basica" Then
    '    emissao_cesta_basica.Show
    'ElseIf x_interno = "emissao_extrato_bancario" Then
    '    emissao_extrato_bancario.Show
    'ElseIf x_interno = "emissao_funcionario" Then
    '    emissao_funcionario.Show
    'ElseIf x_interno = "emissao_funcionario_ficha" Then
    '    emissao_funcionario_ficha.Show
    ElseIf x_interno = "emissao_movimento_bomba" Then
        emissao_movimento_bomba.Show
    'ElseIf x_interno = "emissao_nota_cliente" Then
    '    emissao_nota_cliente.Show
    'ElseIf x_interno = "emissao_recibo" Then
    '    emissao_recibo.Show
    'ElseIf x_interno = "emissao_suspencao" Then
    '    emissao_suspencao.Show
    'ElseIf x_interno = "emissao_vale_transporte" Then
    '    emissao_vale_transporte.Show
    'ElseIf x_interno = "emissao_movimento_digitacao" Then
    '    emissao_movimento_digitacao.Show
    'ElseIf x_interno = "emissao_movimento_lubrificante" Then
    '    emissao_movimento_lubrificante.Show
    'ElseIf x_interno = "emissao_venda_cliente" Then
    '    emissao_venda_cliente.Show
    'ElseIf x_interno = "frm_emissao_cheques_folhas" Then
    '    frm_emissao_cheques_folhas.Show
    'ElseIf x_interno = "emissao_cheque_formulario" Then
    '    emissao_cheque_formulario.Show
    'ElseIf x_interno = "emissao_preco_combustivel" Then
    '    emissao_preco_combustivel.Show
    'ElseIf x_interno = "frm_emissao_lmc" Then
    '    frm_emissao_lmc.Show
    'ElseIf x_interno = "frm_emissao_recibo_folhas" Then
    '    frm_emissao_recibo_folhas.Show
    'ElseIf x_interno = "frm_emissao_recibo_formulario" Then
    '    frm_emissao_recibo_formulario.Show
    'ElseIf x_interno = "listagem_cheque_formulario" Then
    '    listagem_cheque_formulario.Show
    'ElseIf x_interno = "lst_baixa_cheque_devolvido" Then
    '    lst_baixa_cheque_devolvido.Show
    'ElseIf x_interno = "lst_baixa_cheque_devolvido_descontado" Then
    '    lst_baixa_cheque_devolvido_descontado.Show
    'ElseIf x_interno = "lst_baixa_contas_a_pagar_fornecedor" Then
    '    lst_baixa_contas_a_pagar_fornecedor.Show
    'ElseIf x_interno = "lst_baixa_pagar" Then
    '    lst_baixa_pagar.Show
    'ElseIf x_interno = "lst_bordero_deposito" Then
    '    lst_bordero_deposito.Show
    'ElseIf x_interno = "lst_bordero_deposito_avista" Then
    '    lst_bordero_deposito_avista.Show
    'ElseIf x_interno = "lst_caixa" Then
    '    lst_caixa.Show
    'ElseIf x_interno = "lst_cheque" Then
    '    lst_cheque.Show
    'ElseIf x_interno = "lst_cheque_devolvido" Then
    '    lst_cheque_devolvido.Show
    'ElseIf x_interno = "lst_cheque_avista" Then
    '    lst_cheque_avista.Show
    'ElseIf x_interno = "lst_cheque_baixados" Then
    '    lst_cheque_baixados.Show
    'ElseIf x_interno = "lst_cheque_deposito_bancario" Then
    '    lst_cheque_deposito_bancario.Show
    'ElseIf x_interno = "lst_cliente" Then
    '    lst_cliente.Show
    'ElseIf x_interno = "lst_cliente_conveniado" Then
    '    lst_cliente_conveniado.Show
    'ElseIf x_interno = "lst_contas_pagar" Then
    '    lst_contas_pagar.Show
    'ElseIf x_interno = "lst_contas_pagar2" Then
    '    lst_contas_pagar2.Show
    'ElseIf x_interno = "lst_contas_pagar_especial" Then
    '    lst_contas_pagar_especial.Show
    'ElseIf x_interno = "lst_contas_a_pagar_fornecedor" Then
    '    lst_contas_a_pagar_fornecedor.Show
    'ElseIf x_interno = "lst_duplicata_paga" Then
    '    lst_duplicata_paga.Show
    'ElseIf x_interno = "lst_duplicata_receber" Then
    '    lst_duplicata_receber.Show
    'ElseIf x_interno = "lst_entrada_combustivel" Then
    '    lst_entrada_combustivel.Show
    'ElseIf x_interno = "lst_entrada_produto" Then
    '    lst_entrada_produto.Show
    'ElseIf x_interno = "lst_entrada_produto_conferencia" Then
    '    lst_entrada_produto_conferencia.Show
    'ElseIf x_interno = "lst_falta_caixa" Then
    '    lst_falta_caixa.Show
    'ElseIf x_interno = "lst_falta_funcionario" Then
    '    lst_falta_funcionario.Show
    'ElseIf x_interno = "lst_estoque_medio" Then
    '    lst_estoque_medio.Show
    'ElseIf x_interno = "lst_historico" Then
    '    lst_historico.Show
    ElseIf x_interno = "lst_inventario_produto" Then
        lst_inventario_produto.Show
    ' ElseIf x_interno = "lst_leasing_veiculo" Then
    '     lst_leasing_veiculo.Show
    'ElseIf x_interno = "lst_lmc_abertura" Then
    '    lst_lmc_abertura.Show
    'ElseIf x_interno = "lst_nota_abastecimento_convenio" Then
    '    lst_nota_abastecimento_convenio.Show
    ElseIf x_interno = "lst_nota_cliente_emissao" Then
        lst_nota_cliente_emissao.Show
    'ElseIf x_interno = "lst_movimento_cartao" Then
    '    lst_movimento_cartao.Show
    'ElseIf x_interno = "lst_quadro_funcionario" Then
    '    lst_quadro_funcionario.Show
    'ElseIf x_interno = "lst_resumo_lmc" Then
    '    lst_resumo_lmc.Show
    ElseIf x_interno = "lst_venda_cupom" Then
        lst_venda_cupom.Show
    'ElseIf x_interno = "relatorio_cheque_cobranca" Then
    '    relatorio_cheque_cobranca.Show
    'ElseIf x_interno = "relatorio_cheque_folha" Then
    '    relatorio_cheque_folha.Show
    Else
        Screen.MousePointer = 1
    End If
End Sub
Private Sub ClickMenu(x_tipo As String, x_index As Integer)
    Dim i As Integer
    Dim interno As String
    i = -1
    interno = ""
    If tbl_menu.RecordCount > 0 Then
        tbl_menu.Seek ">=", g_usuario, x_tipo, " "
        If Not tbl_menu.NoMatch Then
            Do Until tbl_menu.EOF
                i = i + 1
                If x_index = i Then
                    interno = tbl_menu!interno
                    Exit Do
                End If
                tbl_menu.MoveNext
            Loop
        End If
    End If
    If interno = "" Then
        Exit Sub
    End If
    Chama interno
End Sub
Private Sub Finaliza()
    tbl_dados.Edit
    tbl_dados!Empresa = g_empresa
    tbl_dados![Conta Bancaria] = g_conta_bancaria
    tbl_dados.Update
    tbl_dados.Close
    tbl_empresa.Close
    tbl_configuracao.Close
    tbl_conta_bancaria.Close
    tbl_liberacao_digitacao.Close
    tbl_menu.Close
    bd_sgp.Close
    bd_sgp_b.Close
    bd_sgp_m.Close
    End
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
    tbl_menu.Index = "id_usuario"
    If tbl_menu.RecordCount > 0 Then
        tbl_menu.Seek ">=", g_usuario, "  ", " "
        If Not tbl_menu.NoMatch Then
            Do Until tbl_menu.EOF
                If tbl_menu!Usuario <> g_usuario Then
                    Exit Do
                End If
                If tbl_menu!Tipo <> x_tipo Then
                    x_tipo = tbl_menu!Tipo
                    i = -1
                End If
                i = i + 1
                If tbl_menu!Tipo = "CA" Then
                    l_sair = i
                    If i > 0 Then
                        Load mnu_cadastro_item(i)
                    End If
                    mnu_cadastro_item(i).Caption = "&" & tbl_menu!Menu
                ElseIf tbl_menu!Tipo = "CO" Then
                    i_co = i
                    If i > 0 Then
                        Load mnu_consulta_item(i)
                    End If
                    mnu_consulta_item(i).Caption = "&" & tbl_menu!Menu
                ElseIf tbl_menu!Tipo = "GR" Then
                    i_gr = i
                    If i > 0 Then
                        Load mnu_grafico_item(i)
                    End If
                    mnu_grafico_item(i).Caption = "&" & tbl_menu!Menu
                ElseIf tbl_menu!Tipo = "MO" Then
                    i_mo = i
                    If i > 0 Then
                        Load mnu_movimentacao_item(i)
                    End If
                    mnu_movimentacao_item(i).Caption = "&" & tbl_menu!Menu
                ElseIf tbl_menu!Tipo = "RE" Then
                    i_re = i
                    If i > 0 Then
                        Load mnu_relatorio_item(i)
                    End If
                    mnu_relatorio_item(i).Caption = "&" & tbl_menu!Menu
                End If
                tbl_menu.MoveNext
            Loop
        End If
    End If
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
Private Sub cbo_conta_bancaria_Click()
    If tbl_conta_bancaria.RecordCount > 0 Then
        If cbo_conta_bancaria.ListIndex <> -1 Then
            g_conta_bancaria = Mid(cbo_conta_bancaria, 41, 10)
            g_nome_conta = cbo_conta_bancaria
        Else
            cbo_conta_bancaria.SetFocus
        End If
    Else
        Screen.MousePointer = 11
        'cadastro_conta_bancaria.Show
    End If
End Sub
Private Sub cbo_conta_bancaria_GotFocus()
'    SendMessageLong cbo_conta_bancaria.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_empresa_Click()
    If Not (tbl_empresa.BOF And tbl_empresa.EOF) Then
        If cbo_empresa.ListIndex <> -1 Then
            g_empresa = cbo_empresa.ItemData(cbo_empresa.ListIndex)
            g_nome_empresa = cbo_empresa.Text
            PreencheCboContaBancaria
        Else
            cbo_empresa.SetFocus
        End If
    Else
        Screen.MousePointer = 11
        'cadastro_empresa.Show
    End If
End Sub
Private Sub cbo_empresa_GotFocus()
'    SendMessageLong cbo_empresa.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_empresa_LostFocus()
    If cbo_empresa.ListIndex <> -1 Then
        g_empresa = cbo_empresa.ItemData(cbo_empresa.ListIndex)
        g_nome_empresa = cbo_empresa.Text
        PreencheCboContaBancaria
    End If
    If g_empresa > 0 And tbl_empresa.RecordCount > 0 Then
        BuscaConfiguracao
        tbl_empresa.Seek "=", g_empresa
        If Not tbl_empresa.NoMatch Then
            g_cidade_empresa = Trim(tbl_empresa!Cidade)
        Else
            g_cidade_empresa = "Goiânia"
        End If
    End If
End Sub
Private Sub cmd_calc_Click()
    Dim retval As Long
    retval = Shell("calc", vbNormalFocus)
End Sub
Private Sub cmd_calendario_Click()
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    g_string = ""
End Sub
Private Sub cmd_configuracao_Click()
    Screen.MousePointer = 11
    config_liberacao_caixa.Show 1
End Sub
Private Sub cmd_senha_Click()
    LimpaMenu
    'cbo_empresa.ListIndex = -1
    g_nivel_acesso = 0
    frm_identificacao2.Show 1
    Me.Caption = "Cupom Fiscal Cerrado - " & UCase(Mid(g_nome_usuario, 1, 1)) & LCase(Mid(g_nome_usuario, 2, Len(g_nome_usuario) - 1))
    MontaMenu
    'BuscaDados
End Sub
Private Sub cmd_sql_Click()
    Screen.MousePointer = 11
    super_consulta.Show
End Sub
Private Sub Command1_Click()
    On Error GoTo FileError
    Dim old_sgp As Database
    Dim New_sgp As Database
    Dim Old_contas_pagar As Table
    Dim New_contas_pagar As Table
    Dim Old_baixa_pagar As Table
    Dim New_baixa_pagar As Table
    Dim x_registro(1 To 10) As Long
    
    Set old_sgp = OpenDatabase("OLD_DATA.MDB")
    Set New_sgp = OpenDatabase("SGP_DATA.MDB")
    Set Old_baixa_pagar = old_sgp.OpenTable("baixa_pagar")
    Set New_baixa_pagar = New_sgp.OpenTable("baixa_pagar")
    Set Old_contas_pagar = old_sgp.OpenTable("contas_pagar")
    Set New_contas_pagar = New_sgp.OpenTable("contas_pagar")
    
    Old_baixa_pagar.Index = "id_registro"
    New_baixa_pagar.Index = "id_registro"
    Old_contas_pagar.Index = "id_registro"
    New_contas_pagar.Index = "id_registro"
    
    'baixa_pagar
    Do Until Old_baixa_pagar.EOF
        x_registro(Old_baixa_pagar!Empresa) = x_registro(Old_baixa_pagar!Empresa) + 1
        New_baixa_pagar.AddNew
        New_baixa_pagar!Empresa = Old_baixa_pagar!Empresa
        New_baixa_pagar!Registro = x_registro(Old_baixa_pagar!Empresa)
        New_baixa_pagar!codigo_fornecedor = Old_baixa_pagar!codigo_fornecedor
        New_baixa_pagar!nome_fornecedor = Old_baixa_pagar!nome_fornecedor
        New_baixa_pagar!data_emissao = Old_baixa_pagar!data_emissao
        New_baixa_pagar!Data_Vencimento = Old_baixa_pagar!Data_Vencimento
        New_baixa_pagar!Valor = Old_baixa_pagar!Valor
        New_baixa_pagar!Numero_Documento = Old_baixa_pagar!Numero_Documento
        New_baixa_pagar!local_cobranca = Old_baixa_pagar!local_cobranca
        New_baixa_pagar!codigo_conta = Old_baixa_pagar!codigo_conta
        New_baixa_pagar!Complemento = Old_baixa_pagar!Complemento
        New_baixa_pagar!data_pagamento = Old_baixa_pagar!data_pagamento
        New_baixa_pagar!valor_pagamento = Old_baixa_pagar!valor_pagamento
        New_baixa_pagar.Update
        Old_baixa_pagar.MoveNext
    Loop
    
    'contas_pagar
    Do Until Old_contas_pagar.EOF
        x_registro(Old_contas_pagar!Empresa) = x_registro(Old_contas_pagar!Empresa) + 1
        New_contas_pagar.AddNew
        New_contas_pagar!Empresa = Old_contas_pagar!Empresa
        New_contas_pagar!Registro = x_registro(Old_contas_pagar!Empresa)
        New_contas_pagar!codigo_fornecedor = Old_contas_pagar!codigo_fornecedor
        New_contas_pagar!nome_fornecedor = Old_contas_pagar!nome_fornecedor
        New_contas_pagar!data_emissao = Old_contas_pagar!data_emissao
        New_contas_pagar!Data_Vencimento = Old_contas_pagar!Data_Vencimento
        New_contas_pagar!Valor = Old_contas_pagar!Valor
        New_contas_pagar!Numero_Documento = Old_contas_pagar!Numero_Documento
        New_contas_pagar!local_cobranca = Old_contas_pagar!local_cobranca
        New_contas_pagar!codigo_conta = Old_contas_pagar!codigo_conta
        New_contas_pagar!Complemento = Old_contas_pagar!Complemento
        New_contas_pagar.Update
        Old_contas_pagar.MoveNext
    Loop
    Exit Sub
FileError:
    ErroArquivo New_contas_pagar.Name, "Registroo"
    Exit Sub
End Sub
Private Sub BuscaNumeroDeSerie()
    Dim x_string As String
    Dim NumeroArquivo As Integer
    On Error GoTo FileError
    'busca número de série
    Call Abre_ProtocoloCF(1)
    ComandoCF = Chr(27) + "|35|00|" + Chr(27)
    Envia_ComandoCF
    Fecha_ProtocoloCF
    NumeroArquivo = FreeFile
    Open "MP20FI.RET" For Input As NumeroArquivo
    Input #NumeroArquivo, x_string
    Close NumeroArquivo
    If g_nome_empresa = "T-Kar Posto Shopping Ltda" And x_string = "4708990404338" Then
        Exit Sub
    ElseIf g_nome_empresa = "AUTO POSTO MANTIQUEIRA LTDA" And x_string = "4708990507288" Then
        Exit Sub
    ElseIf g_nome_empresa = "NOLETO E FILHAS LTDA" And x_string = "4708990713532" Then
        Exit Sub
    ElseIf g_nome_empresa = "AUTO POSTO VALE DO PIPIRIPAU LTDA" And x_string = "4708990713545" Then
        Exit Sub
    ElseIf g_nome_empresa = "BRAZUCA AUTO POSTO LTDA" And x_string = "4708990711791" Then
        Exit Sub
    ElseIf g_nome_empresa = "J. A. OLIVEIRA E CIA. LTDA." And x_string = "4708990918066" Then
        Exit Sub
    ElseIf g_nome_empresa = "BISPO E BATISTA LTDA" And x_string = "4708990919192" Then
        Exit Sub
    ElseIf g_nome_empresa = "AUTO POSTO VIA 63 LTDA" And x_string = "4708990711264" Then
        Exit Sub
    ElseIf g_nome_empresa = "POSTO URUAÇU LTDA" And x_string = "4708990711817" Then
        Exit Sub
    End If
    MsgBox "Número de Série da Impressora Fiscal ->" & x_string & "<-" & Chr(13) & "Empresa ->" & Trim(g_empresa) & "<-", vbInformation, "Número de Série"
    MsgBox "O sistema será finalizado", vbCritical, "Erro Interno Fatal"
    End
    Exit Sub
FileError:
    MsgBox "Não foi possível verificar o número de série.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub BuscaRegistroLiberacaoDigitacao()
    With tbl_liberacao_digitacao
        If .RecordCount > 0 Then
            .MoveFirst
            g_cfg_empresa_i = ![Empresa Inicial]
            g_cfg_empresa_f = ![Empresa Final]
            g_cfg_data_i = ![Data Inicial]
            g_cfg_data_f = ![Data Final]
            g_cfg_periodo_i = ![Periodo Inicial]
            g_cfg_periodo_f = ![Periodo Final]
        End If
    End With
End Sub
Private Sub Form_Activate()
    Dim x_chama_cupom As Boolean
    BuscaRegistroLiberacaoDigitacao
    If g_nivel_acesso <= 4 Then
        cmd_configuracao.Enabled = True
    Else
        cmd_configuracao.Enabled = False
    End If
    If g_nivel_acesso = 0 Then
        x_chama_cupom = True
        frm_identificacao2.Show 1
        If g_nivel_acesso <> 0 Then
            Me.Caption = "Cupom Fiscal Cerrado - " & UCase(Mid(g_nome_usuario, 1, 1)) & LCase(Mid(g_nome_usuario, 2, Len(g_nome_usuario) - 1))
            MontaMenu
            BuscaDados
            BuscaConfiguracao
        Else
            Finaliza
        End If
    End If
    If cbo_empresa.ListIndex = -1 Then
        PreencheCboEmpresa
    End If
    If cbo_conta_bancaria.ListIndex = -1 Then
        PreencheCboContaBancaria
    End If
    If cbo_empresa.ListIndex = -1 Then
        cbo_empresa.ListIndex = 0
    End If
    If x_chama_cupom Then
        x_chama_cupom = False
        If g_nome_usuario = "Cupom Fiscal" Then
            Screen.MousePointer = 11
            movimento_cupom_fiscal2.Show
        End If
    End If
End Sub
Private Sub Form_Load()
    'CentraForm Me
    'ChDrive "C"
    'If CDate(Date) >= CDate("30/04/1999") Then
    '    MsgBox Error(71), vbCritical, "Erro Grave"
    '    End
    'End If
    splash2.Show 1
'    Call ChamaDrive
'    ChDir "\VB5\SGP\DATA"
'    Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
'    Set bd_sgp_b = OpenDatabase("SGP_DATA_BAIXA.MDB")
'    Set bd_sgp_m = OpenDatabase("SGP_DATA_MOVIMENTO.MDB")
    Set tbl_dados = bd_sgp.OpenTable("dados")
    Set tbl_empresa = bd_sgp.OpenTable("empresas")
    Set tbl_configuracao = bd_sgp.OpenTable("configuracao")
    Set tbl_conta_bancaria = bd_sgp.OpenTable("Conta_Bancaria")
    Set tbl_liberacao_digitacao = bd_sgp.OpenTable("Liberacao_Digitacao")
    Set tbl_menu = bd_sgp.OpenTable("menu")
    tbl_configuracao.Index = "id_codigo"
    g_data_def = Date
    g_lmc = 0
End Sub
Private Sub PreencheCboEmpresa()
    Dim i As Integer
    With tbl_empresa
        If .RecordCount > 0 Then
            cbo_empresa.Clear
            .Index = "id_codigo"
            .MoveFirst
            Do Until .EOF
                '
                If .RecordCount > 1 Then
                    End
                End If
                If !Nome = "GOIAS MARTINS" Then
                    i = 0
                ElseIf !Nome = "POSTO URUAÇU LTDA" Then
                    i = 0
                ElseIf !Nome = "AUTO POSTO CEGÃO LTDA" Then
                    i = 0
                ElseIf !Nome = "CENTRO OESTE DIST.DERIV.DE PETRÓLEO LTDA" Then
                    i = 0
                'Acertar
                ElseIf !Nome = "URUNAÚTICA E DIESEL LTDA" Then
                    i = 0
                ElseIf !Nome = "NORMA CLAÚDIA FERREIRA" Then
                    i = 0
                ElseIf !Nome = "ROMEO ANTONIO GOEDEL" Then
                    i = 0
                Else
                    End
                End If
                If Not !Inativo Then
                    cbo_empresa.AddItem !Nome
                    cbo_empresa.ItemData(cbo_empresa.NewIndex) = !Codigo
                End If
                .MoveNext
            Loop
            For i = 0 To cbo_empresa.ListCount - 1
                If cbo_empresa.ItemData(i) = g_empresa Then
                    cbo_empresa.ListIndex = i
                    g_nome_empresa = cbo_empresa.Text
                    Exit For
                End If
            Next
        End If
    End With
End Sub
Private Sub PreencheCboContaBancaria()
    Dim i As Integer
    Dim x_flag As Integer
    Dim x_string As String * 50
    cbo_conta_bancaria.Clear
    tbl_conta_bancaria.Index = "id_nome"
    If tbl_conta_bancaria.RecordCount > 0 Then
        tbl_conta_bancaria.Seek ">", g_empresa, "          "
        If Not tbl_conta_bancaria.NoMatch Then
            If tbl_conta_bancaria!Empresa = g_empresa Then
                Do Until tbl_conta_bancaria.EOF
                    If tbl_conta_bancaria!Empresa <> g_empresa Then
                        Exit Do
                    End If
                    x_string = "                                                  "
                    Mid(x_string, 1, 40) = tbl_conta_bancaria!Nome
                    Mid(x_string, 41, 10) = tbl_conta_bancaria!Codigo
                    cbo_conta_bancaria.AddItem x_string
                    cbo_conta_bancaria.ItemData(cbo_conta_bancaria.NewIndex) = Val(tbl_conta_bancaria!Codigo)
                    tbl_conta_bancaria.MoveNext
                Loop
                For i = 0 To cbo_conta_bancaria.ListCount - 1
                    cbo_conta_bancaria.ListIndex = i
                    If Mid(cbo_conta_bancaria, 41, 10) = g_conta_bancaria Then
                        Exit For
                    End If
                Next
            Else
                x_flag = 1
            End If
        Else
            x_flag = 1
        End If
    Else
        x_flag = 1
    End If
    If x_flag = 0 Then
        If cbo_conta_bancaria.ListCount > 0 Then
            cbo_conta_bancaria.ListIndex = 0
        End If
    Else
        cbo_conta_bancaria.ListIndex = -1
        g_conta_bancaria = 0
    End If
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
    If Index = l_sair Then
        If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 6 Then
            Finaliza
        Else
            Exit Sub
        End If
    End If
    ClickMenu "CA", Index
End Sub
Private Sub mnu_consulta_item_Click(Index As Integer)
    ClickMenu "CO", Index
End Sub
Private Sub mnu_grafico_item_Click(Index As Integer)
    ClickMenu "GR", Index
End Sub
Private Sub mnu_movimentacao_item_Click(Index As Integer)
    ClickMenu "MO", Index
End Sub
Private Sub mnu_relatorio_item_Click(Index As Integer)
    ClickMenu "RE", Index
End Sub
Private Sub mnu_sobre_Click()
    If g_lmc = 1 Then
        g_lmc = 3
    Else
        g_lmc = g_lmc + 1
    End If
    Screen.MousePointer = 11
    frm_sobre2.Show 1
End Sub
