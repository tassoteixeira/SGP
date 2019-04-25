VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form menu_principal 
   Caption         =   "Gegênciador de Posto de Combustíveis"
   ClientHeight    =   6030
   ClientLeft      =   3390
   ClientTop       =   1845
   ClientWidth     =   7890
   ControlBox      =   0   'False
   Icon            =   "PRINCIPA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "PRINCIPA.frx":0442
   ScaleHeight     =   6030
   ScaleWidth      =   7890
   Begin VB.ComboBox cbo_conta_bancaria 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5640
      Width           =   4935
   End
   Begin VB.ComboBox cbo_empresa 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5220
      Width           =   4935
   End
   Begin Threed.SSCommand cmd_sql 
      Height          =   675
      Left            =   7080
      TabIndex        =   4
      Top             =   0
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   1191
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PRINCIPA.frx":0888
   End
   Begin VB.Label Label1 
      Caption         =   "Conta Bancária"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label Label7 
      Caption         =   "Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5220
      Width           =   1155
   End
   Begin VB.Menu mnu_ger_cadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnu_cad_bancos 
         Caption         =   "B&ancos"
      End
      Begin VB.Menu mnu_ger_cad_bombas 
         Caption         =   "&Bombas"
      End
      Begin VB.Menu mnu_cad_combustiveis 
         Caption         =   "&Combustíveis"
      End
      Begin VB.Menu mnu_cad_conta_corrente 
         Caption         =   "Co&nta Bancária"
      End
      Begin VB.Menu mnu_cad_convenios 
         Caption         =   "C&onvênios"
      End
      Begin VB.Menu mnu_cad_empresas 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnu_cad_fornecedores 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu mnu_cad_funcionarios 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mnu_cad_grupos 
         Caption         =   "&Grupos"
      End
      Begin VB.Menu mnu_cad_historico 
         Caption         =   "&Históricos"
      End
      Begin VB.Menu mnu_cad_menu 
         Caption         =   "&Menu"
      End
      Begin VB.Menu mnu_cad_produtos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnu_cad_programas 
         Caption         =   "&Programas"
      End
      Begin VB.Menu mnu_cad_vencimento 
         Caption         =   "&Tabela de Vencimento"
      End
      Begin VB.Menu mnu_cad_tipo_documento 
         Caption         =   "&Tipo de Documentos"
      End
      Begin VB.Menu mnu_usuarios 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu separador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnu_movimentacao 
      Caption         =   "&Movimentação"
      Begin VB.Menu mnu_bx_contas_pagar 
         Caption         =   "Baixa de Contas à Pagar"
      End
      Begin VB.Menu mnu_mov_caixa 
         Caption         =   "C&aixa"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_mov_cheques 
         Caption         =   "&Cheques"
      End
      Begin VB.Menu mnu_mov_contas_pagar 
         Caption         =   "Co&ntas à Pagar"
      End
      Begin VB.Menu mnu_mov_bancaria 
         Caption         =   "Ba&ncária"
      End
      Begin VB.Menu mnu_mov_bombas 
         Caption         =   "&Bombas"
      End
      Begin VB.Menu mnu_mov_ent_comb 
         Caption         =   "&Entradas de Combustíveis"
      End
      Begin VB.Menu mnu_mov_med_comb 
         Caption         =   "&Medição de Combustíveis"
      End
      Begin VB.Menu mnu_mov_pedidos_combustiveis 
         Caption         =   "&Pedidos de Combustíveis"
      End
   End
   Begin VB.Menu mnu_relatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnu_rel_avulso 
         Caption         =   "Imprime Cheques (&avulso)"
      End
      Begin VB.Menu mnu_rel_ch_formulario 
         Caption         =   "Imprime &Cheques (formulário)"
      End
      Begin VB.Menu mnu_rel_cheques_baixados 
         Caption         =   "Emissão de Cheques &Baixados"
      End
      Begin VB.Menu mnu_rel_cheques_predatados 
         Caption         =   "Emissão de Cheques &Pré-Datados"
      End
      Begin VB.Menu mnu_rel_contas_pagar 
         Caption         =   "Emissão de Contas à P&agar"
      End
      Begin VB.Menu mnu_rel_contas_pagar2 
         Caption         =   "Emissão de Contas à Pagar (&Especial)"
      End
      Begin VB.Menu mnu_rel_recibos_folha 
         Caption         =   "Emissão de R&ecibos (Folha)"
      End
      Begin VB.Menu mnu_rel_recibo_formulario 
         Caption         =   "Emissão de &Recibos (Formulário)"
      End
      Begin VB.Menu mnu_rel_estrato_bancario 
         Caption         =   "Emissao de &Extrato Bancário"
      End
      Begin VB.Menu mnu_rel_lst_cheque_form 
         Caption         =   "Emissão de  Cheques (&formulario)"
      End
      Begin VB.Menu rel_lmc 
         Caption         =   "Emissão de &L.M.C."
      End
      Begin VB.Menu mnu_rel_historico 
         Caption         =   "Emissão do &Histórico"
      End
      Begin VB.Menu mnu_rel_pagamentos 
         Caption         =   "Emissão dos Pagamentos"
      End
   End
   Begin VB.Menu mnu_graficos 
      Caption         =   "&Gráficos"
      Begin VB.Menu mnu_grf_vendas_combustiveis_2 
         Caption         =   "Vendas de Combustíveis (&Meses)"
      End
      Begin VB.Menu mnu_graf_vendas_combustiveis 
         Caption         =   "&Vendas de Combustíveis"
      End
   End
   Begin VB.Menu mnu_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnu_con_cheque_predatados 
         Caption         =   "Consulta Cheques &Pré-Datados"
      End
   End
   Begin VB.Menu mnu_sobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "menu_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BuscaDados()
    If (tbl_dados.BOF And tbl_dados.EOF) Then
        tbl_dados.AddNew
        tbl_dados!Codigo = 1
        tbl_dados!Empresa = 1
        g_empresa = tbl_dados!Empresa
        tbl_dados.Update
    End If
    tbl_dados.MoveFirst
    g_empresa = tbl_dados!Empresa
    g_conta_bancaria = tbl_dados!conta_bancaria
End Sub
Private Sub Finaliza()
    tbl_dados.Edit
    tbl_dados!Empresa = g_empresa
    tbl_dados!conta_bancaria = g_conta_bancaria
    tbl_dados.Update
    tbl_dados.Close
    tbl_empresa.Close
    tbl_conta_bancaria.Close
    tbl_menu.Close
    bd_sgp.Close
    End
End Sub
Private Sub cbo_conta_bancaria_Click()
    If Not (tbl_conta_bancaria.BOF And tbl_conta_bancaria.EOF) Then
        If cbo_conta_bancaria.ListIndex <> -1 Then
            g_conta_bancaria = Mid(cbo_conta_bancaria, 41, 10)
            g_nome_conta = cbo_conta_bancaria
        Else
            cbo_conta_bancaria.SetFocus
        End If
    Else
        Screen.MousePointer = 11
        frm_cadastro_contas.Show 1
    End If
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
        cadastro_empresa.Show 1
    End If
End Sub
Private Sub cmd_sql_Click()
    Screen.MousePointer = 11
    frm_super_consulta.Show 1
End Sub
Private Sub Form_Activate()
    If cbo_empresa.ListIndex = -1 Then
        PreencheCboEmpresa
    End If
    If cbo_conta_bancaria.ListIndex = -1 Then
        PreencheCboContaBancaria
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    ChDir "\VB4\SGP"
    Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
    Set tbl_dados = bd_sgp.OpenTable("dados")
    Set tbl_empresa = bd_sgp.OpenTable("empresas")
    Set tbl_conta_bancaria = bd_sgp.OpenTable("contas_bancarias")
    Set tbl_menu = bd_sgp.OpenTable("menu")
    g_data_def = Date
    BuscaDados
End Sub
Private Sub PreencheCboEmpresa()
    Dim i As Integer
    cbo_empresa.Clear
    tbl_empresa.Index = "id_codigo"
    If Not (tbl_empresa.BOF And tbl_empresa.EOF) Then
        tbl_empresa.MoveFirst
        Do Until tbl_empresa.EOF
            cbo_empresa.AddItem tbl_empresa!Nome
            cbo_empresa.ItemData(cbo_empresa.NewIndex) = tbl_empresa!Codigo
            tbl_empresa.MoveNext
        Loop
        For i = 0 To cbo_empresa.ListCount - 1
            If cbo_empresa.ItemData(i) = g_empresa Then
                cbo_empresa.ListIndex = i
                g_nome_empresa = cbo_empresa.Text
                Exit For
            End If
        Next
    Else
        cbo_empresa.ListIndex = -1
    End If
End Sub
Private Sub PreencheCboContaBancaria()
    Dim i As Integer
    Dim x_flag As Integer
    Dim x_string As String * 50
    cbo_conta_bancaria.Clear
    tbl_conta_bancaria.Index = "id_nome"
    If Not (tbl_conta_bancaria.BOF And tbl_conta_bancaria.EOF) Then
        tbl_conta_bancaria.Seek ">", g_empresa, "          "
        If Not tbl_conta_bancaria.NoMatch Then
            If tbl_conta_bancaria!Empresa = g_empresa Then
                Do Until tbl_conta_bancaria.EOF
                    If tbl_conta_bancaria!Empresa <> g_empresa Then
                        Exit Do
                    End If
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
Private Sub mnu_bx_contas_pagar_Click()
    Screen.MousePointer = 11
    bai_contas_pagar.Show 1
End Sub
Private Sub mnu_cad_bancos_Click()
    Screen.MousePointer = 11
    frm_cadastro_bancos.Show 1
End Sub
Private Sub mnu_cad_combustiveis_Click()
    Screen.MousePointer = 11
    cadastro_combustivel.Show 1
End Sub
Private Sub mnu_cad_conta_corrente_Click()
    Screen.MousePointer = 11
    frm_cadastro_contas.Show 1
End Sub
Private Sub mnu_cad_convenios_Click()
    Screen.MousePointer = 11
    cadastro_convenio.Show 1
End Sub
Private Sub mnu_cad_empresas_Click()
    Screen.MousePointer = 11
    cadastro_empresa.Show 1
End Sub
Private Sub mnu_cad_fornecedores_Click()
    Screen.MousePointer = 11
    frm_cadastro_fornecedores.Show 1
End Sub
Private Sub mnu_cad_funcionarios_Click()
    Screen.MousePointer = 11
    cadastro_funcionario.Show 1
End Sub
Private Sub mnu_cad_grupos_Click()
    Screen.MousePointer = 11
    cadastro_grupo.Show 1
End Sub
Private Sub mnu_cad_historico_Click()
    Screen.MousePointer = 11
    cadastro_historico.Show 1
End Sub
Private Sub mnu_cad_menu_Click()
    Screen.MousePointer = 11
    cad_menu.Show 1
End Sub
Private Sub mnu_cad_produtos_Click()
    Screen.MousePointer = 11
    cadastro_produto.Show 1
End Sub
Private Sub mnu_cad_programas_Click()
    Screen.MousePointer = 11
    cadastro_programa.Show 1
End Sub
Private Sub mnu_cad_tipo_documento_Click()
    Screen.MousePointer = 11
    frm_cadastro_tipo_documentos.Show 1
End Sub
Private Sub mnu_cad_vencimento_Click()
    Screen.MousePointer = 11
    cadastro_vencimento.Show 1
End Sub
Private Sub mnu_con_cheque_predatados_Click()
'    Screen.MousePointer = 11
'    con_cheque_predatados.Show 1
End Sub
Private Sub mnu_ger_cad_bombas_Click()
    Screen.MousePointer = 11
    cadastro_bomba.Show 1
End Sub
Private Sub mnu_graf_vendas_combustiveis_Click()
    Screen.MousePointer = 11
    frm_graf_venda_combustiveis.Show 1
End Sub
Private Sub mnu_grf_vendas_combustiveis_2_Click()
    Screen.MousePointer = 11
    frm_graf_comp_combustiveis.Show 1
End Sub
Private Sub mnu_mov_bancaria_Click()
    Screen.MousePointer = 11
    movimento_bancario.Show 1
End Sub
Private Sub mnu_mov_bombas_Click()
    Screen.MousePointer = 11
    movimento_bomba.Show 1
End Sub
Private Sub mnu_mov_caixa_Click()
    'Screen.MousePointer = 11
    'frm_movimento_caixa.Show 1
End Sub
Private Sub mnu_mov_cheques_Click()
    Screen.MousePointer = 11
    movimento_cheque.Show 1
End Sub
Private Sub mnu_mov_contas_pagar_Click()
    Screen.MousePointer = 11
    mov_contas_pagar.Show 1
End Sub
Private Sub mnu_mov_ent_comb_Click()
    Screen.MousePointer = 11
    mov_entrada_combustiveis.Show 1
End Sub
Private Sub mnu_mov_med_comb_Click()
    Screen.MousePointer = 11
    mov_medicao_combustiveis.Show 1
End Sub
Private Sub mnu_mov_pedidos_combustiveis_Click()
    Screen.MousePointer = 11
    movimento_pedido_combustivel.Show 1
End Sub
Private Sub mnu_rel_avulso_Click()
    Screen.MousePointer = 11
    frm_emissao_cheques_folhas.Show 1
End Sub
Private Sub mnu_rel_ch_formulario_Click()
    Screen.MousePointer = 11
    frm_emissao_cheques_formulario.Show 1
End Sub
Private Sub mnu_rel_cheques_baixados_Click()
    Screen.MousePointer = 11
    lst_cheque_baixados.Show 1
End Sub
Private Sub mnu_rel_cheques_predatados_Click()
    Screen.MousePointer = 11
    lst_cheque_predatados.Show 1
End Sub
Private Sub mnu_rel_contas_pagar_Click()
    Screen.MousePointer = 11
    lst_contas_pagar.Show 1
End Sub
Private Sub mnu_rel_contas_pagar2_Click()
    Screen.MousePointer = 11
    lst_contas_pagar2.Show 1
End Sub
Private Sub mnu_rel_estrato_bancario_Click()
    Screen.MousePointer = 11
    emissao_extrato_bancario.Show 1
End Sub
Private Sub mnu_rel_historico_Click()
    Screen.MousePointer = 11
    lst_historico.Show 1
End Sub

Private Sub mnu_rel_pagamentos_Click()
    Screen.MousePointer = 11
    lst_baixa_pagar.Show 1
End Sub
Private Sub mnu_rel_recibo_formulario_Click()
    Screen.MousePointer = 11
    frm_emissao_recibo_formulario.Show 1
End Sub
Private Sub mnu_rel_recibos_folha_Click()
    Screen.MousePointer = 11
    frm_emissao_recibo_folhas.Show 1
End Sub
Private Sub mnu_sair_Click()
    If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 6 Then
        Finaliza
    End If
End Sub
Private Sub mnu_sobre_Click()
    Screen.MousePointer = 11
    frm_sobre.Show 1
End Sub
Private Sub mnu_usuarios_Click()
    Screen.MousePointer = 11
    cad_usuario.Show 1
End Sub
Private Sub rel_lmc_Click()
    Screen.MousePointer = 11
    frm_emissao_lmc.Show 1
End Sub
