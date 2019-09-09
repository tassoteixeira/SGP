VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form baixa_contas_pagar 
   Caption         =   "Baixa de Contas à Pagar"
   ClientHeight    =   7785
   ClientLeft      =   2010
   ClientTop       =   465
   ClientWidth     =   8520
   Icon            =   "baixa_contas_pagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_contas_pagar.frx":030A
   ScaleHeight     =   7785
   ScaleWidth      =   8520
   Begin VB.CommandButton cmd_extornar 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   1020
      Picture         =   "baixa_contas_pagar.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Extorna o registro atual."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "baixa_contas_pagar.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   1920
      Picture         =   "baixa_contas_pagar.frx":30BC
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   120
      Picture         =   "baixa_contas_pagar.frx":452E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Altera o registro atual."
      Top             =   6840
      Width           =   795
   End
   Begin VB.Frame frmBaixa 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   8295
      Begin VB.TextBox txt_valor_pagamento 
         Height          =   285
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Frame frm_dados 
         Height          =   2715
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8055
         Begin VB.CheckBox chk_despesa_caixa 
            Caption         =   "Despesa de Caixa"
            ForeColor       =   &H00400040&
            Height          =   195
            Left            =   6300
            TabIndex        =   4
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Número do Registro"
            ForeColor       =   &H00400040&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label lbl_registro 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   3
            Top             =   180
            Width           =   675
         End
         Begin VB.Label lbl_complemento 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   20
            Top             =   2340
            Width           =   4275
         End
         Begin VB.Label lbl_conta 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   18
            Top             =   1980
            Width           =   4275
         End
         Begin VB.Label lbl_local_cobranca 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   16
            Top             =   1620
            Width           =   4275
         End
         Begin VB.Label lbl_valor_vencimento 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6840
            TabIndex        =   14
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lbl_numero_documento 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6840
            TabIndex        =   10
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label lbl_data_vencimento 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   12
            Top             =   1260
            Width           =   1035
         End
         Begin VB.Label lbl_data_emissao 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   8
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lbl_nome_fornecedor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   6
            Top             =   540
            Width           =   4275
         End
         Begin VB.Label Label3 
            Caption         =   "Conta"
            ForeColor       =   &H00400040&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1980
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Número Documento"
            ForeColor       =   &H00400040&
            Height          =   315
            Left            =   5220
            TabIndex        =   9
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Complemento"
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2340
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Código Fornecedor"
            ForeColor       =   &H00400040&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   540
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Data de Emissão"
            ForeColor       =   &H00400040&
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Data do Vencimento"
            ForeColor       =   &H00400040&
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1260
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Valor do Vencimento"
            ForeColor       =   &H00400040&
            Height          =   315
            Left            =   5220
            TabIndex        =   13
            Top             =   1260
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Local de Cobrança"
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1620
            Width           =   1350
         End
      End
      Begin MSMask.MaskEdBox msk_data_pagamento 
         Height          =   300
         Left            =   1860
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Pagamento"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Pagamento"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   2880
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6240
      TabIndex        =   31
      Top             =   6720
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "baixa_contas_pagar.frx":5A28
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "baixa_contas_pagar.frx":6FAA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "baixa_contas_pagar.frx":841C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "baixa_contas_pagar.frx":9916
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7620
      Picture         =   "baixa_contas_pagar.frx":AE10
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6720
      Picture         =   "baixa_contas_pagar.frx":C30A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6840
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   36
      Top             =   60
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
   End
End
Attribute VB_Name = "baixa_contas_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lOpcao As Integer
Dim flag_baixa_contas_pagar As Integer
Dim lTipoRegistro As String
Dim lRegistro As Long
Dim lSQL As String
Dim lData As Date
Dim lNumeroMovimentoCaixa As Long
Dim lNumeroMovimentoCaixaBaixa As Long
Dim lNumeroMovimentoFinanceiro As Integer
Dim lIntegracaoComFinanceiroDiario As Boolean

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxTipoMovimento As Integer
Dim lCxTipoMov As String
Dim lCxIlha As Integer
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxCodigoPortadorFinanceiro As Integer

Private BaixaPagar As New cBaixaPagar
Private Contas As New cContas
Private IntegracaoCaixa As New cIntegracaoCaixa
Private Fornecedor As New cFornecedor
Private LancamentoFinanceiro As New cLancamentoFinanceiro
Private LocalCobranca As New cLocalCobranca
Private MovCaixa As New cMovimentoCaixa
Private MovContaPagar As New cMovimentoContaPagar
Private MovimentoFinanceiro As New cMovimentoFinanceiro
Private PortadorFinanceiro As New cPortadorFinanceiro

Private ConfiguracaoDiversa As New cConfiguracaoDiversa

Private Sub AjustaCaixaPista()
    Dim xString As String
    Dim xOperacao As String
    
    xString = g_string
    xOperacao = RetiraString(2, xString)
    g_string = ""

    lCxData = CDate(RetiraString(3, xString))
    lCxPeriodo = RetiraString(4, xString)
    lCxTipoMovimento = RetiraString(5, xString)
    lCxIlha = Val(RetiraString(6, xString))
    lCxTipoMov = RetiraString(7, xString)
    lCxCodigoLancamentoPadrao = Val(RetiraString(8, xString))
    lCxCodigoPortadorFinanceiro = Val(RetiraString(9, xString))

    If xOperacao = "Incluir" Then
        MSFlexGrid.SetFocus
    ElseIf xOperacao = "Alterar" Then
        lRegistro = RetiraString(10, xString)
        lNumeroMovimentoFinanceiro = Val(RetiraString(11, xString))
        'para manter compatibilidade com movimentos antigos
        'quando nao tinha ordem de lancamento,
        'ou seja aceitava apenas um vale por funcionario.
        If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lRegistro = RetiraString(10, xString)
        lNumeroMovimentoFinanceiro = Val(RetiraString(11, xString))
        If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        End If
        cmd_extornar_Click
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_extornar.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    Dim rsTabela As ADODB.Recordset
    LimpaMSFlexGrid
    lSQL = ""
    lSQL = lSQL & "SELECT nome_fornecedor, data_vencimento, valor, complemento, numero_documento, local_cobranca, codigo_conta, data_emissao, empresa, registro, codigo_fornecedor"
    lSQL = lSQL & "  FROM contas_pagar"
    lSQL = lSQL & " WHERE contas_pagar.empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY data_vencimento"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela("nome_fornecedor").Value
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela("data_vencimento").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = Format(rsTabela("valor").Value, "###,###,##0.00")
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("complemento").Value
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela("numero_documento").Value
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela("local_cobranca").Value
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela("codigo_conta").Value
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela("data_emissao").Value
            MSFlexGrid.Col = 8
            MSFlexGrid.Text = rsTabela("empresa").Value
            MSFlexGrid.Col = 9
            MSFlexGrid.Text = rsTabela("registro").Value
            MSFlexGrid.Col = 10
            MSFlexGrid.Text = rsTabela("codigo_fornecedor").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub AtualTabe()
    BaixaPagar.Empresa = MovContaPagar.Empresa
    BaixaPagar.Registro = MovContaPagar.Registro
    BaixaPagar.CodigoFornecedor = MovContaPagar.CodigoFornecedor
    BaixaPagar.NomeFornecedor = lbl_nome_fornecedor.Caption
    BaixaPagar.DataEmissao = MovContaPagar.DataEmissao
    BaixaPagar.DataVencimento = MovContaPagar.DataVencimento
    BaixaPagar.Valor = MovContaPagar.Valor
    BaixaPagar.NumeroDocumento = MovContaPagar.NumeroDocumento
    BaixaPagar.LocalCobranca = MovContaPagar.LocalCobranca
    BaixaPagar.CodigoConta = MovContaPagar.CodigoConta
    BaixaPagar.Complemento = MovContaPagar.Complemento
    BaixaPagar.DataDigitacao = MovContaPagar.DataDigitacao
    BaixaPagar.DataPagamento = Format(msk_data_pagamento.Text, "dd/mm/yyyy")
    BaixaPagar.ValorPagamento = fValidaValor2(txt_valor_pagamento.Text)
    BaixaPagar.NumeroMovimentoCaixa = MovContaPagar.NumeroMovimentoCaixa
    BaixaPagar.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
    BaixaPagar.TipoBaixa = 1
    If IsNull(BaixaPagar.NumeroMovimentoCaixa) Then
        BaixaPagar.NumeroMovimentoCaixa = 0
    End If
    If IsNull(BaixaPagar.NumeroMovimentoCaixaBaixa) Then
        BaixaPagar.NumeroMovimentoCaixaBaixa = 0
    End If
End Sub
Private Sub Atualtabe_2()
    BaixaPagar.DataPagamento = CDate(msk_data_pagamento.Text)
    BaixaPagar.ValorPagamento = fValidaValor2(txt_valor_pagamento.Text)
    BaixaPagar.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
End Sub
Private Sub MostraVencimento()
    Dim xRegistroNovo As Long
    
    lOpcao = 1
    lTipoRegistro = "C"
    If MovContaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
        If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            xRegistroNovo = MovContaPagar.AlteraNumeroRegistroAutomatico(g_empresa, lRegistro)
            If xRegistroNovo > 0 Then
                MsgBox "O Registro foi corrigido automaticamente para poder ser baixado.", vbInformation, "Correção Automática!"
                lRegistro = xRegistroNovo
                If Not MovContaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
                    MsgBox "Registro nao encontrato para ser baixado." & vbCrLf & "Entre em contato com o suporte técnico.", vbCritical, "Erro de Integridade!"
                End If
                AtualizaMSFlexGrid
            Else
                MsgBox "Não será possível baixar este registro." & vbCrLf & "Entre em contato com o suporte técnico.", vbCritical, "Erro ao tentar corrigir o registro!"
            End If
        End If
        lbl_registro.Caption = MovContaPagar.Registro
        If MovContaPagar.CodigoFornecedor > 0 Then
            If Not Fornecedor.LocalizarCodigo(g_empresa, MovContaPagar.CodigoFornecedor) Then
                lbl_nome_fornecedor.Caption = "** Não Cadastrado **"
            Else
                lbl_nome_fornecedor.Caption = Fornecedor.Nome
            End If
        Else
            lbl_nome_fornecedor.Caption = MovContaPagar.NomeFornecedor
        End If
        lbl_data_emissao.Caption = MovContaPagar.DataEmissao
        lbl_numero_documento.Caption = MovContaPagar.NumeroDocumento
        lbl_data_vencimento.Caption = MovContaPagar.DataVencimento
        lbl_valor_vencimento.Caption = Format(MovContaPagar.Valor, "###,##0.00")
        If LocalCobranca.LocalizarCodigo(MovContaPagar.LocalCobranca, g_empresa) Then
            lbl_local_cobranca.Caption = LocalCobranca.Nome
        Else
            lbl_local_cobranca.Caption = "** Não Cadastrada **"
        End If
        If Contas.LocalizarCodigo(MovContaPagar.CodigoConta, MovContaPagar.Empresa) Then
            lbl_conta.Caption = Contas.Nome
        Else
            lbl_conta.Caption = "** Não Cadastrada **"
        End If
        lbl_complemento.Caption = MovContaPagar.Complemento
        msk_data_pagamento = "__/__/____"
        txt_valor_pagamento = ""
    Else
        MsgBox "Vencimento não cadastrado.", vbInformation, "Erro de consistência!"
    End If
End Sub
Private Sub AtualTela()
    lOpcao = 2
    lTipoRegistro = "B"
    lRegistro = BaixaPagar.Registro
    lData = BaixaPagar.DataPagamento
    lNumeroMovimentoCaixaBaixa = BaixaPagar.NumeroMovimentoCaixaBaixa
    If BaixaPagar.TipoBaixa = 2 Then
        chk_despesa_caixa.Value = 1
    Else
        chk_despesa_caixa.Value = 0
    End If
    lbl_registro.Caption = BaixaPagar.Registro
    If BaixaPagar.CodigoFornecedor > 0 Then
        If Not Fornecedor.LocalizarCodigo(g_empresa, BaixaPagar.CodigoFornecedor) Then
            lbl_nome_fornecedor.Caption = "** Não Cadastrado **"
        Else
            lbl_nome_fornecedor.Caption = Fornecedor.Nome
        End If
    Else
        lbl_nome_fornecedor.Caption = BaixaPagar.NomeFornecedor
    End If
    lbl_data_emissao.Caption = BaixaPagar.DataEmissao
    lbl_numero_documento.Caption = BaixaPagar.NumeroDocumento
    lbl_data_vencimento.Caption = BaixaPagar.DataVencimento
    lbl_valor_vencimento.Caption = Format(BaixaPagar.Valor, "###,##0.00")
    If LocalCobranca.LocalizarCodigo(BaixaPagar.LocalCobranca, g_empresa) Then
        lbl_local_cobranca.Caption = LocalCobranca.Nome
    Else
        lbl_local_cobranca.Caption = "** Não Cadastrada **"
    End If
    If Contas.LocalizarCodigo(BaixaPagar.CodigoConta, BaixaPagar.Empresa) Then
        lbl_conta.Caption = Contas.Nome
    Else
        lbl_conta.Caption = "** Não Cadastrada **"
    End If
    lbl_complemento.Caption = BaixaPagar.Complemento
    msk_data_pagamento.Text = Format(BaixaPagar.DataPagamento, "dd/mm/yyyy")
    txt_valor_pagamento = Format(BaixaPagar.ValorPagamento, "###,##0.00")
    frmBaixa.Enabled = False
    If BaixaPagar.TipoBaixa = 2 Then
        cmd_alterar.Enabled = False
        cmd_extornar.Enabled = False
    Else
        cmd_alterar.Enabled = True
        cmd_extornar.Enabled = True
    End If
End Sub
Private Sub DesativaBotoes()
    cmd_alterar.Enabled = False
    cmd_extornar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = False
End Sub
Private Sub ExcluiMovimentoCaixa()
    If lNumeroMovimentoCaixaBaixa > 0 Then
        If Not MovCaixa.Excluir(g_empresa, lData, lNumeroMovimentoCaixaBaixa) Then
            MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set BaixaPagar = Nothing
    Set Contas = Nothing
    Set IntegracaoCaixa = Nothing
    Set Fornecedor = Nothing
    Set LocalCobranca = Nothing
    Set LancamentoFinanceiro = Nothing
    Set MovCaixa = Nothing
    Set MovContaPagar = Nothing
    Set MovimentoFinanceiro = Nothing
    Set PortadorFinanceiro = Nothing
End Sub
Private Sub cmd_alterar_Click()
    '
    'bd_sgp.Execute "UpDate Baixa_Pagar Set [Data da Digitacao] = Data_Emissao Where [Data da Digitacao] = Null"
    '
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    If lTipoRegistro = "B" Then
        frmBaixa.Enabled = True
        msk_data_pagamento.SetFocus
        Exit Sub
    End If
End Sub
Private Sub cmd_alterar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaBaixaContasPagar
    End If
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If BaixaPagar.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Dim x_flag As Integer
    
    Call GravaAuditoria(1, Me.name, 9, "")
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
    If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
        DesativaBotoes
        cmd_sair.Enabled = True
        LimpaTela
        MSFlexGrid.SetFocus
    Else
        AtivaBotoes
        If BaixaPagar.LocalizarUltimo(g_empresa) Then
            AtualTela
        Else
            LimpaTela
        End If
        'cmd_alterar.SetFocus
        MSFlexGrid.SetFocus
    End If
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xNome As String
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xCodigoHistorico As Integer
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixaBaixa = 0
    
    xNome = ""
    If Fornecedor.CodigoConta = 1 Then
        xNome = "-ESTOQUE"
    End If
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "BAIXA CONTAS A PAGAR" & xNome) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        xCodigoHistorico = IntegracaoCaixa.HistoricoPadrao
    Else
        xContaDebito = "221050003"
        xContaCredito = "111010001"
        xCodigoHistorico = 29
    End If
    
    xComplemento = Trim(lbl_complemento.Caption)  '& " TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
        
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(msk_data_pagamento.Text)
    MovCaixa.NumeroMovimento = 1
    MovCaixa.Valor = fValidaValor(txt_valor_pagamento.Text)
    MovCaixa.NumeroDocumento = lbl_numero_documento.Caption
    MovCaixa.CodigoHistorico = xCodigoHistorico
    MovCaixa.Complemento = xComplemento
    MovCaixa.NumeroContaDebito = xContaDebito
    If Len(Fornecedor.ContaContabil) = 9 Then
        MovCaixa.NumeroContaDebito = Fornecedor.ContaContabil
    End If
    MovCaixa.NumeroContaCredito = xContaCredito
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = True
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixa = True
        lNumeroMovimentoCaixaBaixa = MovCaixa.NumeroMovimento
    Else
        MsgBox "Não foi integrado no caixa o valor=" & txt_valor_pagamento.Text, vbInformation, "Erro de Integridade"
    End If
End Function
Function IncluiMovimentoCaixaAntesBaixa() As Boolean
    Dim xComplemento As String
    Dim xNome As String
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xCodigoHistorico As Integer
    IncluiMovimentoCaixaAntesBaixa = False
    lNumeroMovimentoCaixaBaixa = 0
    
    xNome = ""
    If Fornecedor.CodigoConta = 1 Then
        xNome = "-ESTOQUE"
    End If
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "CONTAS A PAGAR" & xNome) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        xCodigoHistorico = IntegracaoCaixa.HistoricoPadrao
    Else
        xContaDebito = "421030028"
        xContaCredito = "221050003"
        xCodigoHistorico = 18
    End If
    
    xComplemento = Trim(lbl_complemento.Caption)  '& " TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
        
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(lbl_data_emissao.Caption)
    MovCaixa.NumeroMovimento = 1
    MovCaixa.Valor = fValidaValor(lbl_valor_vencimento.Caption)
    MovCaixa.NumeroDocumento = lbl_numero_documento.Caption
    MovCaixa.CodigoHistorico = xCodigoHistorico
    MovCaixa.Complemento = xComplemento
    MovCaixa.NumeroContaDebito = xContaDebito
    MovCaixa.NumeroContaCredito = xContaCredito
    If Len(Fornecedor.ContaContabil) = 9 Then
        MovCaixa.NumeroContaCredito = Fornecedor.ContaContabil
    End If
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = True
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixaAntesBaixa = True
        lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
    Else
        MsgBox "Não foi integrado no caixa o valor=" & lbl_valor_vencimento.Caption, vbInformation, "Erro de Integridade"
    End If
End Function
Private Function IncluiMovimentoFinanceiroDiario() As Boolean
On Error GoTo FileError

IncluiMovimentoFinanceiroDiario = False

        Dim xMovimentoTesouraria As New CadastroDLL.cFinMovimentoTesouraria
        Dim SaldoTesouraria As New CadastroDLL.cFinSaldoTesouraria

        Dim xCodigoTipoMovimento As Integer
        Dim xCodigoFinConta As Integer
        xCodigoTipoMovimento = 0
        xCodigoFinConta = 0
        If LocalCobranca.LocalizarCodigo(BaixaPagar.LocalCobranca, BaixaPagar.Empresa) Then
            xCodigoTipoMovimento = LocalCobranca.CodigoFinTipoMovimento
        End If
        
        If Contas.LocalizarCodigo(BaixaPagar.CodigoConta, BaixaPagar.Empresa) Then
            xCodigoFinConta = Contas.CodigoFinContaTesouraria
        End If

        If xCodigoTipoMovimento = 0 Or xCodigoFinConta = 0 Then
            MsgBox "Não foi possível integrar lançamento ao Financeiro Diário. xCodigoTipoMovimento = " & xCodigoTipoMovimento & " - xCodigoFinConta=" & xCodigoFinConta, vbCritical, "Erro Desconhecido!"
            Exit Function
        End If


        xMovimentoTesouraria.Empresa = g_empresa
        xMovimentoTesouraria.Data = BaixaPagar.DataPagamento
        xMovimentoTesouraria.NumeroMovimento = xMovimentoTesouraria.ProximoCodigo(g_empresa, BaixaPagar.DataPagamento) 'CInt(txtNumeroMovimento.Text)
        xMovimentoTesouraria.CodigoTipoMovimento = xCodigoTipoMovimento
        xMovimentoTesouraria.Historico = lbl_complemento.Caption
        xMovimentoTesouraria.ValorEntrada = 0 'fValidaValor(txtValorEntrada.Text)
        xMovimentoTesouraria.ValorSaida = BaixaPagar.ValorPagamento
        xMovimentoTesouraria.CodigoContaTesouraria = xCodigoFinConta
        xMovimentoTesouraria.RegistroContaPagar = BaixaPagar.Registro
        
        If xMovimentoTesouraria.Incluir Then
            IncluiMovimentoFinanceiroDiario = True
            If Not SaldoTesouraria.AlterarSaldo(g_empresa, xCodigoTipoMovimento, BaixaPagar.DataPagamento, BaixaPagar.ValorPagamento, False) Then
                MsgBox "IncluiMovimentoFinanceiroDiario: Não foi possível atualizar saldo do financeiro diário!", vbCritical, "Erro Desconhecido!"
            End If
        End If

    Exit Function

FileError:
    MsgBox "IncluiMovimentoFinanceiroDiario: Erro não identificado!", vbCritical, "Erro Desconhecido!"

End Function



Private Function IncluiMovimentoFinanceiro(ByVal pFormaPagamento As String) As Boolean
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xHistoricoPadrao As Integer
    Dim xCodigoLancamentoFinanceiro As Integer

On Error GoTo FileError
    
    IncluiMovimentoFinanceiro = False
    xCodigoLancamentoFinanceiro = 9
    If LancamentoFinanceiro.LocalizarCodigo(xCodigoLancamentoFinanceiro) Then

        xContaDebito = LancamentoFinanceiro.ContaDebito
        xContaCredito = LancamentoFinanceiro.ContaCredito
        xHistoricoPadrao = LancamentoFinanceiro.HistoricoPadrao
'        If PortadorFinanceiro.LocalizarCodigo(xPortadorFinanceiro) Then
'            xContaDebito = PortadorFinanceiro.NumeroContaContabil
'        End If

        If IntegracaoCaixa.LocalizarNome(g_empresa, "BAIXA PAGAR " & pFormaPagamento) Then
            xContaDebito = IntegracaoCaixa.ContaDebito
            xContaCredito = IntegracaoCaixa.ContaCredito
            xHistoricoPadrao = IntegracaoCaixa.HistoricoPadrao()
        End If
        MovimentoFinanceiro.Empresa = g_empresa
        MovimentoFinanceiro.Data = BaixaPagar.DataPagamento
        MovimentoFinanceiro.NumeroMovimento = 1
        MovimentoFinanceiro.Valor = BaixaPagar.ValorPagamento
        MovimentoFinanceiro.NumeroDocumento = BaixaPagar.NumeroDocumento
        MovimentoFinanceiro.CodigoHistorico = xHistoricoPadrao
        MovimentoFinanceiro.Complemento = BaixaPagar.Complemento
        MovimentoFinanceiro.NumeroContaDebito = xContaDebito
        MovimentoFinanceiro.NumeroContaCredito = xContaCredito
        MovimentoFinanceiro.CodigoPortador = lCxCodigoPortadorFinanceiro
        MovimentoFinanceiro.CodigoUsuario = g_usuario
        MovimentoFinanceiro.Periodo = 1
        MovimentoFinanceiro.DadosInterno = "BAIXA PAGAR|@|" & BaixaPagar.Registro & "|@|"
        MovimentoFinanceiro.CodigoLancamentoFinanceiro = xCodigoLancamentoFinanceiro
        MovimentoFinanceiro.DataDigitacao = Format(Now, "dd/MM/yyyy")
        MovimentoFinanceiro.HoraDigitacao = Format(Now, "HH:mm:ss")
        MovimentoFinanceiro.DataAlteracao = "00:00:00"
        MovimentoFinanceiro.HoraAlteracao = "00:00:00"
        If MovimentoFinanceiro.Incluir Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & MovimentoFinanceiro.Data & " Portador:" & MovimentoFinanceiro.CodigoPortador & " Per:" & MovimentoFinanceiro.Periodo & " Doc:" & MovimentoFinanceiro.NumeroDocumento & " Vlr:" & MovimentoFinanceiro.Valor)
            IncluiMovimentoFinanceiro = True
        Else
            MsgBox "Não foi possível incluir este movimento financeiro!", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Não será possível integrar com o caixa!" & Chr(10) & "Lançamento Financeiro Inexistente:" & xCodigoLancamentoFinanceiro, vbCritical, "Erro de Verificação!"
    End If
    Exit Function
    
FileError:
    MsgBox "ExcluiMovimentoFinanceiro: Erro não identificado!", vbCritical, "Erro Desconhecido!"
End Function
Private Function ExcluiMovimentoFinanceiro() As Boolean
    Dim xRegistroLocalizado As Boolean
    
On Error GoTo FileError
    
    ExcluiMovimentoFinanceiro = False
    xRegistroLocalizado = False
    If lNumeroMovimentoFinanceiro > 0 Then
        If MovimentoFinanceiro.LocalizarCodigo(g_empresa, lCxData, lNumeroMovimentoFinanceiro) Then
            xRegistroLocalizado = True
        Else
            MsgBox "Registro Financeiro não foi localizado para exclusao!", vbCritical, "Erro de Integridade!"
        End If
    Else
        If MovimentoFinanceiro.LocalizarRegistroEspecialDoc(g_empresa, BaixaPagar.DataPagamento, 1, 1, BaixaPagar.Complemento, BaixaPagar.NumeroDocumento, "421030028", "D") Then
            xRegistroLocalizado = True
        Else
            MsgBox "Registro Financeiro não foi localizado para exclusao!", vbCritical, "Erro de Integridade!"
        End If
    End If
    If xRegistroLocalizado Then
        If MovimentoFinanceiro.Excluir(g_empresa, BaixaPagar.DataPagamento, MovimentoFinanceiro.NumeroMovimento) Then
            ExcluiMovimentoFinanceiro = True
        Else
            MsgBox "Não foi possível excluir o registro Financeiro!", vbCritical, "Erro de Integridade!"
        End If
    End If
    Exit Function
    
FileError:
    MsgBox "ExcluiMovimentoFinanceiro: Erro não identificado!", vbCritical, "Erro Desconhecido!"
End Function
Private Function ExcluiMovimentoFinanceiroDiario(ByVal pRegistroContasAPagar As Integer) As Boolean
    Dim xRegistroLocalizado As Boolean
    Dim xMovimentoTesouraria As New CadastroDLL.cFinMovimentoTesouraria
    Dim SaldoTesouraria As New CadastroDLL.cFinSaldoTesouraria

    
On Error GoTo FileError
    
    ExcluiMovimentoFinanceiroDiario = False
    xRegistroLocalizado = False
    If xMovimentoTesouraria.LocalizarRegistroContaAPagar(g_empresa, pRegistroContasAPagar) Then
        If xMovimentoTesouraria.Excluir(g_empresa, xMovimentoTesouraria.Data, xMovimentoTesouraria.NumeroMovimento) Then
            ExcluiMovimentoFinanceiroDiario = True
            If Not SaldoTesouraria.AlterarSaldo(g_empresa, xMovimentoTesouraria.CodigoTipoMovimento, xMovimentoTesouraria.Data, xMovimentoTesouraria.ValorSaida, True) Then
                MsgBox "ExcluiMovimentoFinanceiroDiario: Não foi possível atualizar saldo do financeiro diário!", vbCritical, "Erro Desconhecido!"
            End If
        Else
            MsgBox "Não foi possível excluir o registro Financeiro Diário!", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Registro Financeiro Diário não foi localizado para exclusão!", vbCritical, "Erro de Integridade!"
    End If
    
    Exit Function
    
FileError:
    MsgBox "ExcluiMovimentoFinanceiroDiario: Erro não identificado!", vbCritical, "Erro Desconhecido!"
End Function
Private Sub LimpaTela()
    lbl_registro.Caption = ""
    lbl_nome_fornecedor.Caption = ""
    lbl_data_emissao.Caption = ""
    lbl_numero_documento.Caption = ""
    lbl_data_vencimento.Caption = ""
    lbl_valor_vencimento.Caption = ""
    lbl_local_cobranca.Caption = ""
    lbl_conta.Caption = ""
    lbl_complemento.Caption = ""
    msk_data_pagamento = "__/__/____"
    txt_valor_pagamento = ""
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 10
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Razão Social do Fornecedor"
    MSFlexGrid.ColWidth(i) = 2500
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Vencimento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 6
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Complemento"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Número do Documento"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Loc. Cobrança"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Cod. Conta"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 6
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data de Emissão"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Empresa"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Registro"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Cod. Fornecedor"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub cmd_extornar_Click()
    Dim valor_x As String
    
    Call GravaAuditoria(1, Me.name, 19, "")
    If lbl_registro.Caption <> "" Then
        If (MsgBox("Deseja Realmente Extornar Esta Baixa?", 4 + 32 + 256, "Extorno de Baixa de Conas à Pagar!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt.Venc:" & lbl_data_vencimento.Caption & "Dt.Pg:" & msk_data_pagamento.Text & " Vlr:" & txt_valor_pagamento.Text & " Forn:" & lbl_nome_fornecedor.Caption & " Reg:" & lbl_registro.Caption)
            Call ExcluiMovimentoCaixa
            Call ExcluiMovimentoFinanceiro
            If lIntegracaoComFinanceiroDiario Then
                Call ExcluiMovimentoFinanceiroDiario(BaixaPagar.Registro)
            End If
            MovContaPagar.Empresa = BaixaPagar.Empresa
            MovContaPagar.Registro = BaixaPagar.Registro
            MovContaPagar.CodigoFornecedor = BaixaPagar.CodigoFornecedor
            MovContaPagar.NomeFornecedor = BaixaPagar.NomeFornecedor
            MovContaPagar.DataEmissao = BaixaPagar.DataEmissao
            MovContaPagar.DataVencimento = BaixaPagar.DataVencimento
            MovContaPagar.Valor = BaixaPagar.Valor
            MovContaPagar.NumeroDocumento = BaixaPagar.NumeroDocumento
            MovContaPagar.LocalCobranca = BaixaPagar.LocalCobranca
            MovContaPagar.CodigoConta = BaixaPagar.CodigoConta
            MovContaPagar.Complemento = BaixaPagar.Complemento
            MovContaPagar.DataDigitacao = BaixaPagar.DataDigitacao
            MovContaPagar.NumeroMovimentoCaixa = BaixaPagar.NumeroMovimentoCaixa
            If BaixaPagar.Excluir(g_empresa, lRegistro) Then
                If Not MovContaPagar.Incluir Then
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
                End If
            Else
                MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
            End If
            AtualizaMSFlexGrid
            If BaixaPagar.LocalizarUltimo(g_empresa) Then
                AtualTela
            Else
                LimpaTela
            End If
        End If
    End If
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
End Sub
Private Sub cmd_ok_Click()
    Dim xFormaPagamento As String
    Dim xCodigoLocalCobranca As Integer
    
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        'Inicio SELECIONA FORMA DE PAGAMENTO
        g_string = "Selecione a forma de pagamento!|@|"
        
        If lIntegracaoComFinanceiroDiario Then
        
           g_string = g_string & ObtenhaTodosLocalCobranca
        
        Else
            g_string = g_string & 2 & "|@|"
            g_string = g_string & "1|@|DINHEIRO|@|"
            g_string = g_string & "1|@|BANCO|@|"
        End If
        
        opcaoGeral.Show 1
        xFormaPagamento = RetiraGString(2)
        xCodigoLocalCobranca = RetiraGString(1)
        
        g_string = ""
        'Fim SELECIONA FORMA DE PAGAMENTO
        If lOpcao = 1 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt.Venc:" & lbl_data_vencimento.Caption & "Dt.Pg:" & msk_data_pagamento.Text & " Vlr:" & txt_valor_pagamento.Text & " Forn:" & lbl_nome_fornecedor.Caption & " Reg:" & Me.lbl_registro.Caption)
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            
            If lIntegracaoComFinanceiroDiario Then
                BaixaPagar.LocalCobranca = xCodigoLocalCobranca
            End If
            
            If BaixaPagar.Incluir Then
                IncluiMovimentoFinanceiro (xFormaPagamento)
                If lIntegracaoComFinanceiroDiario Then
                    Call IncluiMovimentoFinanceiroDiario
                End If
                If Not MovContaPagar.Excluir(g_empresa, lRegistro) Then
                    MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
            End If
            If lCxPeriodo > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
            AtualizaMSFlexGrid
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De: Dt.Venc:" & BaixaPagar.DataVencimento & "Dt.Pg:" & BaixaPagar.DataPagamento & " Vlr:" & BaixaPagar.ValorPagamento & " Forn:" & BaixaPagar.NomeFornecedor & " Reg:" & BaixaPagar.Registro)
            Call GravaAuditoria(1, Me.name, 10, "Para: Dt.Venc:" & lbl_data_vencimento.Caption & "Dt.Pg:" & msk_data_pagamento.Text & " Vlr:" & txt_valor_pagamento.Text & " Forn:" & lbl_nome_fornecedor.Caption & " Reg:" & lbl_registro.Caption)
            Call ExcluiMovimentoCaixa
            Call ExcluiMovimentoFinanceiro
            
            If lIntegracaoComFinanceiroDiario Then
                Call ExcluiMovimentoFinanceiroDiario(lRegistro)
            End If
            
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            Atualtabe_2
            If BaixaPagar.Alterar(g_empresa, lRegistro) Then
                IncluiMovimentoFinanceiro (xFormaPagamento)
                If lIntegracaoComFinanceiroDiario Then
                    Call IncluiMovimentoFinanceiroDiario
                End If
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade."
            End If
            If lCxPeriodo > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
        End If
        lOpcao = 0
        If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        Else
            MsgBox "Não foi possível localiar o registro!", vbInformation, "Erro de Integridade."
        End If
        MSFlexGrid.SetFocus
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_baixa_pagar.Name, "Fornecedoro"
    Exit Sub
End Sub
Private Function ObtenhaTodosLocalCobranca() As String
    Dim rsLocalCobranca As ADODB.Recordset
    Dim xSQL As String
    Dim xRetorno As String

    On Error GoTo FileError
        
        Set rsLocalCobranca = New ADODB.Recordset
        xRetorno = ""
        xSQL = "SELECT Codigo, Nome, CodigoFinTipoMovimento, Empresa FROM Local_Cobrancas WHERE Empresa = " & g_empresa
        
        Set rsLocalCobranca = Conectar.RsConexao(xSQL)
        
        With rsLocalCobranca
        
            If .RecordCount > 0 Then
                xRetorno = .RecordCount & "|@|"
                Do Until .EOF
                    xRetorno = xRetorno & rsLocalCobranca("Codigo").Value & "|@|" & UCase(rsLocalCobranca("Nome").Value) & "|@|"
                    .MoveNext
                Loop
            End If
        End With
        
        ObtenhaTodosLocalCobranca = xRetorno
        
        Exit Function

FileError:
    MsgBox "Erro ao processar Baixa Duplicata Receber", vbInformation, "ProcessaBaixaDuplicataReceber"
    Exit Function
End Function

Private Sub ProcessaBaixaContasPagar()
    Dim xData As Date
    Dim rsBaixaPagar As ADODB.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    Set rsBaixaPagar = New ADODB.Recordset
    xSQL = "SELECT Registro, Data_Pagamento FROM Baixa_Pagar WHERE Empresa = " & g_empresa & " AND Data_Pagamento >= " & xData & " AND [Despesa de Caixa] = False" & " ORDER BY Data_Pagamento"
    Set rsBaixaPagar = Conectar.RsConexao(xSQL)
    With rsBaixaPagar
        If .RecordCount > 0 Then
            Do Until .EOF
                If BaixaPagar.LocalizarCodigo(g_empresa, rsBaixaPagar("Registro").Value) Then
                    AtualTela
                    If Not IncluiMovimentoCaixaAntesBaixa Then
                        MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                    Else
                        If Not IncluiMovimentoCaixa Then
                            MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                        Else
                            BaixaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                            BaixaPagar.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
                            If BaixaPagar.Alterar(g_empresa, rsBaixaPagar("Registro").Value) = False Then
                                MsgBox "Não foi possível alterar este registro de baixa de contas a pagar!", vbInformation, "Erro de Integridade."
                            End If
                        End If
                    End If
                End If
                .MoveNext
            Loop
        Else
            MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
        End If
    End With
    rsBaixaPagar.Close
    Set rsBaixaPagar = Nothing
    MsgBox "Processamento Concluído!"
    Exit Sub
    
FileError:
    MsgBox "Erro ao processar Baixa Duplicata Receber", vbInformation, "ProcessaBaixaDuplicataReceber"
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_pagamento.Text) Then
        MsgBox "Informe a data de pagamento.", vbInformation, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf Not fValidaValor(txt_valor_pagamento.Text) > 0 Then
        MsgBox "Informe o valor do pagamento.", vbInformation, "Atenção!"
        txt_valor_pagamento.SetFocus
    ElseIf Not ValidaCamposFinanceiro Then
    Else
        ValidaCampos = True
    End If
End Function
Private Function ValidaCamposFinanceiro() As Boolean
    ValidaCamposFinanceiro = False
    If lCxPeriodo = 0 Then
        ValidaCamposFinanceiro = True
        Exit Function
    Else
        If CDate(msk_data_pagamento.Text) <> lCxData Then
            MsgBox "A data de pagamento não aceita." & vbCrLf & "Deve ser igual a do movimento financeiro." & vbCrLf & "Ou seja:" & Format(lCxData, "dd/mm/yyyy"), vbInformation, "Atenção!"
            msk_data_pagamento.SetFocus
        Else
            ValidaCamposFinanceiro = True
        End If
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_baixa_contas_pagar.Show 1
    If Len(g_string) > 0 Then
        lRegistro = RetiraGString(1)
        If BaixaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        Else
            MsgBox "Não foi possível localiar o registro!", vbInformation, "Erro de Integridade."
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If BaixaPagar.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If BaixaPagar.LocalizarProximo() Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If BaixaPagar.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If flag_baixa_contas_pagar = 0 Then
        DesativaBotoes
        If RetiraGString(1) = "Financeiro" Then
            AjustaCaixaPista
        Else
            If BaixaPagar.LocalizarUltimo(g_empresa) Then
                AtivaBotoes
                AtualTela
            Else
                cmd_sair.Enabled = True
            End If
            MSFlexGrid.SetFocus
        End If
    Else
        flag_baixa_contas_pagar = 0
    End If
End Sub
Private Sub Form_Deactivate()
    flag_baixa_contas_pagar = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Call GravaAuditoria(1, Me.name, 1, "")
    Screen.MousePointer = 1
    CentraForm Me
    
    AtualizaMSFlexGrid
    lOpcao = 0
    lCxPeriodo = 0
    lNumeroMovimentoFinanceiro = 0
    lIntegracaoComFinanceiroDiario = False
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "CONTAS A PAGAR:INTEGRADO COM FINANCEIRO") Then
        lIntegracaoComFinanceiroDiario = ConfiguracaoDiversa.Verdadeiro
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_Click()
    lRegistro = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 9))
    If lRegistro > 0 Then
        DesativaBotoes
        frmBaixa.Enabled = True
        MostraVencimento
        cmd_alterar.Enabled = False
        cmd_extornar.Enabled = False
        cmd_sair.Enabled = False
        cmd_ok.Visible = True
        cmd_cancelar.Visible = True
    End If
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        lRegistro = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 9))
        If lRegistro > 0 Then
            DesativaBotoes
            frmBaixa.Enabled = True
            MostraVencimento
            cmd_alterar.Enabled = False
            cmd_extornar.Enabled = False
            cmd_sair.Enabled = False
            cmd_ok.Visible = True
            cmd_cancelar.Visible = True
            msk_data_pagamento.SetFocus
        End If
    End If
End Sub
Private Sub msk_data_pagamento_GotFocus()
    If msk_data_pagamento.Text = "__/__/____" Then
        If lCxPeriodo > 0 Then
            msk_data_pagamento.Text = Format(lCxData, "dd/mm/yyyy")
        Else
            msk_data_pagamento.Text = Format(lbl_data_vencimento.Caption, "dd/mm/yyyy")
        End If
    End If
End Sub
Private Sub msk_data_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txt_valor_pagamento_GotFocus()
    If fValidaValor2(txt_valor_pagamento.Text) = 0 Then
        txt_valor_pagamento = Format(lbl_valor_vencimento.Caption, "###,###,##0.00")
    End If
End Sub
Private Sub txt_valor_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        'SendKeys "{TAB}"
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_pagamento_LostFocus()
    txt_valor_pagamento = Format(txt_valor_pagamento, "###,##0.00")
End Sub
