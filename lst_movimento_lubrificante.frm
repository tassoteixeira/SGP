VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_movimento_lubrificante 
   Caption         =   "Emissão do Movimento de Diversos (Óleo/Borr.)"
   ClientHeight    =   5295
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "lst_movimento_lubrificante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_movimento_lubrificante.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Visualiza movimentação de diversos (óleos/borracharia)."
      Top             =   4380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_movimento_lubrificante.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Imprime movimentação de diversos (óleos/borracharia)."
      Top             =   4380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_movimento_lubrificante.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4380
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDataInicial 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   8
         Top             =   660
         Width           =   1095
      End
      Begin VB.CheckBox chkComissaoProduto 
         Caption         =   "Comissão conforme cadastro de produto"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CheckBox chkPeriodoInicial 
         Caption         =   "Apenas para a data inicial"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   3540
         Width           =   3375
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_lubrificante.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_lubrificante.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6060
         Picture         =   "lst_movimento_lubrificante.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CheckBox chk_comissao 
         Caption         =   "&Imprimir Comissão"
         Height          =   255
         Left            =   4740
         TabIndex        =   22
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox cbo_funcionario 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3120
         Width           =   4875
      End
      Begin VB.CheckBox chk_acumulado 
         Caption         =   "Acumular Vendas"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cbo_produto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2340
         Width           =   4875
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1920
         Width           =   4875
      End
      Begin VB.ComboBox cboTipoSubEstoque 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Comissão"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Período Inicial"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Tip&o de de relatório"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "&Tipo do Sub-Estoque"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3900
         TabIndex        =   12
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3900
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_movimento_lubrificante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de variáveis padrão para relatório
Dim l_quantidade As Currency
Dim l_valor As Currency
Dim l_hist_ch_predatado As Currency
Dim l_hist_ch_vista As Currency
Dim l_hist_dinheiro As Currency
Dim l_hist_nota_firma As Currency
Dim l_hist_total As Currency
Dim l_dif_caixa As Currency
Dim l_nome_funcionario As String
Dim lCodigoProduto As Long

Dim lQtdComposicao As Integer
Dim lTotalComposicao As Currency
Dim lValorComposicao(0 To 30) As Currency
Dim lNomeComposicao(0 To 30) As String
Dim lTotalComissao As Currency


Dim lSQL As String
Dim rs As New adodb.Recordset
Dim rsMovComposicaoCaixa As New adodb.Recordset
Dim rsMovLubrificante As New adodb.Recordset
Dim rsProduto As New adodb.Recordset


Private Estoque As New cEstoque
Private Funcionario As New cFuncionario
Private MovimentoComposicaoCaixa As New cMovimentoComposicaoCaixa
Private Produto As New cProduto
Private TabelaPremiacao As New cPremiacao
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rs = Nothing
    Set rsMovComposicaoCaixa = Nothing
    
    Set Estoque = Nothing
    Set Funcionario = Nothing
    Set MovimentoComposicaoCaixa = Nothing
    Set Produto = Nothing
    Set TabelaPremiacao = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    
    l_nome_funcionario = ""
    lQtdComposicao = 0
    lTotalComposicao = 0
    lTotalComissao = 0
    For i = 0 To 30
        lValorComposicao(i) = 0
        lNomeComposicao(i) = ""
    Next
    
    l_quantidade = 0
    l_valor = 0
    l_hist_ch_predatado = 0
    l_hist_ch_vista = 0
    l_hist_dinheiro = 0
    l_hist_nota_firma = 0
    l_hist_total = 0
    l_dif_caixa = 0
End Sub
Private Sub PreparaDatas()
    Dim x_data As Date
    Dim x_data_teste As String
    x_data = CDate("01/" & Month(g_data_def) & "/" & Year(g_data_def))
    If Month(x_data) = 1 Then
        x_data = CDate("01/12/" & Year(g_data_def) - 1)
    Else
        x_data = CDate("01/" & Month(g_data_def) - 1 & "/" & Year(g_data_def))
    End If
    x_data_teste = "32/" & Format(x_data, "mm") & "/" & Year(x_data)
    Do Until IsDate(x_data_teste)
        x_data_teste = Val(Mid(x_data_teste, 1, 2)) - 1 & Mid(x_data_teste, 3, 8)
    Loop
    txtDataInicial.Text = x_data
    txtDataFinal.Text = CDate(x_data_teste)
End Sub
Private Sub Relatorio()
    Dim flag_imprime As Boolean
    
    flag_imprime = False
    ZeraVariaveis
    'Verifica movimento
    If Not g_caixa_unificado Then
        If chk_comissao.Value = 0 Then
            If Not MovimentoComposicaoCaixa.ExisteRegistroData(g_empresa, CDate(txtDataInicial.Text), Val(cbo_periodo_i.Text), 1, cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)) Then
'                If (MsgBox("Histórico não cadastrado!" & Chr(10) & "Deseja continuar?", vbYesNo + vbDefaultButton2, "Erro de integridade!")) = 7 Then
'                    txtdatainicial.SetFocus
'                    Exit Sub
'                End If
            End If
        End If
    End If
    If chk_comissao.Value = 0 Then
        CalculaHistorico
    End If
    
    
    lSQL = ""
    lSQL = lSQL & " SELECT Movimento_Lubrificante.[Codigo do Produto2], Movimento_Lubrificante.[Valor Venda], Movimento_Lubrificante.Quantidade, Movimento_Lubrificante.[Valor Total], Movimento_Lubrificante.Data, Movimento_Lubrificante.Periodo, Movimento_Lubrificante.[Codigo do Funcionario], Produto.Nome, Produto.Unidade"
    lSQL = lSQL & "   FROM Movimento_Lubrificante, Produto"
    lSQL = lSQL & "  WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    If chkPeriodoInicial.Value = 0 Then
        lSQL = lSQL & "    AND Movimento_Lubrificante.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "    AND Movimento_Lubrificante.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "    AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & "    AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    Else
        lSQL = lSQL & "    AND ( Movimento_Lubrificante.Data = " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "          AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & "          OR  Movimento_Lubrificante.Data > " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "          AND Movimento_Lubrificante.Data < " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "          OR  Movimento_Lubrificante.Data = " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "          AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
        lSQL = lSQL & "         ) "
    End If
    lSQL = lSQL & "    AND Movimento_Lubrificante.[Codigo do Produto2] = Produto.Codigo"
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 Then
        lSQL = lSQL & "    AND Movimento_Lubrificante.[Codigo do Funcionario] = " & cbo_funcionario.ItemData(cbo_funcionario.ListIndex)
    End If
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
    If cbo_produto.ItemData(cbo_produto.ListIndex) > 0 Then
        lSQL = lSQL & "    AND Movimento_Lubrificante.[Codigo do Produto2] = " & cbo_produto.ItemData(cbo_produto.ListIndex)
    End If
    If chk_comissao.Value = 1 Then
        lSQL = lSQL & "    AND Produto.Comissao = " & preparaBooleano(True)
    End If
    If cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex) > 0 Then
        lSQL = lSQL & "    AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    End If
    lSQL = lSQL & "  ORDER BY Produto.Nome, Movimento_Lubrificante.Data, Movimento_Lubrificante.Periodo, Movimento_Lubrificante.[Codigo do Funcionario]"
    
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        If chk_acumulado.Value = 0 Then
            LoopMovimentoLubrificante
            ImpDados
        Else
            LoopMovimentoLubrificanteAcumulado
            ImpDados
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    
    
    
'    With tbl_movimento_lubrificante
'        .Index = "id_data"
'        .Seek ">=", g_empresa, CDate(txtdatainicial.Text), cbo_periodo_i.Text, cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex) - 1, cbo_produto.ItemData(cbo_produto.ListIndex), cbo_funcionario.ItemData(cbo_funcionario.ListIndex)
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or !Data > CDate(txtdatafinal.Text) Then
'                    Exit Do
'                End If
'                If !Data <= CDate(txtdatafinal.Text) And !Periodo <= Val(cbo_periodo_f.Text) And (![Codigo do Funcionario] = Val(cbo_funcionario.ItemData(cbo_funcionario.ListIndex)) Or cbo_funcionario.ListIndex = 0) Then
'                    flag_imprime = True
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
    
    
    If chk_acumulado Then
        If cbo_funcionario.ListIndex < (cbo_funcionario.ListCount - 1) Then
            cbo_funcionario.ListIndex = cbo_funcionario.ListIndex + 1
            cbo_funcionario.SetFocus
        End If
    Else
        cmd_sair.SetFocus
    End If
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    If l_valor <> 0 Then
        ImpTotal
        If cbo_funcionario.ListIndex <> -1 Then
            If Not g_caixa_unificado Then
                If chk_comissao.Value = 0 Then
                    ImpHistorico
                End If
            Else
                ImpFim
            End If
        End If
        If chk_comissao.Value = 1 Then
            ImpPremiacao
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório da Venda de Produtos|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopMovimentoLubrificante()
    'loop movimento dos lubrificantes
    Dim i As Integer
    Dim x_linha As String
    Dim x_valor As Currency
    
    With rsMovLubrificante
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 60 Then
                x_linha = String(137, "-")
                Mid(x_linha, 1, 1) = "+"
                Mid(x_linha, 12, 1) = "+"
                Mid(x_linha, 55, 1) = "+"
                Mid(x_linha, 61, 1) = "+"
                Mid(x_linha, 80, 1) = "+"
                Mid(x_linha, 96, 1) = "+"
                Mid(x_linha, 115, 1) = "+"
                Mid(x_linha, 126, 1) = "+"
                Mid(x_linha, 130, 1) = "+"
                Mid(x_linha, 137, 1) = "+"
                Mid(x_linha, 14, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            'Le tabela auxiliar
            Call ImpDet(![Codigo do Produto2], !Nome, !Unidade, ![Valor Venda], !Quantidade, ![Valor Total], !Data, !Periodo, ![Codigo do Funcionario])
            l_valor = l_valor + ![Valor Total]
            l_quantidade = l_quantidade + !Quantidade
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopMovimentoLubrificanteAcumulado()
    'loop movimento dos lubrificantes
    Dim xLinha As String
    Dim xQuantidade As Currency
    Dim xNome As String
    Dim xUnidade As String
    Dim xValorTotal As Currency
    Dim xValorVenda As Currency
    Dim xData As Date
    Dim xPeriodo As String
    Dim xCodigoFuncionario As Integer
    
    With rsMovLubrificante
        Do Until .EOF
            If lCodigoProduto <> ![Codigo do Produto2] Then
                If xValorTotal > 0 Then
                    If lPagina = 0 Then
                        ImpCab
                    End If
                    'xValorVenda = l_valor / l_quantidade
                    If Estoque.LocalizarCodigo(g_empresa, lCodigoProduto) Then
                        If Estoque.PrecoVenda > 0 Then
                            xValorVenda = Estoque.PrecoVenda
                        End If
                    End If
                    If Me.chkComissaoProduto.Value = 1 Then
                        If Produto.LocalizarCodigo(lCodigoProduto) Then
                            lTotalComissao = lTotalComissao + (xValorTotal * Produto.PercentualComissao / 100)
                        End If
                        Call ImpDet(lCodigoProduto, xNome, xUnidade, xValorVenda, xQuantidade, xValorTotal, xData, xPeriodo, Produto.PercentualComissao)
                    Else
                        Call ImpDet(lCodigoProduto, xNome, xUnidade, xValorVenda, xQuantidade, xValorTotal, xData, xPeriodo, xCodigoFuncionario)
                    End If
                    xQuantidade = 0
                    xValorTotal = 0
                    If lLinha >= 60 Then
                        xLinha = String(137, "-")
                        Mid(xLinha, 1, 1) = "+"
                        Mid(xLinha, 12, 1) = "+"
                        Mid(xLinha, 55, 1) = "+"
                        Mid(xLinha, 61, 1) = "+"
                        Mid(xLinha, 80, 1) = "+"
                        Mid(xLinha, 96, 1) = "+"
                        Mid(xLinha, 115, 1) = "+"
                        Mid(xLinha, 126, 1) = "+"
                        Mid(xLinha, 130, 1) = "+"
                        Mid(xLinha, 137, 1) = "+"
                        Mid(xLinha, 14, 22) = " Cerrado Informática. "
                        BioImprime "@Printer.Print " & xLinha
                        BioImprime "@@Printer.NewPage"
                        ImpCab
                    End If
                End If
                lCodigoProduto = ![Codigo do Produto2]
            End If
        
            xValorTotal = xValorTotal + ![Valor Total]
            xNome = !Nome
            xUnidade = !Unidade
            xQuantidade = xQuantidade + !Quantidade
            xData = !Data
            xPeriodo = !Periodo
            xCodigoFuncionario = ![Codigo do Funcionario]
            l_valor = l_valor + ![Valor Total]
            l_quantidade = l_quantidade + !Quantidade
            .MoveNext
        Loop
    End With
    If xValorTotal <> 0 Then
        If lPagina = 0 Then
            ImpCab
        End If
        'xValorVenda = l_valor / l_quantidade
        If Estoque.LocalizarCodigo(g_empresa, lCodigoProduto) Then
            If Estoque.PrecoVenda > 0 Then
                xValorVenda = Estoque.PrecoVenda
            End If
        End If
        If Me.chkComissaoProduto.Value = 1 Then
            If Produto.LocalizarCodigo(lCodigoProduto) Then
                lTotalComissao = lTotalComissao + (xValorTotal * Produto.PercentualComissao / 100)
            End If
            Call ImpDet(lCodigoProduto, xNome, xUnidade, xValorVenda, xQuantidade, xValorTotal, xData, xPeriodo, Produto.PercentualComissao)
        Else
            Call ImpDet(lCodigoProduto, xNome, xUnidade, xValorVenda, xQuantidade, xValorTotal, xData, xPeriodo, xCodigoFuncionario)
        End If
        xQuantidade = 0
        xValorTotal = 0
    End If

End Sub
Private Sub CalculaHistorico()
    Dim i As Integer
    
    
    
    If txtDataInicial.Text = txtDataFinal.Text And cbo_periodo_i.Text = cbo_periodo_f.Text Then  ' And txt_ilha_i.Text = txt_ilha_f.Text Then
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "   SELECT Movimento_Composicao_Caixa.[Codigo do Funcionario],"
        lSQL = lSQL & "          Funcionario.Nome"
        lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Funcionario"
        lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data = " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo = " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] = " & 1 'Val(txt_ilha_i.Text)
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 2
        lSQL = lSQL & "      AND Funcionario.Empresa = " & g_empresa
        lSQL = lSQL & "      AND Funcionario.Codigo = Movimento_Composicao_Caixa.[Codigo do Funcionario]"
        'Abre RecordSet
        Set rsMovComposicaoCaixa = New adodb.Recordset
        Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
        If rsMovComposicaoCaixa.RecordCount > 0 Then
            l_nome_funcionario = rsMovComposicaoCaixa("Nome").Value
        End If
    End If
    
    
    
    
    'loop movimento do historico
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Composicao_Caixa.Ordem,"
    lSQL = lSQL & "          Movimento_Composicao_Caixa.[Codigo da Composicao],"
    lSQL = lSQL & "          SUM(Valor) AS Total,"
    lSQL = lSQL & "          Composicao_Caixa.Nome AS NomeComposicao"
    lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & 1 'Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & 1 'Val(txt_ilha_f.Text)
    If cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex) = 0 Then
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 2
    Else
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    End If
    lSQL = lSQL & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQL = lSQL & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
    i = -1
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        lQtdComposicao = rsMovComposicaoCaixa.RecordCount
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            lValorComposicao(i) = rsMovComposicaoCaixa("Total").Value
            lNomeComposicao(i) = rsMovComposicaoCaixa("NomeComposicao").Value
            lTotalComposicao = lTotalComposicao + rsMovComposicaoCaixa("Total").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
End Sub
Private Sub ImpTotal()
    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 12, 1) = "+"
    Mid(x_linha, 55, 1) = "+"
    Mid(x_linha, 61, 1) = "+"
    Mid(x_linha, 80, 1) = "+"
    Mid(x_linha, 96, 1) = "+"
    Mid(x_linha, 115, 1) = "+"
    Mid(x_linha, 126, 1) = "+"
    Mid(x_linha, 130, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 61, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    Mid(x_linha, 96, 1) = "|"
    Mid(x_linha, 115, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Mid(x_linha, 63, 16) = "TOTAL DAS VENDAS"
    i = Len(Format(l_quantidade, "####,##0.00"))
    Mid(x_linha, 84 + 11 - i, i) = Format(l_quantidade, "####,##0.00")
    i = Len(Format(l_valor, "##,###,##0.00"))
    Mid(x_linha, 101 + 13 - i, i) = Format(l_valor, "##,###,##0.00")
'    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
'    BioImprime "@@y_local = Printer.CurrentY"
    'ImprimeTexto "  ", 1, 2, 2, 1
'    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
'    BioImprime "@@Printer.CurrentY = y_local"
'    BioImprime "@@Printer.Print " & "  "
    BioImprime "@@Printer.FontBold = False"
End Sub
Private Sub ImpFim()
    Dim x_linha As String
    x_linha = "+-----------------------------------------------------------+------------------+---------------+------------------+---------------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpHistorico()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    l_dif_caixa = lTotalComposicao - l_valor
    BioImprime "@Printer.Print " & "+-----------------------------------------------------------+------------------+---------------+------------------+---------------------+"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "|                             RESUMO DO CAIXA                                  |"
    BioImprime "@Printer.Print " & "+----------------------+------------------+------------------------------------+"
    BioImprime "@Printer.Print " & "| NOME DA OPERACAO     |     V A L O R    |                                    |"
    BioImprime "@Printer.Print " & "+----------------------+------------------+------------------------------------+"
    For i = 0 To lQtdComposicao - 1
        x_linha = "|                      |                  |                                    |"
        Mid(x_linha, 3, 20) = lNomeComposicao(i)
        If lValorComposicao(i) > 0 Then
            i2 = Len(Format(lValorComposicao(i), "##,###,##0.00"))
            Mid(x_linha, 29 + 13 - i2, i2) = Format(lValorComposicao(i), "##,###,##0.00")
        End If
        If i = lQtdComposicao - 1 Then
            If l_dif_caixa <> 0 Then
                Mid(x_linha, 45, 20) = "DIFERENÇA DE CAIXA.:"
                i2 = Len(Format(l_dif_caixa, "###,###,##0.00;##,###,##0.00-"))
                Mid(x_linha, 65 + 14 - i2, i2) = Format(l_dif_caixa, "###,###,##0.00;##,###,##0.00-")
            End If
        End If
        BioImprime "@Printer.Print " & x_linha
    Next
    x_linha = "| TOTAL............... |                  | RESP..:                            |"
    If lTotalComposicao > 0 Then
        i = Len(Format(lTotalComposicao, "##,###,##0.00"))
        Mid(x_linha, 29 + 13 - i, i) = Format(lTotalComposicao, "##,###,##0.00")
    End If
    Mid(x_linha, 53, 24) = l_nome_funcionario
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+----------------------+------------------+------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpPremiacao()
    Dim x_linha As String
    Dim i As Integer
    Dim x_indice_alcancado As Currency
    Dim x_percentual_premiacao As Currency
    Dim x_valor_premiacao As Currency
    
    x_indice_alcancado = 0
    x_percentual_premiacao = 0
    x_valor_premiacao = 0
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+--------------------------------------------+--------------+---------------+--+---------------+------------------++--------------------+"
    If chkComissaoProduto.Value = 0 Then
        If TabelaPremiacao.LocalizarCodigo(g_empresa, CDate("01/" & Format(txtDataInicial.Text, "mm") & "/" & Format(txtDataInicial.Text, "yyyy"))) Then
            x_indice_alcancado = l_valor * 100 / TabelaPremiacao.ValorBase
            If x_indice_alcancado >= TabelaPremiacao.PercentualBase1 Then
                x_percentual_premiacao = TabelaPremiacao.PercentualComissao1
                x_valor_premiacao = l_valor * TabelaPremiacao.PercentualComissao1 / 100
            ElseIf x_indice_alcancado >= TabelaPremiacao.PercentualBase2 Then
                x_percentual_premiacao = TabelaPremiacao.PercentualComissao2
                x_valor_premiacao = l_valor * TabelaPremiacao.PercentualComissao2 / 100
            ElseIf x_indice_alcancado >= TabelaPremiacao.PercentualBase3 Then
                x_percentual_premiacao = TabelaPremiacao.PercentualComissao3
                x_valor_premiacao = l_valor * TabelaPremiacao.PercentualComissao3 / 100
            End If
        End If
        x_linha = "| ÍNDICE ALCANCADO.:                         | PERCENTUAL DO PREMIO...:     | VALOR DO PREMIO...:                  |                    |"
        i = Len(Format(x_indice_alcancado, "##0"))
        Mid(x_linha, 41 + 3 - i, i) = Format(x_indice_alcancado, "##0")
        Mid(x_linha, 44, 1) = "%"
        i = Len(Format(x_percentual_premiacao, "##0"))
        Mid(x_linha, 72 + 3 - i, i) = Format(x_percentual_premiacao, "##0")
        Mid(x_linha, 75, 1) = "%"
    ElseIf chkComissaoProduto.Value = 1 Then
        x_valor_premiacao = lTotalComissao
        x_linha = "|                                            |                              | VALOR DO PREMIO...:                  |                    |"
    End If
    i = Len(Format(x_valor_premiacao, "##,###,##0.00"))
    Mid(x_linha, 102 + 13 - i, i) = Format(x_valor_premiacao, "##,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "+--------------------------------------------+------------------------------+--------------------------------------+--------------------+"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(x_codigo_produto As Long, x_nome As String, x_unidade As String, x_valor_venda As Currency, x_quantidade As Currency, x_valor_total As Currency, x_data As Date, x_periodo As String, x_codigo_funcionario As Integer)
    Dim x_linha As String
    Dim i As Integer
    
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 12, 1) = "|"
    Mid(x_linha, 55, 1) = "|"
    Mid(x_linha, 61, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    Mid(x_linha, 96, 1) = "|"
    Mid(x_linha, 115, 1) = "|"
    Mid(x_linha, 126, 1) = "|"
    Mid(x_linha, 130, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    i = Len(Format(x_codigo_produto, "#000"))
    Mid(x_linha, 5 + 4 - i, i) = Format(x_codigo_produto, "#000")
    Mid(x_linha, 14, 40) = x_nome
    Mid(x_linha, 57, 3) = x_unidade
    i = Len(Format(x_valor_venda, "##,###,##0.00"))
    Mid(x_linha, 66 + 13 - i, i) = Format(x_valor_venda, "##,###,##0.00")
    i = Len(Format(x_quantidade, "####,##0.00"))
    Mid(x_linha, 84 + 11 - i, i) = Format(x_quantidade, "####,##0.00")
    i = Len(Format(x_valor_total, "##,###,##0.00"))
    Mid(x_linha, 101 + 13 - i, i) = Format(x_valor_total, "##,###,##0.00")
    Mid(x_linha, 116, 10) = Format(x_data, "dd/mm/yyyy")
    Mid(x_linha, 128, 1) = x_periodo
    i = Len(Format(x_codigo_funcionario, "#00"))
    Mid(x_linha, 132 + 3 - i, i) = Format(x_codigo_funcionario, "#00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "| VENDAS DE PRODUTOS                                       Goiânia, " & txtDataEmissao.Text & " |"
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = txtDataInicial.Text
    Mid(x_linha, 42, 10) = txtDataFinal.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_linha, 29, 1) = cbo_periodo_i
    Mid(x_linha, 49, 1) = cbo_periodo_f
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| FUNCIONARIO.............:                                                    |"
    Mid(x_linha, 29, 3) = Format(cbo_funcionario.ItemData(cbo_funcionario.ListIndex), "000")
    Mid(x_linha, 33, 40) = cbo_funcionario
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| GRUPO...................:                                                    |"
    Mid(x_linha, 29, 40) = cbo_grupo
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PRODUTO.................:                                                    |"
    Mid(x_linha, 29, 40) = cbo_produto
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| TIPO DO SUB-ESTOQUE.....:                                                    |"
    Mid(x_linha, 29, 40) = cboTipoSubEstoque
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+------------------------------------------+-----+------------------+---------------+------------------+----------+---+------+"
    x_linha = "|  CODIGO  | DISCRIMINACAO DOS PRODUTOS               | UN  |  PRECO DE VENDA  |   QUANTIDADE  |  TOTAL DE VENDA  |DATA SAIDA|PER| FUNC.|"
    If Me.chkComissaoProduto.Value = 1 Then
        Mid(x_linha, 131, 6) = "PERCEN"
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "+----------+------------------------------------------+-----+------------------+---------------+------------------+----------+---+------+"
End Sub
Private Sub cbo_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_produto.SetFocus
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoSubEstoque.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_acumulado.SetFocus
    End If
End Sub
Private Sub cboTipoSubEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
Private Sub chk_acumulado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_comissao.SetFocus
    End If
End Sub
Private Sub chk_comissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_funcionario.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = txtDataEmissao.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        txtDataEmissao.Text = RetiraGString(1)
        txtDataInicial.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = txtDataFinal.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
    Else
        txtDataFinal.Text = RetiraGString(1)
    End If
    g_string = ""
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = txtDataInicial.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
    Call SelecionaImpressoraPadrao("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        'If SelecionaImpressoraEpson(Me) Then
            If chk_comissao.Value = 1 Then
                If Not TabelaPremiacao.LocalizarCodigo(g_empresa, CDate("01/" & Format(txtDataInicial.Text, "mm") & "/" & Format(txtDataInicial.Text, "yyyy"))) Then
                    If (MsgBox("Não existe o registro " & Format(txtDataInicial.Text, "mm") & "/" & Format(txtDataInicial.Text, "yyyy") & " cadastrado na tabela de premiação." & Chr(10) & "A premiação não será calculada." & Chr(10) & "Imprime com a premiação zerada?", 4 + 32 + 256, "Erro de Verificação!")) = 7 Then
                        cmd_sair.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        'End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    ElseIf Not IsDate(txtDataInicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        txtDataInicial.SetFocus
    ElseIf Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(txtDataInicial.Text) & ".", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cboTipoSubEstoque.ListIndex = -1 Then
        MsgBox "Selecione o tipo do Sub-Estoque.", vbInformation, "Atenção!"
        cboTipoSubEstoque.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", vbInformation, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf cbo_funcionario.ListIndex = -1 Then
        MsgBox "Selecione o funcionario.", vbInformation, "Atenção!"
        cbo_funcionario.SetFocus
    ElseIf cbo_funcionario.ListIndex < 1 And chk_comissao Then
        MsgBox "Para tirar comissão, tem que selecionar um funcionario.", vbInformation, "Atenção!"
        cbo_funcionario.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        Call SelecionaImpressoraPadrao("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        'If SelecionaImpressoraEpson(Me) Then
            If chk_comissao.Value = 1 And Me.chkComissaoProduto.Value = 0 Then
                If Not TabelaPremiacao.LocalizarCodigo(g_empresa, CDate("01/" & Format(txtDataInicial.Text, "mm") & "/" & Format(txtDataInicial.Text, "yyyy"))) Then
                    If (MsgBox("Não existe o registro " & Format(txtDataInicial.Text, "mm") & "/" & Format(txtDataInicial.Text, "yyyy") & " cadastrado na tabela de premiação." & Chr(10) & "A premiação não será calculada." & Chr(10) & "Imprime com a premiação zerada?", 4 + 32 + 256, "Erro de Verificação!")) = 7 Then
                        cmd_sair.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        'End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(txtDataEmissao.Text) Then
        txtDataEmissao.Text = Format(Date, "dd/mm/yyyy")
        txtDataInicial.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        txtDataFinal.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        cboTipoSubEstoque.ListIndex = 0
        cbo_grupo.ListIndex = 0
        cbo_produto.ListIndex = 0
        chk_acumulado.Value = 0
        chk_comissao.Value = 0
        cbo_funcionario.ListIndex = 0
        cbo_periodo_i.SetFocus
        If g_nivel_acesso = 4 Then
            PreparaDatas
            cbo_periodo_f.ListIndex = 3
            cboTipoSubEstoque.ListIndex = 0
            chk_acumulado.Value = 1
            chk_comissao.Value = 1
            cbo_funcionario.SetFocus
        End If
        chkPeriodoInicial.Value = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    PreencheCboGrupo
    PreencheCboProduto
    PreencheCboFuncionario
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo_i.Clear
    cbo_periodo_f.Clear
    cbo_periodo_i.AddItem 1
    cbo_periodo_f.AddItem 1
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 1
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 1
    cbo_periodo_i.AddItem 2
    cbo_periodo_f.AddItem 2
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 2
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 2
    cbo_periodo_i.AddItem 3
    cbo_periodo_f.AddItem 3
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 3
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 3
    cbo_periodo_i.AddItem 4
    cbo_periodo_f.AddItem 4
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 4
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 4
End Sub
Private Sub PreencheCboTipoMovimento()
    Dim rstTipoSubEstoque As adodb.Recordset
    cboTipoSubEstoque.Clear
    cboTipoSubEstoque.AddItem "Todos os Tipos de Sub-Estoque"
    Set rstTipoSubEstoque = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoSubEstoque ORDER BY Codigo")
    Do Until rstTipoSubEstoque.EOF
        cboTipoSubEstoque.AddItem rstTipoSubEstoque!Codigo & " " & rstTipoSubEstoque!Nome
        cboTipoSubEstoque.ItemData(cboTipoSubEstoque.NewIndex) = rstTipoSubEstoque!Codigo
        rstTipoSubEstoque.MoveNext
    Loop
    rstTipoSubEstoque.Close
    Set rstTipoSubEstoque = Nothing
End Sub
Private Sub PreencheCboGrupo()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cbo_grupo.AddItem rs("Nome").Value
            cbo_grupo.ItemData(cbo_grupo.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub PreencheCboProduto()
    cbo_produto.Clear
    
    cbo_produto.AddItem "Todos os Produtos"
    cbo_produto.ItemData(cbo_produto.NewIndex) = 0
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    Set rsProduto = Conectar.RsConexao(lSQL)
    If rsProduto.RecordCount > 0 Then
        Do Until rsProduto.EOF
                cbo_produto.AddItem rsProduto!Nome
                cbo_produto.ItemData(cbo_produto.NewIndex) = rsProduto!Codigo
            rsProduto.MoveNext
        Loop
    End If
    rsProduto.Close
    Set rsProduto = Nothing
End Sub
Private Sub PreencheCboFuncionario()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Funcionario"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Situacao = " & Chr(39) & "A" & Chr(39)
    lSQL = lSQL & "      AND Periodo < " & 5
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    cbo_funcionario.Clear
    cbo_funcionario.AddItem "Todos os Funcionários"
    cbo_funcionario.ItemData(cbo_funcionario.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cbo_funcionario.AddItem rs("Nome").Value
            cbo_funcionario.ItemData(cbo_funcionario.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub txtDataEmissao_GotFocus()
    txtDataEmissao.Text = fDesmascaraData(txtDataEmissao.Text)
    txtDataEmissao.SelStart = 0
    txtDataEmissao.SelLength = 4
    txtDataEmissao.MaxLength = 8
End Sub
Private Sub txtDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataInicial.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataEmissao_LostFocus()
    txtDataEmissao.MaxLength = 10
    txtDataEmissao.Text = fMascaraData(txtDataEmissao.Text)
End Sub
Private Sub txtDataFinal_GotFocus()
    txtDataFinal.Text = fDesmascaraData(txtDataFinal.Text)
    txtDataFinal.SelStart = 0
    txtDataFinal.SelLength = 2
    txtDataFinal.MaxLength = 8
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataFinal_LostFocus()
    txtDataFinal.MaxLength = 10
    txtDataFinal.Text = fMascaraData(txtDataFinal.Text)
End Sub
Private Sub txtDataInicial_GotFocus()
    txtDataInicial.Text = fDesmascaraData(txtDataInicial.Text)
    txtDataInicial.SelStart = 0
    txtDataInicial.SelLength = 2
    txtDataInicial.MaxLength = 8
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataFinal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataInicial_LostFocus()
    txtDataInicial.MaxLength = 10
    txtDataInicial.Text = fMascaraData(txtDataInicial.Text)
    If IsDate(txtDataInicial.Text) Then
        txtDataFinal.Text = txtDataInicial.Text
    End If
End Sub
