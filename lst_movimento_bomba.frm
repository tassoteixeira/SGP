VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_movimento_bomba 
   Caption         =   "Emissão do Movimento de Bombas"
   ClientHeight    =   3135
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   7155
   Icon            =   "lst_movimento_bomba.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   7155
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1260
      Picture         =   "lst_movimento_bomba.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Visualiza movimentação das bombas."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3180
      Picture         =   "lst_movimento_bomba.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime movimentação das bombas."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5100
      Picture         =   "lst_movimento_bomba.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txt_ilha_f 
         Height          =   300
         Left            =   5160
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1500
         Width           =   255
      End
      Begin VB.TextBox txt_ilha_i 
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1500
         Width           =   255
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_bomba.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_bomba.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6300
         Picture         =   "lst_movimento_bomba.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5160
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
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Ilha Final"
         Height          =   300
         Left            =   3840
         TabIndex        =   16
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Ilha Inicial"
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   1275
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
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_movimento_bomba"
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
Const l_ULTIMO_INDICE As Integer = 40
'Fim de variáveis padrão para relatório
Dim l_abertura_bomba(1 To l_ULTIMO_INDICE) As Currency ' (1 To 20)
Dim l_encerrante_bomba(1 To l_ULTIMO_INDICE) As Currency ' (1 To 20)
Dim l_litros_saida(1 To l_ULTIMO_INDICE) As Currency ' (1 To 20)
Dim l_valor_saida(1 To l_ULTIMO_INDICE) As Currency ' (1 To 20)
Dim lLitrosAfericao As Currency
Dim lValorAfericao As Currency
Dim lValeAbastecimentoEmitido As Currency
Dim lCartaFrete As Currency
Dim lSuprimentoCaixa As Currency
Dim lAcrescimoPrecoPersonalizado As Currency
Dim lBombaAvista As Currency
Dim lLitrosA As Currency
Dim lLitrosAA As Currency
Dim lLitrosD As Currency
Dim lLitrosDA As Currency
Dim lLitrosG As Currency
Dim lLitrosGA As Currency
Dim lLitrosL As Currency
Dim lLitrosB As Currency
Dim lLitrosT As Currency
Dim lValorA As Currency
Dim lValorAA As Currency
Dim lValorD As Currency
Dim lValorDA As Currency
Dim lValorG As Currency
Dim lValorGA As Currency
Dim lValorL As Currency
Dim lValorB As Currency
Dim lValorT As Currency
Dim lLucroA As Currency
Dim lLucroAA As Currency
Dim lLucroD As Currency
Dim lLucroDA As Currency
Dim lLucroG As Currency
Dim lLucroGA As Currency
Dim lLucroT As Currency

Dim lQtdComposicao1 As Integer
Dim lQtdComposicao2 As Integer
Dim lQtdComposicao3 As Integer
Dim lTotalComposicao1 As Currency
Dim lTotalComposicao2 As Currency
Dim lTotalComposicao3 As Currency
Dim lValorComposicao1(0 To l_ULTIMO_INDICE) As Currency ' (1 To 30)
Dim lValorComposicao2(0 To l_ULTIMO_INDICE) As Currency ' (1 To 30)
Dim lValorComposicao3(0 To l_ULTIMO_INDICE) As Currency ' (1 To 30)
Dim lNomeComposicao1(0 To l_ULTIMO_INDICE) As String ' (1 To 30)
Dim lNomeComposicao2(0 To l_ULTIMO_INDICE) As String ' (1 To 30)
Dim lNomeComposicao3(0 To l_ULTIMO_INDICE) As String ' (1 To 30)

Dim l_dif_caixa(1 To 3) As Currency
Dim l_nome_funcionario As String

Dim lSQL As String
Private rsMovBomba As New adodb.Recordset
Private rsMovComposicaoCaixa As New adodb.Recordset
Private rsMovLubrificante As New adodb.Recordset
Private rsTabela As New adodb.Recordset

Private Combustivel As New cCombustivel
Private EntradaCombustivel As New cEntradaCombustivel
Private Funcionario As New cFuncionario
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovCheque As New cMovimentoCheque
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoCartaFrete As New cMovimentoCartaFrete
Private MovimentoComposicaoCaixa As New cMovimentoComposicaoCaixa
Private SuprimentoCaixa As New cSuprimentoCaixa
Private rsTotalizador As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsMovComposicaoCaixa = Nothing
    Set Combustivel = Nothing
    Set EntradaCombustivel = Nothing
    Set Funcionario = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovCheque = Nothing
    Set MovCupomFiscalItem = Nothing
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoCartaFrete = Nothing
    Set MovimentoComposicaoCaixa = Nothing
    Set SuprimentoCaixa = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    l_nome_funcionario = ""
    lQtdComposicao1 = 0
    lQtdComposicao2 = 0
    lQtdComposicao3 = 0
    lTotalComposicao1 = 0
    lTotalComposicao2 = 0
    lTotalComposicao3 = 0
    For i = 0 To l_ULTIMO_INDICE
        lValorComposicao1(i) = 0
        lValorComposicao2(i) = 0
        lValorComposicao3(i) = 0
        lNomeComposicao1(i) = ""
        lNomeComposicao2(i) = ""
        lNomeComposicao3(i) = ""
    Next
    
    For i = 1 To l_ULTIMO_INDICE
        l_abertura_bomba(i) = 0
        l_encerrante_bomba(i) = 0
        l_litros_saida(i) = 0
        l_valor_saida(i) = 0
    Next
    lBombaAvista = 0
    lLitrosAfericao = 0
    lValorAfericao = 0
    lValeAbastecimentoEmitido = 0
    lCartaFrete = 0
    lSuprimentoCaixa = 0
    lAcrescimoPrecoPersonalizado = 0
    lLitrosA = 0
    lLitrosAA = 0
    lLitrosD = 0
    lLitrosDA = 0
    lLitrosG = 0
    lLitrosGA = 0
    lLitrosL = 0
    lLitrosB = 0
    lLitrosT = 0
    lValorA = 0
    lValorAA = 0
    lValorD = 0
    lValorDA = 0
    lValorG = 0
    lValorGA = 0
    lValorL = 0
    lValorB = 0
    lValorT = 0
    
    lLucroA = 0
    lLucroAA = 0
    lLucroD = 0
    lLucroDA = 0
    lLucroG = 0
    lLucroGA = 0
    lLucroT = 0
    
    For i = 1 To 3
        l_dif_caixa(i) = 0
    Next
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento
    If Not MovimentoComposicaoCaixa.ExisteRegistroData(g_empresa, CDate(msk_data_i.Text), Val(txt_ilha_i.Text), Val(cbo_periodo_i.Text), 1) Then
        If (MsgBox("Composição não cadastrada!" & Chr(10) & "Deseja continuar?", vbYesNo + vbDefaultButton2, "Erro de integridade!")) = 7 Then
            msk_data_i.SetFocus
            Exit Sub
        End If
    End If
    CalculaHistorico
    CalculaLubrificante
    
    
    
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Data, Periodo, [Numero da Ilha], SubCaixa, Abertura, Encerrante, "
    lSQL = lSQL & "          [Codigo da Bomba], [Quantidade da Saida], [Preco de Venda], [Numero do Tanque], "
    lSQL = lSQL & "          [Preco de Custo], [Tipo de Combustivel]"
    lSQL = lSQL & "     FROM "
    If g_nome_usuario = "L.M.C." Then
        lSQL = lSQL & "Movimento_Bomba_LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        lSQL = lSQL & "Movimento_Bomba_Cupom"
    Else
        lSQL = lSQL & "Movimento_Bomba"
    End If
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    lSQL = lSQL & " ORDER BY Data, Periodo, [Numero da Ilha], SubCaixa, [Codigo da Bomba]"
    'Abre RecordSet
    Set rsMovBomba = Conectar.RsConexao(lSQL)
    If rsMovBomba.RecordCount > 0 Then
        ImpDados
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
    cmd_sair.SetFocus
    If g_nivel_acesso = 4 Or g_usuario = 8 Then
        msk_data_i.Text = Format(CDate(msk_data_i.Text) + 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(msk_data_f.Text) + 1, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub ImpDados()
    Dim xIndiceLoop As Integer
    LoopMovimentoBomba
    CalculaAfericao
    CalculaValeAbastecimentoEmitido
    lCartaFrete = MovimentoCartaFrete.TotalCartaFrete(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), 1)
    lSuprimentoCaixa = SuprimentoCaixa.TotalPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    lAcrescimoPrecoPersonalizado = MovCupomFiscalItem.TotalAcrescimo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), True)
    If lLitrosT > 0 Then
        ImpCab
        Call ImpDet(1, l_abertura_bomba(1), l_encerrante_bomba(1), l_litros_saida(1), l_valor_saida(1), lNomeComposicao1(0), "@N@" & lValorComposicao1(0))
        Call ImpDet(2, l_abertura_bomba(2), l_encerrante_bomba(2), l_litros_saida(2), l_valor_saida(2), lNomeComposicao1(1), "@N@" & lValorComposicao1(1))
        Call ImpDet(3, l_abertura_bomba(3), l_encerrante_bomba(3), l_litros_saida(3), l_valor_saida(3), lNomeComposicao1(2), "@N@" & lValorComposicao1(2))
        Call ImpDet(4, l_abertura_bomba(4), l_encerrante_bomba(4), l_litros_saida(4), l_valor_saida(4), lNomeComposicao1(3), "@N@" & lValorComposicao1(3))
        Call ImpDet(5, l_abertura_bomba(5), l_encerrante_bomba(5), l_litros_saida(5), l_valor_saida(5), lNomeComposicao1(4), "@N@" & lValorComposicao1(4))
        Call ImpDet(6, l_abertura_bomba(6), l_encerrante_bomba(6), l_litros_saida(6), l_valor_saida(6), lNomeComposicao1(5), "@N@" & lValorComposicao1(5))
        Call ImpDet(7, l_abertura_bomba(7), l_encerrante_bomba(7), l_litros_saida(7), l_valor_saida(7), lNomeComposicao1(6), "@N@" & lValorComposicao1(6))
        Call ImpDet(8, l_abertura_bomba(8), l_encerrante_bomba(8), l_litros_saida(8), l_valor_saida(8), lNomeComposicao1(7), "@N@" & lValorComposicao1(7))
        Call ImpDet(9, l_abertura_bomba(9), l_encerrante_bomba(9), l_litros_saida(9), l_valor_saida(9), lNomeComposicao1(8), "@N@" & lValorComposicao1(8))
        Call ImpDet(10, l_abertura_bomba(10), l_encerrante_bomba(10), l_litros_saida(10), l_valor_saida(10), lNomeComposicao1(9), "@N@" & lValorComposicao1(9))
        Call ImpDet(11, l_abertura_bomba(11), l_encerrante_bomba(11), l_litros_saida(11), l_valor_saida(11), lNomeComposicao1(10), "@N@" & lValorComposicao1(10))
        Call ImpDet(12, l_abertura_bomba(12), l_encerrante_bomba(12), l_litros_saida(12), l_valor_saida(12), lNomeComposicao1(11), "@N@" & lValorComposicao1(11))
        Call ImpDet(13, l_abertura_bomba(13), l_encerrante_bomba(13), l_litros_saida(13), l_valor_saida(13), "Responsavel: ", "@A@" & l_nome_funcionario)
        Call ImpDet(14, l_abertura_bomba(14), l_encerrante_bomba(14), l_litros_saida(14), l_valor_saida(14), "Total Geral: ", "@N@" & lTotalComposicao1)
        
       
        For xIndiceLoop = 15 To l_ULTIMO_INDICE
            If l_encerrante_bomba(xIndiceLoop) > 0 Then
                Call ImpDet(xIndiceLoop, l_abertura_bomba(xIndiceLoop), l_encerrante_bomba(xIndiceLoop), l_litros_saida(xIndiceLoop), l_valor_saida(xIndiceLoop), "", "@N@" & 0)
            End If
        Next
        
        
'        If l_encerrante_bomba(15) > 0 Then
'            Call ImpDet(15, l_abertura_bomba(15), l_encerrante_bomba(15), l_litros_saida(15), l_valor_saida(15), "", "@N@" & 0)
'        End If
'        If l_encerrante_bomba(16) > 0 Then
'            Call ImpDet(16, l_abertura_bomba(16), l_encerrante_bomba(16), l_litros_saida(16), l_valor_saida(16), "", "@N@" & 0)
'        End If
'        If l_encerrante_bomba(17) > 0 Then
'            Call ImpDet(17, l_abertura_bomba(17), l_encerrante_bomba(17), l_litros_saida(17), l_valor_saida(17), "", "@N@" & 0)
'        End If
'        If l_encerrante_bomba(18) > 0 Then
'            Call ImpDet(18, l_abertura_bomba(18), l_encerrante_bomba(18), l_litros_saida(18), l_valor_saida(18), "", "@N@" & 0)
'        End If
'        If l_encerrante_bomba(19) > 0 Then
'            Call ImpDet(19, l_abertura_bomba(19), l_encerrante_bomba(19), l_litros_saida(19), l_valor_saida(19), "", "@N@" & 0)
'        End If
'        If l_encerrante_bomba(20) > 0 Then
'            Call ImpDet(20, l_abertura_bomba(20), l_encerrante_bomba(20), l_litros_saida(20), l_valor_saida(20), "", "@N@" & 0)
'        End If
        
        
        
        Call LoopMovimentoLubrificante
        If Not g_caixa_unificado Then
            Call ImpResumoHistoricos
        End If
        Call ImpResumoCombustiveis
        Call ImpResumoChequePreDatado
        Call ImpResumoBaixaDuplicaraReceber
        Call ImpResumoBaixaContasPagar
        Call ImpResumoDiferencaPreco
        Call ImpResumoMedicaoCombustiveis
        If Not g_caixa_unificado Then
            If g_usuario = 8 Then
                Call ImpCodificacaoContabil
            End If
        End If
        If g_caixa_unificado = False Then
            'Call ImpResumoProtocoloEntrega
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Movimento de Bomba/Composição do Caixa|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopMovimentoBomba()
    Dim i As Integer
    Dim xCustoUnitario As Currency
    'loop movimento das bombas
    With rsMovBomba
        Do Until .EOF
            i = ![Codigo da Bomba]
            If l_abertura_bomba(i) = 0 Then
                l_abertura_bomba(i) = !Abertura
            End If
            l_encerrante_bomba(i) = !Encerrante
            l_litros_saida(i) = l_litros_saida(i) + ![Quantidade da Saida]
            l_valor_saida(i) = l_valor_saida(i) + (![Quantidade da Saida] * ![Preco de Venda])
            xCustoUnitario = ![Preco de Custo]
            If msk_data_i.Text = msk_data_f.Text Then
                xCustoUnitario = CustoUltimaEntrada(![Tipo de Combustivel], CDate(msk_data_i.Text))
            End If
            Select Case Trim(![Tipo de Combustivel])
                Case "A"
                    lLitrosA = lLitrosA + ![Quantidade da Saida]
                    lValorA = lValorA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroA = lLucroA + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
                Case "AA"
                    lLitrosAA = lLitrosAA + ![Quantidade da Saida]
                    lValorAA = lValorAA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroAA = lLucroAA + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
                Case "D"
                    lLitrosD = lLitrosD + ![Quantidade da Saida]
                    lValorD = lValorD + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroD = lLucroD + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
                Case "DA"
                    lLitrosDA = lLitrosDA + ![Quantidade da Saida]
                    lValorDA = lValorDA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroDA = lLucroDA + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
                Case "G"
                    lLitrosG = lLitrosG + ![Quantidade da Saida]
                    lValorG = lValorG + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroG = lLucroG + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
                Case "GA"
                    lLitrosGA = lLitrosGA + ![Quantidade da Saida]
                    lValorGA = lValorGA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    lLucroGA = lLucroGA + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
            End Select
            lLucroT = lLucroT + Format(![Quantidade da Saida] * (![Preco de Venda] - xCustoUnitario), "#########0.00")
            If ![Numero do Tanque] = 2 Then
                lBombaAvista = lBombaAvista + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
            End If
            lLitrosT = lLitrosT + ![Quantidade da Saida]
            lValorT = lValorT + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopMovimentoLubrificante()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Movimento_Lubrificante.[Codigo do Produto2], SUM(Movimento_Lubrificante.Quantidade) AS Quantidade, SUM(Movimento_Lubrificante.[Valor Total]) AS [Valor Total], Produto.Nome"
    lSQL = lSQL & "     FROM Movimento_Lubrificante, Produto"
    lSQL = lSQL & "    WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    'lSQl = lSQl & "      AND Movimento_Lubrificante.[Tipo do Movimento] = " & 1
    lSQL = lSQL & "      AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
    lSQL = lSQL & " GROUP BY Produto.Nome, Movimento_Lubrificante.[Codigo do Produto2]"
    lSQL = lSQL & " ORDER BY Produto.Nome, Movimento_Lubrificante.[Codigo do Produto2]"
    'Abre RecordSet
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        ImpCabLubrificante
        rsMovLubrificante.MoveFirst
        Do Until rsMovLubrificante.EOF
            ImpDetLubrificante
            rsMovLubrificante.MoveNext
        Loop
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
End Sub
Function CustoUltimaEntrada(ByVal xTipoCombustivel As String, ByVal xData As Date) As Currency
    CustoUltimaEntrada = 0
    If EntradaCombustivel.LocalizarUltimoCombustivel(g_empresa, xData + 1, xTipoCombustivel) Then
        CustoUltimaEntrada = EntradaCombustivel.ValorLitro
    End If
End Function
Private Sub CalculaAfericao()
    Dim i As Integer
    Dim xValor As Currency
    Dim xCusto As Currency
    'loop movimento de Afericao
    lValorAfericao = 0
    'Busca Litragem
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", "")
    lLitrosA = lLitrosA - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", "")
    lLitrosAA = lLitrosAA - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", "")
    lLitrosD = lLitrosD - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", "")
    lLitrosDA = lLitrosDA - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", "")
    lLitrosG = lLitrosG - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    xValor = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", "")
    lLitrosGA = lLitrosGA - xValor
    lLitrosAfericao = lLitrosAfericao + xValor
    
    'Busca Valor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroA = lLucroA - (xValor - xCusto)
    lValorA = lValorA - xValor
    lValorAfericao = lValorAfericao + xValor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroAA = lLucroAA - (xValor - xCusto)
    lValorAA = lValorAA - xValor
    lValorAfericao = lValorAfericao + xValor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroD = lLucroD - (xValor - xCusto)
    lValorD = lValorD - xValor
    lValorAfericao = lValorAfericao + xValor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroDA = lLucroDA - (xValor - xCusto)
    lValorDA = lValorDA - xValor
    lValorAfericao = lValorAfericao + xValor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroG = lLucroG - (xValor - xCusto)
    lValorG = lValorG - xValor
    lValorAfericao = lValorAfericao + xValor
    xCusto = MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", "")
    xValor = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", "")
    lLucroT = lLucroT - (xValor - xCusto)
    lLucroGA = lLucroGA - (xValor - xCusto)
    lValorGA = lValorGA - xValor
    lValorAfericao = lValorAfericao + xValor
    'Calcula Lucro
    

    
    
    'Calcula Total
    
    lLitrosT = lLitrosT - lLitrosAfericao
    lValorT = lValorT - lValorAfericao
    
    
'    With tbl_movimento_afericao
'        .Seek ">=", g_empresa, CDate(msk_data_i), cbo_periodo_i, 0, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f And !Transferencia = False Then
'                    Select Case Trim(![Tipo de Combustivel])
'                        Case "A"
'                            lLitrosA = lLitrosA - !Quantidade
'                            lValorA = lValorA - Format(![Valor Total], "#########0.00")
'                            lLucroA = lLucroA - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                        Case "AA"
'                            lLitrosAA = lLitrosAA - !Quantidade
'                            lValorAA = lValorAA - Format(![Valor Total], "#########0.00")
'                            lLucroAA = lLucroAA - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                        Case "D"
'                            lLitrosD = lLitrosD - !Quantidade
'                            lValorD = lValorD - Format(![Valor Total], "#########0.00")
'                            lLucroD = lLucroD - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                        Case "DA"
'                            lLitrosDA = lLitrosDA - !Quantidade
'                            lValorDA = lValorDA - Format(![Valor Total], "#########0.00")
'                            lLucroDA = lLucroDA - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                        Case "G"
'                            lLitrosG = lLitrosG - !Quantidade
'                            lValorG = lValorG - Format(![Valor Total], "#########0.00")
'                            lLucroG = lLucroG - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                        Case "GA"
'                            lLitrosGA = lLitrosGA - !Quantidade
'                            lValorGA = lValorGA - Format(![Valor Total], "#########0.00")
'                            lLucroGA = lLucroGA - Format(![Valor Total] - (![Preco de Custo] * !Quantidade), "#########0.00")
'                    End Select
'                    lLitrosT = lLitrosT - !Quantidade
'                    lValorT = lValorT - Format(![Valor Total], "#########0.00")
'                    lLitrosAfericao = lLitrosAfericao + !Quantidade
'                    lValorAfericao = lValorAfericao + ![Valor Total]
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
End Sub
Private Sub CalculaHistorico()
    Dim i As Integer
    
    If msk_data_i.Text = msk_data_f.Text And cbo_periodo_i.Text = cbo_periodo_f.Text And txt_ilha_i.Text = txt_ilha_f.Text Then
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Movimento_Composicao_Caixa.[Codigo do Funcionario],"
        lSQL = lSQL & "       Funcionario.Nome"
        lSQL = lSQL & "  FROM Movimento_Composicao_Caixa, Funcionario"
        lSQL = lSQL & " WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Movimento_Composicao_Caixa.Data = " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND Movimento_Composicao_Caixa.Periodo = " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "   AND Movimento_Composicao_Caixa.[Numero da Ilha] = " & Val(txt_ilha_i.Text)
        lSQL = lSQL & "   AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 1
        lSQL = lSQL & "   AND Funcionario.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Funcionario.Codigo = Movimento_Composicao_Caixa.[Codigo do Funcionario]"
        'Abre RecordSet
        Set rsMovComposicaoCaixa = New adodb.Recordset
        Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
        If rsMovComposicaoCaixa.RecordCount > 0 Then
            l_nome_funcionario = rsMovComposicaoCaixa("Nome").Value
        End If
    End If
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Composicao_Caixa.Ordem,"
    lSQL = lSQL & "          Movimento_Composicao_Caixa.[Codigo da Composicao],"
    lSQL = lSQL & "          SUM(Valor) AS Total,"
    lSQL = lSQL & "          Composicao_Caixa.Nome AS NomeComposicao"
    lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    If g_caixa_unificado Then
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] > " & 0
    Else
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 1
    End If
    lSQL = lSQL & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQL = lSQL & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    lSQL = lSQL & " ORDER BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
    i = -1
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        lQtdComposicao1 = rsMovComposicaoCaixa.RecordCount
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            lValorComposicao1(i) = rsMovComposicaoCaixa("Total").Value
            lNomeComposicao1(i) = rsMovComposicaoCaixa("NomeComposicao").Value
            lTotalComposicao1 = lTotalComposicao1 + rsMovComposicaoCaixa("Total").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
End Sub
Private Sub CalculaLubrificante()
    Dim i As Integer
    
    'loop movimento de lubrificante
    lSQL = ""
    lSQL = lSQL & "   SELECT SUM([Valor Total]) AS ValorTotal"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    'lSQl = lSQl & "      AND [Tipo do Movimento] = " & 1
    'Abre RecordSet
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        If Not IsNull(rsMovLubrificante!ValorTotal) Then
            lValorL = rsMovLubrificante!ValorTotal
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
End Sub
Private Sub CalculaValeAbastecimentoEmitido()
    Dim xSQL As String
    lValeAbastecimentoEmitido = 0
    xSQL = ""
    xSQL = xSQL & "   SELECT SUM(Valor) AS Total"
    xSQL = xSQL & "     FROM Movimento_Vale_Abastecimento_Emitido"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    xSQL = xSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    xSQL = xSQL & "      AND Periodo >= " & Val(cbo_periodo_i.Text)
    xSQL = xSQL & "      AND Periodo <= " & Val(cbo_periodo_f.Text)
    xSQL = xSQL & "      AND [Tipo de Movimento] >= " & 0
    Set rsTotalizador = New adodb.Recordset
    Set rsTotalizador = Conectar.RsConexao(xSQL)
    If Not rsTotalizador.EOF Then
        If Not IsNull(rsTotalizador("Total").Value) Then
            lValeAbastecimentoEmitido = rsTotalizador("Total").Value
        End If
    End If
    rsTotalizador.Close
    xSQL = ""
    xSQL = xSQL & "   SELECT SUM(Valor) AS Total"
    xSQL = xSQL & "     FROM Movimento_Vale_Abastecimento_Recebido"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    xSQL = xSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    xSQL = xSQL & "      AND Periodo >= " & Val(cbo_periodo_i.Text)
    xSQL = xSQL & "      AND Periodo <= " & Val(cbo_periodo_f.Text)
    xSQL = xSQL & "      AND [Tipo de Movimento] >= " & 0
    Set rsTotalizador = New adodb.Recordset
    Set rsTotalizador = Conectar.RsConexao(xSQL)
    If Not rsTotalizador.EOF Then
        If Not IsNull(rsTotalizador("Total").Value) Then
            lValeAbastecimentoEmitido = lValeAbastecimentoEmitido + rsTotalizador("Total").Value
        End If
    End If
    rsTotalizador.Close
    Set rsTotalizador = Nothing
End Sub
Private Sub ImpResumoDiferencaPreco()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim x_valor As Currency
    If lBombaAvista > 0 Then
        x_valor = 0
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
'        x_linha = "| VENDAS DE BOMBAS A VISTA.:                           Cheques A Vista + Dinheiro.:                          Diferenca.:                |"
'        i = Len(Format(lBombaAvista, "###,###,##0.00"))
'        Mid(x_linha, 30 + 14 - i, i) = Format(lBombaAvista, "###,###,##0.00")
'        x_valor = x_valor + lValorComposicao3(1) + lValorComposicao3(2)
'        i = Len(Format(x_valor, "###,###,##0.00"))
'        Mid(x_linha, 85 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
'        x_valor = x_valor - lBombaAvista
'        i = Len(Format(x_valor, "###,###,##0.00"))
'        Mid(x_linha, 122 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
'        BioImprime "@Printer.Print " & x_linha
        
        x_linha = "| Cheques A Vista + Dinheiro.:                                                                                                          |"
        x_valor = x_valor + lValorComposicao3(0) + lValorComposicao3(1)
        i = Len(Format(x_valor, "###,###,##0.00"))
        Mid(x_linha, 32 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
        x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
        If g_usuario = 8 Then
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
        End If
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontName = Draft 10cpi"
    End If
End Sub
Private Sub ImpResumoHistoricos()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Composicao_Caixa.Ordem,"
    lSQL = lSQL & "          Movimento_Composicao_Caixa.[Codigo da Composicao],"
    lSQL = lSQL & "          SUM(Valor) AS Total,"
    lSQL = lSQL & "          Composicao_Caixa.Nome AS NomeComposicao"
    lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 2
    lSQL = lSQL & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQL = lSQL & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    lSQL = lSQL & " ORDER BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
    i = -1
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        lQtdComposicao2 = rsMovComposicaoCaixa.RecordCount
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            lValorComposicao2(i) = rsMovComposicaoCaixa("Total").Value
            lNomeComposicao2(i) = rsMovComposicaoCaixa("NomeComposicao").Value
            lTotalComposicao2 = lTotalComposicao2 + rsMovComposicaoCaixa("Total").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Composicao_Caixa.Ordem,"
    lSQL = lSQL & "          Movimento_Composicao_Caixa.[Codigo da Composicao],"
    lSQL = lSQL & "          SUM(Valor) AS Total,"
    lSQL = lSQL & "          Composicao_Caixa.Nome AS NomeComposicao"
    lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] > " & 0
    lSQL = lSQL & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQL = lSQL & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    lSQL = lSQL & " ORDER BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
    i = -1
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        lQtdComposicao3 = rsMovComposicaoCaixa.RecordCount
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            lValorComposicao3(i) = rsMovComposicaoCaixa("Total").Value
            lNomeComposicao3(i) = rsMovComposicaoCaixa("NomeComposicao").Value
            lTotalComposicao3 = lTotalComposicao3 + rsMovComposicaoCaixa("Total").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    BioImprime "@Printer.Print " & "+----+----------------------------------+-----+------------+------+------------+"
    BioImprime "@Printer.Print " & "|  COMPOSIÇÃO DO CAIXA DE ÓLEO/LUBRIF.  |   COMPOSIÇÃO DOS CAIXAS UNIFICADOS   |"
    BioImprime "@Printer.Print " & "+---------------------------------------+--------------------------------------+"
    For i = 0 To lQtdComposicao3 - 1
        x_linha = "| .....................:                | ....................:                |"
        Mid(x_linha, 3, 20) = lNomeComposicao2(i)
        If lValorComposicao2(i) > 0 Then
            i2 = Len(Format(lValorComposicao2(i), "#,###,##0.00"))
            Mid(x_linha, 28 + 12 - i2, i2) = Format(lValorComposicao2(i), "#,###,##0.00")
        End If
        Mid(x_linha, 43, 20) = lNomeComposicao3(i)
        If lValorComposicao3(i) > 0 Then
            i2 = Len(Format(lValorComposicao3(i), "#,###,##0.00"))
            Mid(x_linha, 68 + 12 - i2, i2) = Format(lValorComposicao3(i), "#,###,##0.00")
        End If
        BioImprime "@Printer.Print " & x_linha
    Next
    x_linha = "| Total................:                | Total Geral.........:                |"
    If lTotalComposicao2 > 0 Then
        i = Len(Format(lTotalComposicao2, "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(lTotalComposicao2, "#,###,##0.00")
    End If
    If lTotalComposicao3 > 0 Then
        i = Len(Format(lTotalComposicao3, "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(lTotalComposicao3, "#,###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpResumoProtocoloEntrega()
    Dim x_linha As String
    Dim xChequeVista As Currency
    Dim xDinheiro As Currency
    Dim i As Integer
    
    
    
    For i = 0 To lQtdComposicao3 - 1
        If UCase(lNomeComposicao3(i)) Like "*CH*" And UCase(lNomeComposicao3(i)) Like "*VISTA*" Then
            xChequeVista = lValorComposicao3(i)
        End If
        If UCase(lNomeComposicao3(i)) Like "*DINHEIRO*" Then
            xDinheiro = lValorComposicao3(i)
        End If
    Next
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "                                                                                "
    BioImprime "@Printer.Print " & "                    PROTOCOLO DE ENTREGA DE VALORES                             "
    BioImprime "@Printer.Print " & "                                                                                "
    
    x_linha = "  Cheques à Vista......:                 _____________________________________  "
    If xChequeVista > 0 Then
        i = Len(Format(xChequeVista, "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(xChequeVista, "#,###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "                                                                                "
    x_linha = "  Dinheiro.............:                 _____________________________________  "
    If xDinheiro > 0 Then
        i = Len(Format(xDinheiro, "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(xDinheiro, "#,###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "                                                                                "
    x_linha = "  Outros...............:                 _____________________________________  "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "                                                                                "
End Sub
Private Sub ImpResumoCombustiveis()
    Dim x_linha As String
    Dim i As Integer
    If g_caixa_unificado Then
        BioImprime "@Printer.Print " & "+--+--------+-+----------++--------+----+---------+--+------------+------------+"
    Else
        BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    End If
    BioImprime "@Printer.Print " & "|COMBUSTÍVEL|    LITROS   |    VALOR    |LUCRO  MEDIO|LUCRO  VENDA| % DO LUCRO |"
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    Call ImpDetCombustivel("ÁLCOOL    ", lLitrosA, lValorA, lLucroA, "A ")
    Call ImpDetCombustivel("ÁLCOOL +  ", lLitrosAA, lValorAA, lLucroAA, "AA")
    Call ImpDetCombustivel("DIESEL    ", lLitrosD, lValorD, lLucroD, "D ")
    Call ImpDetCombustivel("DIESEL +  ", lLitrosDA, lValorDA, lLucroDA, "DA")
    Call ImpDetCombustivel("GASOLINA  ", lLitrosG, lValorG, lLucroG, "G ")
    Call ImpDetCombustivel("GASOLINA +", lLitrosGA, lValorGA, lLucroGA, "GA")
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    Call ImpDetCombustivel("SUB-TOTAL ", lLitrosT, lValorT, lLucroT, "TT")
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    Call ImpDetCombustivel("ÓLEOS/LUBR", lLitrosL, lValorL, 0, "  ")
    Call ImpDetCombustivel("BORRA/LAV.", lLitrosB, lValorB, 0, "  ")
    Call ImpDetCombustivel("AFERIÇÕES", lLitrosAfericao, lValorAfericao, 0, "AF")
    If lValeAbastecimentoEmitido > 0 Then
        Call ImpDetCombustivel("V.ABAST.EMI", 0, lValeAbastecimentoEmitido, 0, "VE")
    End If
    If lAcrescimoPrecoPersonalizado > 0 Then
        Call ImpDetCombustivel("ACRES.PREÇO", 0, lAcrescimoPrecoPersonalizado, 0, "AP")
    End If
    If lCartaFrete > 0 Then
        Call ImpDetCombustivel("CARTA FRETE", 0, lCartaFrete, 0, "CF")
    End If
    If lSuprimentoCaixa > 0 Then
        Call ImpDetCombustivel("SUPR. CAIXA", 0, lSuprimentoCaixa, 0, "CF")
    End If
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    x_linha = "|TOTAL GERAL|             |             | DIFERENÇA DE CAIXA...:               |"
    lLitrosT = lLitrosT + lLitrosL + lLitrosB + lLitrosAfericao
    lValorT = lValorT + lValorL + lValorB + lValorAfericao
    i = Len(Format(lLitrosT, "##,###,##0.0"))
    Mid(x_linha, 15 + 12 - i, i) = Format(lLitrosT, "##,###,##0.0")
    i = Len(Format(lValorT, "#,###,##0.00"))
    Mid(x_linha, 28 + 12 - i, i) = Format(lValorT, "#,###,##0.00")
    l_dif_caixa(1) = lTotalComposicao1 - lValorT
    If g_caixa_unificado Then
        l_dif_caixa(1) = l_dif_caixa(1) + lTotalComposicao2 - lValorL - lValorB - lCartaFrete - lSuprimentoCaixa
    End If
    If l_dif_caixa(1) <> 0 Then
        l_dif_caixa(1) = l_dif_caixa(1) + lValorL - lValeAbastecimentoEmitido - lAcrescimoPrecoPersonalizado - lCartaFrete - lSuprimentoCaixa
        i = Len(Format(l_dif_caixa(1), "#,###,##0.00;####,##0.00-"))
        Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(1), "#,###,##0.00;####,##0.00-")
    End If
    BioImprime "@Printer.Print " & x_linha
    If Not g_caixa_unificado Then
        x_linha = "|           |             |             | DIFERENÇA CAIXA ÓLEOS:               |"
        l_dif_caixa(2) = lTotalComposicao2 - lValorL
        If l_dif_caixa(2) <> 0 Then
            i = Len(Format(l_dif_caixa(2), "#,###,##0.00;####,##0.00-"))
            Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(2), "#,###,##0.00;####,##0.00-")
        End If
        BioImprime "@Printer.Print " & x_linha
        'x_linha = "|           |             |             | DIFERENÇA CAIXA BORR.:               |"
        'l_dif_caixa(3) = lTotalComposicao3 - lValorB
        'If l_dif_caixa(3) <> 0 Then
        '    i = Len(Format(l_dif_caixa(3), "#,###,##0.00;####,##0.00-"))
        '    Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(3), "#,###,##0.00;####,##0.00-")
        'End If
        'BioImprime "@Printer.Print " & x_linha
    End If
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+--------------------------------------+"
End Sub
Private Sub ImpResumoMedicaoCombustiveis()
    Dim x_linha As String
    Dim i As Integer
    If cbo_periodo_i.Text <> cbo_periodo_f.Text Then
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & "+-----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-----------+"
        BioImprime "@Printer.Print " & "|COMBUSTÍVEL| EST.INICIAL | ENTRADAS    | QTD. SAIDAS | EST.ESCRIT. | EST.  FINAL | PERDA/SOBRA | PER.SOB. R$ |LASTRO TANQUE|           |"
        BioImprime "@Printer.Print " & "+-----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-----------+"
        Call ImpDetMedicaoCombustivel("A ", "N")
        Call ImpDetMedicaoCombustivel("A ", "S")
        Call ImpDetMedicaoCombustivel("AA", "N")
        Call ImpDetMedicaoCombustivel("AA", "S")
        Call ImpDetMedicaoCombustivel("D ", "N")
        Call ImpDetMedicaoCombustivel("D ", "S")
        Call ImpDetMedicaoCombustivel("DA", "N")
        Call ImpDetMedicaoCombustivel("DA", "S")
        Call ImpDetMedicaoCombustivel("G ", "N")
        Call ImpDetMedicaoCombustivel("G ", "S")
        Call ImpDetMedicaoCombustivel("GA", "N")
        Call ImpDetMedicaoCombustivel("GA", "S")
        BioImprime "@Printer.Print " & "+-----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-----------+"
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
End Sub
Private Sub ImpResumoChequePreDatado()
    Dim x_linha As String
    Dim i As Integer
    Dim x_total As Currency
    x_total = MovCheque.TotalEmissaoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), cbo_periodo_i.Text, cbo_periodo_f, "0", "P")
    If x_total > 0 Then
        x_linha = "| CHEQUES PRÉ-DATADOS PARA VENCIMENTO EM.:             TOTAL.:                 |"
        Mid(x_linha, 44, 10) = Format(CDate(msk_data_i) + 1, "dd/mm/yyyy")
        i = Len(Format(x_total, "###,###,##0.00"))
        Mid(x_linha, 64 + 14 - i, i) = Format(x_total, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
        x_linha = "+------------------------------------------------------------------------------+"
        If g_usuario = 8 Then
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
        End If
        BioImprime "@Printer.Print " & x_linha
    End If
End Sub
Private Sub ImpResumoBaixaContasPagar()
    Dim x_linha As String
    Dim i As Integer
    Dim xTotal As Currency
    
    xTotal = 0
    
    lSQL = ""
    lSQL = lSQL & "SELECT SUM(Valor_Pagamento) AS Total"
    lSQL = lSQL & "  FROM Baixa_Pagar"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data_Pagamento >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Data_Pagamento <= " & preparaData(CDate(msk_data_f.Text))
    Set rsTabela = Conectar.RsConexao(lSQL)
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela!total) Then
            xTotal = rsTabela!total
        End If
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    
    If xTotal > 0 Then
        x_linha = "| PAGAMENTO EFETUADO DE CONTAS À PAGAR...:                                     |"
        i = Len(Format(xTotal, "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(xTotal, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
        x_linha = "+------------------------------------------------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
    End If
End Sub
Private Sub ImpResumoBaixaDuplicaraReceber()
    Dim x_linha As String
    Dim i As Integer
    Dim xTotalDinheiro As Currency
    Dim xTotalChVista As Currency
    Dim xTotalChPrazo As Currency
    Dim xTotal As Currency
    xTotalDinheiro = 0
    xTotalChVista = 0
    xTotalChPrazo = 0
    
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Pago]) AS TotalDinheiro,"
    lSQL = lSQL & "       SUM([Valor Pago Cheque Vista]) AS TotalChVista,"
    lSQL = lSQL & "       SUM([Valor Pago Cheque Prazo]) AS TotalChPrazo"
    lSQL = lSQL & "  FROM Baixa_Duplicata_Receber"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data do Pagamento] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND [Data do Pagamento] <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Periodo <= " & Val(cbo_periodo_f.Text)
    Set rsTabela = Conectar.RsConexao(lSQL)
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela!TotalDinheiro) Then
            xTotalDinheiro = rsTabela!TotalDinheiro
        End If
        If Not IsNull(rsTabela!TotalChVista) Then
            xTotalChVista = rsTabela!TotalChVista
        End If
        If Not IsNull(rsTabela!TotalChPrazo) Then
            xTotalChPrazo = rsTabela!TotalChPrazo
        End If
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    
    
'    With tbl_baixa_duplicata_receber
'        If .RecordCount > 0 Then
'            .Seek ">=", g_empresa, CDate(msk_data_i.Text), 0, 0
'            If Not .NoMatch Then
'                Do Until .EOF
'                    If !Empresa <> g_empresa Or ![Data do Pagamento] > CDate(msk_data_f.Text) Then
'                        Exit Do
'                    End If
'                    If !Periodo >= cbo_periodo_i.Text Or !Periodo <= cbo_periodo_f.Text Then
'                        xTotalDinheiro = xTotalDinheiro + ![Valor Pago]
'                        xTotalChVista = xTotalChVista + ![Valor Pago Cheque Vista]
'                        xTotalChPrazo = xTotalChPrazo + ![Valor Pago Cheque Prazo]
'                    End If
'                    .MoveNext
'                Loop
'            End If
'        End If
'    End With
    If xTotalDinheiro > 0 Then
        x_linha = "| DUPLICATAS RECEBIDAS EM DINHEIRO.......:                                     |"
        i = Len(Format(xTotalDinheiro, "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(xTotalDinheiro, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
    End If
    xTotal = xTotalDinheiro + xTotalChVista
    If xTotalChVista > 0 Then
        x_linha = "| DUPLICATAS RECEBIDAS EM CHEQUE A VISTA.:                                     |"
        i = Len(Format(xTotalChVista, "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(xTotalChVista, "###,###,##0.00")
        i = Len(Format(xTotal, "###,###,##0.00"))
        Mid(x_linha, 64 + 14 - i, i) = Format(xTotal, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
    End If
    xTotal = xTotalDinheiro + xTotalChVista + xTotalChPrazo
    If xTotalChPrazo > 0 Then
        x_linha = "| DUPLICATAS RECEBIDAS EM CHEQUE A PRAZO.:                                     |"
        i = Len(Format(xTotalChPrazo, "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(xTotalChPrazo, "###,###,##0.00")
        i = Len(Format(xTotal, "###,###,##0.00"))
        Mid(x_linha, 64 + 14 - i, i) = Format(xTotal, "###,###,##0.00")
        BioImprime "@Printer.Print " & x_linha
    End If
    If (xTotalChPrazo + xTotalChVista + xTotalChPrazo) > 0 Then
        x_linha = "+------------------------------------------------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
    End If
End Sub
Private Sub ImpDetCombustivel(ByVal x_combustivel As String, ByVal x_litros As Currency, ByVal x_valor As Currency, ByVal xLucro As Currency, ByVal xTipoCombustivel As String)
    Dim x_linha As String
    Dim i As Integer
    Dim xLucroLitro As Currency
    Dim xCustoUltimaEntrada As Currency
    If xTipoCombustivel = "  " Or xTipoCombustivel = "AF" Or xTipoCombustivel = "AA" Or xTipoCombustivel = "DA" Or xTipoCombustivel = "GA" Then
        If x_valor = 0 Then
            Exit Sub
        End If
    End If
    x_linha = "|           |             |             |            |            |            |"
    Mid(x_linha, 3, 10) = x_combustivel
    If CCur(x_litros) > 0 Then
        i = Len(Format(x_litros, "##,###,##0.0"))
        Mid(x_linha, 15 + 12 - i, i) = Format(x_litros, "##,###,##0.0")
    End If
    If CCur(x_valor) > 0 Then
        i = Len(Format(x_valor, "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(x_valor, "#,###,##0.00")
        If x_litros > 0 Then
            If xTipoCombustivel = "TT" Then
                xLucroLitro = Format(xLucro / x_litros, "00000000.0000")
            ElseIf xTipoCombustivel = "AF" Then
                xLucroLitro = 0
            Else
                If msk_data_i.Text = msk_data_f.Text Then
                    xCustoUltimaEntrada = CustoUltimaEntrada(xTipoCombustivel, CDate(msk_data_i.Text))
                    xLucroLitro = Format((x_valor / x_litros), "00000000.0000") - xCustoUltimaEntrada
                Else
                    xLucroLitro = Format(xLucro / x_litros, "00000000.0000")
                End If
            End If
            If xTipoCombustivel <> "AF" Then
                i = Len(Format(xLucroLitro, "###,##0.0000"))
                Mid(x_linha, 42 + 12 - i, i) = Format(xLucroLitro, "###,##0.0000")
                i = Len(Format(xLucro, "#,###,##0.00"))
                Mid(x_linha, 55 + 12 - i, i) = Format(xLucro, "#,###,##0.00")
                
                xLucroLitro = Format(xLucro * 100 / x_valor, "##0.0000")
                
                i = Len(Format(xLucroLitro, "##0.0000"))
                Mid(x_linha, 72 + 8 - i, i) = Format(xLucroLitro, "##0.0000")
            End If
        End If
    End If
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpDetMedicaoCombustivel(ByVal pTipoCombustivel As String, ByVal pSemNF As String)
    Dim xLinha As String
    Dim i As Integer
    Dim xEstoqueInicial As Currency
    Dim xEstoqueFinal As Currency
    Dim xQtdEntrada As Currency
    Dim xQtdSaida As Currency
    Dim xEstEscritural As Currency
    Dim xPerdaSobra As Currency
    
    'BioImprime "@Printer.Print " & "+-----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
    '                                         1         2         3         4         5         6         7         8         9        10        11        12   12
    '                                12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
    'BioImprime "@Printer.Print " & "|COMBUSTÍVEL| EST.INICIAL | ENTRADAS    | QTD. SAIDAS | EST.ESCRIT. | EST.  FINAL | PERDA/SOBRA | PER.SOB. R$ |             |"
    'BioImprime "@Printer.Print " & "+-----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
    xLinha = "|           |             |             |             |             |             |             |             |             |           |"
    
    
    If pSemNF = "S" Then
        xQtdEntrada = EntradaCombustivel.TotalEntradaPeriodoSN(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), pTipoCombustivel)
        If xQtdEntrada > 0 Then
            If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
                Mid(xLinha, 2, 11) = Combustivel.Nome
            End If
            i = Len(Format(xQtdEntrada, "####,###,##0"))
            Mid(xLinha, 28 + 12 - i, i) = Format(xQtdEntrada, "####,###,##0")
            BioImprime "@Printer.Print " & xLinha
        End If
        Exit Sub
    End If
    
    xEstoqueInicial = 0
    xEstoqueFinal = 0
    
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Mid(xLinha, 2, 11) = Combustivel.Nome
    End If
    
    xEstoqueInicial = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(msk_data_i.Text), pTipoCombustivel, 0)
    xEstoqueFinal = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(msk_data_f.Text) + 1, pTipoCombustivel, 0)
    xQtdEntrada = EntradaCombustivel.TotalEntradaPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), pTipoCombustivel, 0)
    
    xQtdSaida = MovimentoBomba.TotalVendaPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), pTipoCombustivel, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    xQtdSaida = xQtdSaida - MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), pTipoCombustivel, "")
    
    i = Len(Format(xEstoqueInicial, "####,###,##0"))
    Mid(xLinha, 14 + 12 - i, i) = Format(xEstoqueInicial, "####,###,##0")
    i = Len(Format(xQtdEntrada, "####,###,##0"))
    Mid(xLinha, 28 + 12 - i, i) = Format(xQtdEntrada, "####,###,##0")
    
    xEstEscritural = xEstoqueInicial + xQtdEntrada - xQtdSaida
    xPerdaSobra = xEstoqueFinal - xEstEscritural
    
    i = Len(Format(xQtdSaida, "#,###,##0.00"))
    Mid(xLinha, 42 + 12 - i, i) = Format(xQtdSaida, "#,###,##0.00")
    
    i = Len(Format(xEstEscritural, "#,###,##0.00"))
    Mid(xLinha, 56 + 12 - i, i) = Format(xEstEscritural, "#,###,##0.00")
    
    i = Len(Format(xEstoqueFinal, "####,###,##0"))
    Mid(xLinha, 70 + 12 - i, i) = Format(xEstoqueFinal, "####,###,##0")
    
    i = Len(Format(xPerdaSobra, "#,###,##0.00"))
    Mid(xLinha, 84 + 12 - i, i) = Format(xPerdaSobra, "#,###,##0.00")
    
    If xEstoqueInicial <> 0 Or xEstoqueFinal <> 0 Or xQtdSaida <> 0 Then
        BioImprime "@Printer.Print " & xLinha
    End If
End Sub
Private Sub ImpDetLubrificante()
    Dim xLinha As String
    Dim xValorUnitario As Currency
    Dim i As Integer
    
    '                  1         2         3         4         5         6         7         8         9        10        11        12   12
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
    xLinha = "|    |                                        |            |      |            |"
    
    
    i = Len(Format(rsMovLubrificante("Codigo do Produto2").Value, "###0"))
    Mid(xLinha, 2 + 4 - i, i) = Format(rsMovLubrificante("Codigo do Produto2").Value, "###0")
    Mid(xLinha, 7, 40) = rsMovLubrificante("Nome").Value
            
    xValorUnitario = Format(rsMovLubrificante("Valor Total").Value / rsMovLubrificante("Quantidade").Value, "0000000000.00")
    i = Len(Format(xValorUnitario, "#,###,###.00"))
    Mid(xLinha, 48 + 12 - i, i) = Format(xValorUnitario, "#,###,###.00")
    
    
    i = Len(Format(rsMovLubrificante("Quantidade").Value, "##,###"))
    Mid(xLinha, 61 + 6 - i, i) = Format(rsMovLubrificante("Quantidade").Value, "##,###")
            
    i = Len(Format(rsMovLubrificante("Valor Total").Value, "#,###,###.00"))
    Mid(xLinha, 68 + 12 - i, i) = Format(rsMovLubrificante("Valor Total").Value, "#,###,###.00")
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDet(x_i As Integer, x_abertura As Currency, x_encerrante As Currency, x_litros As Currency, x_valor As Currency, x_historico As String, x_variavel As String)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|  |          |          |         |              |                            |"
    If x_i > 0 Then
        Mid(x_linha, 2, 2) = Format(x_i, "00")
    End If
    If CCur(x_abertura) Or CCur(x_encerrante) > 0 Then
        i = Len(Format(x_abertura, "####,##0.0"))
        Mid(x_linha, 5 + 10 - i, i) = Format(x_abertura, "####,##0.0")
        i = Len(Format(x_encerrante, "####,##0.0"))
        Mid(x_linha, 16 + 10 - i, i) = Format(x_encerrante, "####,##0.0")
        i = Len(Format(x_litros, "###,##0.0"))
        Mid(x_linha, 27 + 9 - i, i) = Format(x_litros, "###,##0.0")
        i = Len(Format(x_valor, "##,###,##0.00"))
        Mid(x_linha, 37 + 13 - i, i) = Format(x_valor, "##,###,##0.00")
    End If
    Mid(x_linha, 52, 13) = x_historico
    If Mid(x_variavel, 1, 3) = "@N@" Then
        If CCur(Mid(x_variavel, 4, Len(x_variavel) - 3)) > 0 Then
            x_variavel = Mid(x_variavel, 4, Len(x_variavel) - 3)
            i = Len(Format(x_variavel, "##,###,##0.00"))
            Mid(x_linha, 66 + 13 - i, i) = Format(x_variavel, "##,###,##0.00")
        End If
    Else
        Mid(x_linha, 65, 15) = Mid(x_variavel, 4, Len(x_variavel) - 3)
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    Dim i As Integer
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
'    '                  1         2         3         4         5         6         7         8
'    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
'    '                                             123456789012345678901234567890
    x_linha = "| MOVIMENTO DAS BOMBAS                                      cidade, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(x_linha, 24, 8) = "- L.M.C."
    End If
    If fEcfInstalada Then
        Mid(x_linha, 24, 8) = "- E.C.F."
    End If
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_linha, 29, 1) = cbo_periodo_i
    Mid(x_linha, 49, 1) = cbo_periodo_f
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "+--+----------+----------+---------+--------------+----------------------------+"
    BioImprime "@Printer.Print " & "|N.| ABERTURA |ENCERRANTE|LTS.SAIDA|VALOR DA SAIDA| COMPOSIÇÃO DO CAIXA        |"
    BioImprime "@Printer.Print " & "+--+----------+----------+---------+--------------+----------------------------+"
'    BioImprime "@Printer.Print " & "+--+---------------------+---------+--------------+----------------------------+"
'    BioImprime "@Printer.Print " & "|N.| OBSERVACAO          |LTS.SAIDA|VALOR DA SAIDA| COMPOSIÇÃO DO CAIXA        |"
'    BioImprime "@Printer.Print " & "+--+---------------------+---------+--------------+----------------------------+"
End Sub
Private Sub ImpCabLubrificante()
    Dim xLinha As String
    BioImprime "@Printer.Print " & "+--+-+--------+----------+---------+----------+---+--------+------+------------+"

    '                                        1         2         3         4         5         6         7         8         9        10        11        12   12
    '                               12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
    BioImprime "@Printer.Print " & "|COD.|NOME DO PRODUTO                         |VLR.UNITARIO|QUANT.| VLR. TOTAL |"
    BioImprime "@Printer.Print " & "+----+----------------------------------------+------------+------+------------+"
End Sub
Private Sub ImpCodificacaoContabil()
    Dim x_linha As String
    Dim i As Integer
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
'                       1         2         3         4         5         6         7         8         9        10        11        12        13     13
'              12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "|        VENDA A VISTA       | VENDAS C/CHEQUE POS DATADO |                            |                            |                   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| DEBITAR...:   1-9          | DEBITAR...:                |                            |                            |                   |"
    'If l_hist_ch_predatado(3) > 0 Then
    '    Mid(x_linha, 44, 5) = "166-0"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CREDITAR..: 185-6          | CREDITAR..:                |                            |                            |                   |"
    'If l_hist_ch_predatado(3) > 0 Then
    '    Mid(x_linha, 44, 5) = "137-6"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| HISTORICO.:   2-7          | HISTORICO.:                |                            |                            |                   |"
    'If l_hist_ch_predatado(3) > 0 Then
    '    Mid(x_linha, 44, 5) = "  2-7"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| VALOR.....:                | VALOR.....:                |                            |                            |                   |"
    'If l_hist_ch_predatado(3) > 0 Then
    '    i = Len(Format(l_hist_ch_predatado(3), "###,###,##0.00"))
    '    Mid(x_linha, 44 + 14 - i, i) = Format(l_hist_ch_predatado(3), "###,###,##0.00")
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+----------------------------+----------------------------+----------------------------+----------------------------+-------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|    VENDAS C/ CARTAO VISA   | VENDAS C/ CARTAO CREDICARD |    VENDAS C/ CARTAO SOLO   | VENDAS C/ CARTAO HIPERCHEQ |                   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| DEBITAR...:                | DEBITAR...:                | DEBITAR...:                | DEBITAR...:                |                   |"
    'If l_hist_visa(1) > 0 Then
    '    Mid(x_linha, 15, 5) = "164-3"
    'End If
    'If l_hist_dinners(1) > 0 Then
    '    Mid(x_linha, 44, 5) = "  5-1"
    'End If
    'If l_hist_amex(1) > 0 Then
    '    Mid(x_linha, 73, 5) = "163-5"
    'End If
    'If l_hist_hipercheque(1) > 0 Then
    '    Mid(x_linha, 102, 5) = "165-1"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CREDITAR..:                | CREDITAR..:                | CREDITAR..:                | CREDITAR..:                |                   |"
    'If l_hist_visa(1) > 0 Then
    '    Mid(x_linha, 15, 5) = "136-8"
    'End If
    'If l_hist_dinners(1) > 0 Then
    '    Mid(x_linha, 44, 5) = "136-8"
    'End If
    'If l_hist_amex(1) > 0 Then
    '    Mid(x_linha, 73, 5) = "136-8"
    'End If
    'If l_hist_hipercheque(1) > 0 Then
    '    Mid(x_linha, 102, 5) = "136-8"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| HISTORICO.:                | HISTORICO.:                | HISTORICO.:                | HISTORICO.:                |                   |"
    'If l_hist_visa(1) > 0 Then
    '    Mid(x_linha, 15, 5) = "  2-7"
    'End If
    'If l_hist_dinners(1) > 0 Then
    '    Mid(x_linha, 44, 5) = "  2-7"
    'End If
    'If l_hist_amex(1) > 0 Then
    '    Mid(x_linha, 73, 5) = "  2-7"
    'End If
    'If l_hist_hipercheque(1) > 0 Then
    '    Mid(x_linha, 102, 5) = "  2-7"
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| VALOR.....:                | VALOR.....:                | VALOR.....:                | VALOR.....:                |                   |"
    'If l_hist_visa(1) > 0 Then
    '    i = Len(Format(l_hist_visa(1), "###,###,##0.00"))
    '    Mid(x_linha, 15 + 14 - i, i) = Format(l_hist_visa(1), "###,###,##0.00")
    'End If
    'If l_hist_dinners(1) > 0 Then
    '    i = Len(Format(l_hist_dinners(1), "###,###,##0.00"))
    '    Mid(x_linha, 44 + 14 - i, i) = Format(l_hist_dinners(1), "###,###,##0.00")
    'End If
    'If l_hist_amex(1) > 0 Then
    '    i = Len(Format(l_hist_amex(1), "###,###,##0.00"))
    '    Mid(x_linha, 73 + 14 - i, i) = Format(l_hist_amex(1), "###,###,##0.00")
    'End If
    'If l_hist_hipercheque(1) > 0 Then
    '    i = Len(Format(l_hist_hipercheque(1), "###,###,##0.00"))
    '    Mid(x_linha, 102 + 14 - i, i) = Format(l_hist_hipercheque(1), "###,###,##0.00")
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+--- Cerrado Informática. ---+----------------------------+----------------------------+----------------------------+-------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_ilha_i.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_GotFocus()
    SendMessageLong cbo_periodo_i.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
    Else
        msk_data_f = RetiraGString(1)
    End If
    g_string = " "
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", 64, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f < cbo_periodo_i Then
        MsgBox "Periodo final deve ser maior.", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf Not Val(txt_ilha_i) > 0 Then
        MsgBox "A ilha inicial deve ser maior que 0.", 64, "Atenção!"
        txt_ilha_i.SetFocus
    ElseIf Val(txt_ilha_f) < Val(txt_ilha_i) Then
        MsgBox "A ilha final deve ser igual ou maior que " & txt_ilha_i & ".", 64, "Atenção!"
        txt_ilha_f.SetFocus
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
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        txt_ilha_i.Text = 1
        txt_ilha_f.Text = 1
        cbo_periodo_i.SetFocus
    End If
    Screen.MousePointer = 1
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
    
    If g_nome_usuario = "L.M.C." Then
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
    ElseIf fEcfInstalada Then
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
    End If
    
    
    If g_nome_usuario = "L.M.C." Then
        Me.Caption = Me.Caption & " - LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        Me.Caption = Me.Caption & " - ECF"
    End If
    PreencheCboPeriodo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 2
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_i_LostFocus()
    If IsDate(msk_data_i.Text) Then
        msk_data_f.Text = msk_data_i.Text
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
Private Sub txt_ilha_f_GotFocus()
    txt_ilha_f.SelStart = 0
    txt_ilha_f.SelLength = Len(txt_ilha_f)
End Sub
Private Sub txt_ilha_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_ilha_i_GotFocus()
    txt_ilha_i.SelStart = 0
    txt_ilha_i.SelLength = Len(txt_ilha_i)
End Sub
Private Sub txt_ilha_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_ilha_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
