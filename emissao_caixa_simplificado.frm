VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_caixa_simplificado 
   Caption         =   "Emissão do Caixa Simplificado"
   ClientHeight    =   3465
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   7155
   Icon            =   "emissao_caixa_simplificado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   7155
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1260
      Picture         =   "emissao_caixa_simplificado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Visualiza Caixa (Simplificado)."
      Top             =   2520
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3180
      Picture         =   "emissao_caixa_simplificado.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprime Caixa (Simplificado)."
      Top             =   2520
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5100
      Picture         =   "emissao_caixa_simplificado.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2520
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6915
      Begin VB.CheckBox chkImprimeLubrificante 
         Caption         =   "Imprime Venda de Lubrificante Detalhada"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1920
         Width           =   3495
      End
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
         Picture         =   "emissao_caixa_simplificado.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "emissao_caixa_simplificado.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6300
         Picture         =   "emissao_caixa_simplificado.frx":6C74
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
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_caixa_simplificado"
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
Dim lBombaAbertura(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaEncerrante(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaLitros(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaValorTotal(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaTipoPreco(1 To gQUANTIDADE_MAXIMA_BICO) As String
Dim lBombaCombustivel(1 To gQUANTIDADE_MAXIMA_BICO) As String
Dim lBombaLitrosAfericao As Currency
Dim lBombaValorAfericao As Currency
Dim lValeAbastecimentoEmitido As Currency
Dim lCartaFrete As Currency
Dim lSuprimentoCaixa As Currency
Dim lAcrescimoPrecoPersonalizado As Currency
Dim lBombaLitrosA As Currency
Dim lBombaLitrosAA As Currency
Dim lBombaLitrosD As Currency
Dim lBombaLitrosDA As Currency
Dim lBombaLitrosG As Currency
Dim lBombaLitrosGA As Currency
Dim l_litros_l As Currency
Dim l_litros_b As Currency
Dim lBombaTotalLitros As Currency
Dim lBombaValorA As Currency
Dim lBombaValorAA As Currency
Dim lBombaValorD As Currency
Dim lBombaValorDA As Currency
Dim lBombaValorG As Currency
Dim lBombaValorGA As Currency
Dim l_valor_l As Currency
Dim l_valor_b As Currency
Dim lBombaTotalValor As Currency

Dim lQtdComposicao As Integer
Dim lTotalComposicao As Currency
Dim lValorComposicao(0 To 30) As Currency
Dim lNomeComposicao(0 To 30) As String
Dim lUltimoBico As Integer

Dim lDiferencaCaixa As Currency
Dim l_nome_funcionario As String

Dim lSQL As String
Private rsMovComposicaoCaixa As New ADODB.Recordset
Private rsMovLubrificante As New ADODB.Recordset

Private Bomba As New cBomba
Private Combustivel As New cCombustivel
Private Funcionario As New cFuncionario
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoCartaFrete As New cMovimentoCartaFrete
Private MovimentoComposicaoCaixa As New cMovimentoComposicaoCaixa
Private SuprimentoCaixa As New cSuprimentoCaixa
Private rsTotalizador As New ADODB.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsMovLubrificante = Nothing
    Set rsMovComposicaoCaixa = Nothing
    
    Set Bomba = Nothing
    Set Combustivel = Nothing
    Set Funcionario = Nothing
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
    lQtdComposicao = 0
    lTotalComposicao = 0
    For i = 0 To 30
        lValorComposicao(i) = 0
        lNomeComposicao(i) = ""
    Next
    
    For i = 1 To gQUANTIDADE_MAXIMA_BICO
        lBombaAbertura(i) = 0
        lBombaEncerrante(i) = 0
        lBombaLitros(i) = 0
        lBombaValorTotal(i) = 0
        lBombaTipoPreco(i) = ""
        lBombaCombustivel(i) = ""
    Next
    lBombaLitrosAfericao = 0
    lBombaValorAfericao = 0
    lValeAbastecimentoEmitido = 0
    lCartaFrete = 0
    lSuprimentoCaixa = 0
    lAcrescimoPrecoPersonalizado = 0
    lBombaLitrosA = 0
    lBombaLitrosAA = 0
    lBombaLitrosD = 0
    lBombaLitrosDA = 0
    lBombaLitrosG = 0
    lBombaLitrosGA = 0
    l_litros_l = 0
    l_litros_b = 0
    lBombaTotalLitros = 0
    lBombaValorA = 0
    lBombaValorAA = 0
    lBombaValorD = 0
    lBombaValorDA = 0
    lBombaValorG = 0
    lBombaValorGA = 0
    l_valor_l = 0
    l_valor_b = 0
    lBombaTotalValor = 0
    
    lDiferencaCaixa = 0
    
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica se exiete composição de caixa
    If Not MovimentoComposicaoCaixa.ExisteRegistroData(g_empresa, CDate(msk_data_i.Text), Val(txt_ilha_i.Text), Val(cbo_periodo_i.Text), 1) Then
        If (MsgBox("Composição não cadastrada!" & Chr(10) & "Deseja continuar?", vbYesNo + vbDefaultButton2, "Erro de integridade!")) = 7 Then
            msk_data_i.SetFocus
            Exit Sub
        End If
    End If
    TotalizaComposicaoCaixa
    TotalizaLubrificante
    If MovimentoBomba.ExisteMovimentoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_i.Text)) Then
        ImpDados
    End If
    cmd_sair.SetFocus
    If g_nivel_acesso = 4 Or g_usuario = 8 Then
        msk_data_i.Text = Format(CDate(msk_data_i.Text) + 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(msk_data_f.Text) + 1, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub ImpDados()
    Dim i As Integer
    LoopMovimentoBomba
    CalculaAfericao
    CalculaValeAbastecimentoEmitido
    lCartaFrete = MovimentoCartaFrete.TotalCartaFrete(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), 1)
    lSuprimentoCaixa = SuprimentoCaixa.TotalPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    lAcrescimoPrecoPersonalizado = MovCupomFiscalItem.TotalAcrescimo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), True)
    If lBombaTotalLitros > 0 Then
        ImpCab
        For i = 1 To lUltimoBico
            Call ImpDetBomba(i, lBombaAbertura(i), lBombaEncerrante(i), lBombaLitros(i), lBombaValorTotal(i), lBombaTipoPreco(i), lBombaCombustivel(i))
        Next
        If chkImprimeLubrificante.Value = 1 Then
            LoopMovimentoLubrificante
        End If
        ImpResumoCombustiveis
        If UCase(g_nome_empresa) Like "*OLIVEIRA*" Then
            ImpResumoGeral
        End If
        If UCase(g_nome_empresa) Like "*REDENÇÃO*" Then
            ImpResumoRedencao
        End If
        'ImpResumoDeposito
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Caixa (Simplificado)|@|"
        frm_preview.Show 1
    End If
End Sub
Private Function LocalizaCodigoComposicao(ByVal pNome As String) As Integer
    Dim rsComposicaoCaixa As New ADODB.Recordset
    LocalizaCodigoComposicao = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Composicao_Caixa.Codigo"
    lSQL = lSQL & "     FROM Composicao_Caixa"
    lSQL = lSQL & "    WHERE Composicao_Caixa.Nome LIKE " & preparaTexto("%" & pNome & "%")
    'Abre RecordSet
    Set rsComposicaoCaixa = Conectar.RsConexao(lSQL)
    If rsComposicaoCaixa.RecordCount > 0 Then
        LocalizaCodigoComposicao = rsComposicaoCaixa("Codigo").Value
    End If
    rsComposicaoCaixa.Close
    Set rsComposicaoCaixa = Nothing
End Function
Private Function LocalizaValorComposicao(ByVal pCodigo As Integer, ByVal pTipoMovimento As Integer) As Currency
    LocalizaValorComposicao = 0
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT SUM(Valor) AS Total"
    lSQL = lSQL & "     FROM Movimento_Composicao_Caixa"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "      AND Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    If pTipoMovimento > 0 Then
        lSQL = lSQL & "      AND [Tipo do Movimento] = " & pTipoMovimento
    End If
    lSQL = lSQL & "      AND [Codigo da Composicao] = " & pCodigo
    'Abre RecordSet
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQL)
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        If Not IsNull(rsMovComposicaoCaixa("Total").Value) Then
            LocalizaValorComposicao = rsMovComposicaoCaixa("Total").Value
        End If
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
End Function
Private Sub LoopMovimentoBomba()
    Dim xBico As Integer
    'loop movimento das bombas
    lUltimoBico = MovimentoBomba.UltimoBicoComMovimento(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    For xBico = 1 To lUltimoBico
        'Le apenas para utilizar o tipo de combustivel do bico
        If MovimentoBomba.LocalizarPrimeiroPeriodoBico(g_empresa, CDate(msk_data_i.Text), xBico, 0) Then
        End If
        lBombaAbertura(xBico) = MovimentoBomba.AberturaBicoDataPeriodo(g_empresa, CDate(msk_data_i.Text), xBico, Val(cbo_periodo_i.Text))
        lBombaEncerrante(xBico) = MovimentoBomba.EncerranteBicoDataPeriodo(g_empresa, CDate(msk_data_i.Text), xBico, Val(cbo_periodo_i.Text))
        lBombaLitros(xBico) = MovimentoBomba.TotalLitrosBicoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
        lBombaValorTotal(xBico) = MovimentoBomba.TotalValorBicoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
        If Bomba.LocalizarCodigo(g_empresa, xBico) Then
            lBombaTipoPreco(xBico) = Bomba.TipoPreco
        End If
        lBombaCombustivel(xBico) = MovimentoBomba.TipoCombustivel
        Select Case Trim(MovimentoBomba.TipoCombustivel)
            Case "A"
                lBombaLitrosA = lBombaLitrosA + lBombaLitros(xBico)
                lBombaValorA = lBombaValorA + lBombaValorTotal(xBico)
            Case "AA"
                lBombaLitrosAA = lBombaLitrosAA + lBombaLitros(xBico)
                lBombaValorAA = lBombaValorAA + lBombaValorTotal(xBico)
            Case "D"
                lBombaLitrosD = lBombaLitrosD + lBombaLitros(xBico)
                lBombaValorD = lBombaValorD + lBombaValorTotal(xBico)
            Case "DA"
                lBombaLitrosDA = lBombaLitrosDA + lBombaLitros(xBico)
                lBombaValorDA = lBombaValorDA + lBombaValorTotal(xBico)
            Case "G"
                lBombaLitrosG = lBombaLitrosG + lBombaLitros(xBico)
                lBombaValorG = lBombaValorG + lBombaValorTotal(xBico)
            Case "GA"
                lBombaLitrosGA = lBombaLitrosGA + lBombaLitros(xBico)
                lBombaValorGA = lBombaValorGA + lBombaValorTotal(xBico)
        End Select
        lBombaTotalLitros = lBombaTotalLitros + lBombaLitros(xBico)
        lBombaTotalValor = lBombaTotalValor + lBombaValorTotal(xBico)
    Next
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
    'Abre RecordSet
    Set rsMovLubrificante = New ADODB.Recordset
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        ImpCabLubrificante
        rsMovLubrificante.MoveFirst
        Do Until rsMovLubrificante.EOF
            ImpDetLubrificante
            rsMovLubrificante.MoveNext
        Loop
    End If
End Sub
Private Sub CalculaAfericao()
    Dim xLitrosAfericao As Currency
    Dim xValorAfericao As Currency
    
    'Alcool
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosA = lBombaLitrosA - xLitrosAfericao
    lBombaValorA = lBombaValorA - xValorAfericao
    'Alcool Aditivado
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosAA = lBombaLitrosAA - xLitrosAfericao
    lBombaValorAA = lBombaValorAA - xValorAfericao
    'Diesel
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosD = lBombaLitrosD - xLitrosAfericao
    lBombaValorD = lBombaValorD - xValorAfericao
    'Diesel Aditivado
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosDA = lBombaLitrosDA - xLitrosAfericao
    lBombaValorDA = lBombaValorDA - xValorAfericao
    'Gasolina
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosG = lBombaLitrosG - xLitrosAfericao
    lBombaValorG = lBombaValorG - xValorAfericao
    'Gasolina Aditivada
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", "")
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", "")
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    lBombaLitrosGA = lBombaLitrosGA - xLitrosAfericao
    lBombaValorGA = lBombaValorGA - xValorAfericao
    
    lBombaTotalLitros = lBombaTotalLitros - lBombaLitrosAfericao
    lBombaTotalValor = lBombaTotalValor - lBombaValorAfericao
End Sub
Private Sub TotalizaComposicaoCaixa()
    Dim i As Integer
    
    If msk_data_i.Text = msk_data_f.Text And cbo_periodo_i.Text = cbo_periodo_f.Text And txt_ilha_i.Text = txt_ilha_f.Text Then
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "   SELECT Movimento_Composicao_Caixa.[Codigo do Funcionario],"
        lSQL = lSQL & "          Funcionario.Nome"
        lSQL = lSQL & "     FROM Movimento_Composicao_Caixa, Funcionario"
        lSQL = lSQL & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data = " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo = " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] = " & Val(txt_ilha_i.Text)
        lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 1
        lSQL = lSQL & "      AND Funcionario.Empresa = " & g_empresa
        lSQL = lSQL & "      AND Funcionario.Codigo = Movimento_Composicao_Caixa.[Codigo do Funcionario]"
        'Abre RecordSet
        Set rsMovComposicaoCaixa = New ADODB.Recordset
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
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    lSQL = lSQL & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQL = lSQL & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New ADODB.Recordset
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
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
End Sub
Private Sub TotalizaLubrificante()
    'loop movimento de lubrificante
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT SUM(Movimento_Lubrificante.[Valor Total]) AS Total"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(txt_ilha_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(txt_ilha_f.Text)
    'Abre RecordSet
    Set rsMovLubrificante = New ADODB.Recordset
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    rsMovLubrificante.MoveFirst
    If Not IsNull(rsMovLubrificante("Total").Value) Then
        l_valor_l = rsMovLubrificante("Total").Value
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
    Set rsTotalizador = New ADODB.Recordset
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
    Set rsTotalizador = New ADODB.Recordset
    Set rsTotalizador = Conectar.RsConexao(xSQL)
    If Not rsTotalizador.EOF Then
        If Not IsNull(rsTotalizador("Total").Value) Then
            lValeAbastecimentoEmitido = lValeAbastecimentoEmitido + rsTotalizador("Total").Value
        End If
    End If
    rsTotalizador.Close
    Set rsTotalizador = Nothing
End Sub
Private Sub ImpResumoGeral()
    Dim xLinha As String
    
    BioImprime "@Printer.Print " & ""
    BioImprime "@Printer.Print " & ""
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| DESCRIÇAO               |    VALOR    | DESCRIÇAO              |    VALOR    |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Combustiveis Dinheiro   |             | Depósitos Bancários    |             |"
    BioImprime "@Printer.Print " & "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Lub.Adit.Outros Dinheiro|             |                        |             |"
    BioImprime "@Printer.Print " & "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Combustíveis Cheques    |             | Boletos                |             |"
    BioImprime "@Printer.Print " & "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| *** SUB-TOTAL           |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Cheques Pré p/ Dia      |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Recebimento de Notas    |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "|                         |             | Fundo Caixa            |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Cheques Troco           |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| Fundo de Caixa          |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & "| *** TOTAIS              |             |                        |             |"
    BioImprime "@Printer.Print " & "+-------------------------+-------------+------------------------+-------------+"




End Sub
Private Sub ImpResumoCombustiveis()
    Dim xLinha As String
    Dim i As Integer
    Dim xQtdLinha As Integer
    Dim xOrdemComposicao As Integer
    
    xQtdLinha = 0
    xOrdemComposicao = -1
    If lBombaLitrosA > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosAA > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosD > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosDA > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosG > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosGA > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If l_valor_l > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    If lBombaLitrosAfericao > 0 Then
        xQtdLinha = xQtdLinha + 1
    End If
    
    If chkImprimeLubrificante.Value = 0 Then
        BioImprime "@Printer.Print " & "+--+--------+----+--------+----+--------++----------+--------------+--------+--+"
    Else
        BioImprime "@Printer.Print " & "+----+------+-------------+-------------+-----+------------+------+------------+"
    End If
    BioImprime "@Printer.Print " & "|COMBUSTÍVEL|    LITROS   |    VALOR    |          COMPOSIÇÃO DE CAIXA         |"
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+--------------------------------------+"
    
    
    
    
    If lBombaLitrosA > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("ÁLCOOL    ", lBombaLitrosA, lBombaValorA, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lBombaLitrosAA > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("ÁLCOOL +  ", lBombaLitrosAA, lBombaValorAA, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lBombaLitrosD > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("DIESEL    ", lBombaLitrosD, lBombaValorD, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lBombaLitrosDA > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("DIESEL +  ", lBombaLitrosDA, lBombaValorDA, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lBombaLitrosG > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("GASOLINA  ", lBombaLitrosG, lBombaValorG, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lBombaLitrosGA > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("GASOLINA +", lBombaLitrosGA, lBombaValorGA, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    xOrdemComposicao = xOrdemComposicao + 1
    Call ImpDetCombustivel("----------", 0, 0, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    xOrdemComposicao = xOrdemComposicao + 1
    Call ImpDetCombustivel("SUB-TOTAL ", lBombaTotalLitros, lBombaTotalValor, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    xOrdemComposicao = xOrdemComposicao + 1
    Call ImpDetCombustivel("----------", 0, 0, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    xOrdemComposicao = xOrdemComposicao + 1
    Call ImpDetCombustivel("ÓLEOS/LUBR", l_litros_l, l_valor_l, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    If lBombaLitrosAfericao > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("AFERIÇÕES", lBombaLitrosAfericao, lBombaValorAfericao, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lValeAbastecimentoEmitido > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("V.ABAST.EMI", 0, lValeAbastecimentoEmitido, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lAcrescimoPrecoPersonalizado > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("ACRES.PREÇO", 0, lAcrescimoPrecoPersonalizado, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    ElseIf lAcrescimoPrecoPersonalizado < 0 Then
        lAcrescimoPrecoPersonalizado = 0
    End If
    If lCartaFrete > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("CARTA FRETE", 0, lCartaFrete, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    If lSuprimentoCaixa > 0 Then
        xOrdemComposicao = xOrdemComposicao + 1
        Call ImpDetCombustivel("SUPR. CAIXA", 0, lSuprimentoCaixa, lNomeComposicao(xOrdemComposicao), lValorComposicao(xOrdemComposicao))
    End If
    
    
    For i = (xOrdemComposicao + 1) To 20
        If lValorComposicao(i) > 0 Then
            Call ImpDetCombustivel("", 0, 0, lNomeComposicao(i), lValorComposicao(i))
        Else
            Exit For
        End If
    Next
    
    Call ImpDetCombustivel("", 0, 0, "+--------------------------------------+", 0)
    Call ImpDetCombustivel("", 0, 0, "Responsavel: " & l_nome_funcionario, 0)
    Call ImpDetCombustivel("", 0, 0, "Total Geral:", lTotalComposicao)
    Call ImpDetCombustivel("----------", 0, 0, "+--------------------------------------+", 0)
    
    lBombaTotalLitros = lBombaTotalLitros + l_litros_l + l_litros_b + lBombaLitrosAfericao
    'Soma Bomba + Lubrificante + Borracharia + Aferição
    lBombaTotalValor = lBombaTotalValor + l_valor_l + l_valor_b + lBombaValorAfericao + lValeAbastecimentoEmitido + lSuprimentoCaixa + lAcrescimoPrecoPersonalizado
    lDiferencaCaixa = lTotalComposicao - lBombaTotalValor
'        lDiferencaCaixa= lDiferencaCaixa + l_valor_l - lValeAbastecimentoEmitido - lAcrescimoPrecoPersonalizado - lCartaFrete
    Call ImpDetCombustivel("TOT. GERAL", lBombaTotalLitros, lBombaTotalValor, "DIFERENÇA DE CAIXA...:", lDiferencaCaixa)
    Call ImpDetCombustivel("----------", 0, 0, "+--------------------------------------+", 0)
    
    
    If lQtdComposicao > xQtdLinha Then
    End If
    
End Sub
Private Sub ImpResumoDeposito()
    Dim xLinha As String
    Dim i As Integer
    Dim xValor As Currency
    
    xValor = 0
    BioImprime "@@Printer.FontName = Draft 10cpi"
    xLinha = "  Cheques A Vista + Dinheiro.:                                    "
    xValor = xValor + lValorComposicao(0) + lValorComposicao(1)
    i = Len(Format(xValor, "###,###,##0.00"))
    Mid(xLinha, 32 + 14 - i, i) = Format(xValor, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
End Sub
Private Sub ImpResumoRedencao()
    Dim xLinha As String
    Dim i As Integer
    Dim xValorDinheiroComb As Currency
    Dim xValorDinheiroLubri As Currency
    Dim xValorCheque As Currency
    Dim xValor As Currency
    Dim xCodigo As Integer

    BioImprime "@@Printer.FontName = Draft 10cpi"
    xLinha = ""
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & xLinha
    
    
    xCodigo = LocalizaCodigoComposicao("DINHEIRO")
    If xCodigo > 0 Then
        xValorDinheiroComb = LocalizaValorComposicao(xCodigo, 1)
        xValorDinheiroLubri = LocalizaValorComposicao(xCodigo, 2)
    End If
    xValor = xValorDinheiroComb
    xValor = xValor + xValorDinheiroLubri
    xCodigo = LocalizaCodigoComposicao("vista")
    If xCodigo > 0 Then
        xValorCheque = LocalizaValorComposicao(xCodigo, 1)
    End If
    xValor = xValor + xValorCheque
    
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DESCRIÇAO               |    VALOR    | DESCRIÇAO              |    VALOR    |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Combustiveis Dinheiro   |             | Depósitos Bancários    |             |"
    i = Len(Format(xValorDinheiroComb, "##,###,##0.00"))
    Mid(xLinha, 28 + 13 - i, i) = Format(xValorDinheiroComb, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Lub.Adit.Outros Dinheiro|             |                        |             |"
    i = Len(Format(xValorDinheiroLubri, "##,###,##0.00"))
    Mid(xLinha, 28 + 13 - i, i) = Format(xValorDinheiroLubri, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Combustíveis Cheques    |             | Boletos                |             |"
    i = Len(Format(xValorCheque, "##,###,##0.00"))
    Mid(xLinha, 28 + 13 - i, i) = Format(xValorCheque, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                         |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| *** SUB-TOTAL           |             |                        |             |"
    i = Len(Format(xValor, "##,###,##0.00"))
    Mid(xLinha, 28 + 13 - i, i) = Format(xValor, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Cheques Pré p/ Dia      |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Recebimento de Notas    |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                         |             | Fundo Caixa            |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Cheques Troco           |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Fundo de Caixa          |             |                        |             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| *** TOTAIS              |             |                        |             |"
    i = Len(Format(xValor, "##,###,##0.00"))
    Mid(xLinha, 28 + 13 - i, i) = Format(xValor, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------+-------------+------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    
    
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    xLinha = ""
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDetBomba(ByVal pBico As Integer, ByVal pAbertura As Currency, ByVal pEncerrante As Currency, ByVal pLitros As Currency, ByVal pValorTotal As Currency, ByVal pTipoPreco As String, ByVal pCombustivel As String)
    Dim xLinha As String
    Dim xValorUnitario As String
    Dim i As Integer
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|  |             |             |         |          |              |        |  |"
    Mid(xLinha, 2, 2) = Format(pBico, "00")
    i = Len(Format(pAbertura, "#,###,##0.00"))
    Mid(xLinha, 6 + 12 - i, i) = Format(pAbertura, "#,###,##0.00")
    i = Len(Format(pEncerrante, "#,###,##0.00"))
    Mid(xLinha, 20 + 12 - i, i) = Format(pEncerrante, "#,###,##0.00")
    i = Len(Format(pLitros, "##,##0.00"))
    Mid(xLinha, 33 + 9 - i, i) = Format(pLitros, "##,##0.00")
    If pLitros = 0 Then
        xValorUnitario = 0
        If MovimentoBomba.LocalizarCodigo(g_empresa, CDate(msk_data_f.Text), Val(cbo_periodo_f.Text), pBico, 999) Then
            xValorUnitario = MovimentoBomba.PrecoVenda
        End If
    Else
        xValorUnitario = Format(pValorTotal / pLitros, "00000.0000")
    End If
    i = Len(Format(xValorUnitario, "###0.0000"))
    Mid(xLinha, 43 + 9 - i, i) = Format(xValorUnitario, "###0.0000")
    i = Len(Format(pValorTotal, "##,###,##0.00"))
    Mid(xLinha, 54 + 13 - i, i) = Format(pValorTotal, "##,###,##0.00")
    
    If pTipoPreco = "V" Then
        Mid(xLinha, 70, 5) = "Vista"
    ElseIf pTipoPreco = "P" Then
        Mid(xLinha, 70, 5) = "Prazo"
    End If
    Mid(xLinha, 78, 2) = pCombustivel
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDetCombustivel(ByVal pCombustivel As String, ByVal pLitros As Currency, ByVal pTotal As Currency, ByVal pComposicao As String, ByVal pValorComposicao As Currency)
    Dim xLinha As String
    Dim i As Integer
    Dim xLucroLitro As Currency
    Dim xCustoUltimaEntrada As Currency
    'If xTipoCombustivel = "  " Or xTipoCombustivel = "AF" Or xTipoCombustivel = "AA" Or xTipoCombustivel = "DA" Or xTipoCombustivel = "GA" Then
    '    If x_valor = 0 Then
    '        Exit Sub
    '    End If
    'End If
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "| ÁLCOOL    |        202,3|      315,59 | Dinheiro                      619,00 |"
    xLinha = "|           |             |             |                                      |"
    
    
    
    If pCombustivel = "----------" Then
        Mid(xLinha, 1, 41) = "+-----------+-------------+-------------+"
    Else
        Mid(xLinha, 3, 10) = pCombustivel
        If CCur(pLitros) > 0 Then
            i = Len(Format(pLitros, "#,###,##0.00"))
            Mid(xLinha, 15 + 12 - i, i) = Format(pLitros, "#,###,##0.00")
        End If
        If pTotal > 0 Then
            i = Len(Format(pTotal, "#,###,##0.00"))
            Mid(xLinha, 28 + 12 - i, i) = Format(pTotal, "#,###,##0.00")
        End If
    End If
    If pValorComposicao > 0 Then
        i = Len(Format(pValorComposicao, "#,###,##0.00"))
        Mid(xLinha, 67 + 12 - i, i) = Format(pValorComposicao, "#,###,##0.00")
    End If
    
    If pValorComposicao < 0 Then
        i = Len(Format(pValorComposicao, "#,###,##0.00;####,##0.00-"))
        Mid(xLinha, 67 + 12 - i, i) = Format(pValorComposicao, "#,###,##0.00;####,##0.00-")
    End If
    
    If pComposicao = "+--------------------------------------+" Then
        Mid(xLinha, 41, 40) = pComposicao
    ElseIf pComposicao = "DIFERENÇA DE CAIXA...:" Then
        If pValorComposicao = 0 Then
            Mid(xLinha, 43, 20) = "CAIXA OK            "
        ElseIf pValorComposicao < 0 Then
            Mid(xLinha, 43, 20) = "FALTOU NO CAIXA     "
        ElseIf pValorComposicao > 0 Then
            Mid(xLinha, 43, 20) = "PASSOU NO CAIXA     "
        End If
    ElseIf Mid(pComposicao, 1, 12) = "Responsavel:" Then
        Mid(xLinha, 43, 37) = pComposicao
    Else
        Mid(xLinha, 43, 20) = pComposicao
    End If
    BioImprime "@Printer.Print " & xLinha
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
    x_linha = "| RELATÓRIO DO CAIXA (SIMPLIFICADO)                         cidade, __/__/____ |"
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
    BioImprime "@Printer.Print " & "+--+-------------+-------------+---------+----------+--------------+--------+--+"
    BioImprime "@Printer.Print " & "|N.|   ABERTURA  |  ENCERRANTE |LTS.SAIDA|VLR LITRO |VALOR DA SAIDA| PREÇO  |CB|"
    BioImprime "@Printer.Print " & "+--+-------------+-------------+---------+----------+--------------+--------+--+"
End Sub
Private Sub ImpCabLubrificante()
    Dim xLinha As String
    BioImprime "@Printer.Print " & "+--+-+-----------+-------------+---------+----+-----+------+------++--------+--+"
    BioImprime "@Printer.Print " & "|COD.|NOME DO PRODUTO                         |VLR.UNITARIO|QUANT.| VLR. TOTAL |"
    BioImprime "@Printer.Print " & "+----+----------------------------------------+------------+------+------------+"
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
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
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
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
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f < cbo_periodo_i Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf Not Val(txt_ilha_i) > 0 Then
        MsgBox "A ilha inicial deve ser maior que 0.", vbInformation, "Atenção!"
        txt_ilha_i.SetFocus
    ElseIf Val(txt_ilha_f) < Val(txt_ilha_i) Then
        MsgBox "A ilha final deve ser igual ou maior que " & txt_ilha_i & ".", vbInformation, "Atenção!"
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
        'MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
        Me.Caption = Me.Caption & " - LMC"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        'MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
        Me.Caption = Me.Caption & " - ECF"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
    Else
        'MovimentoBomba.NomeTabela = "Movimento_Bomba"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
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
