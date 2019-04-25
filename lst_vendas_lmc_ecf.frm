VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_vendas_lmc_ecf 
   Caption         =   "Vendas do L.M.C. x Documentos Fiscais"
   ClientHeight    =   5490
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "lst_vendas_lmc_ecf.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_vendas_lmc_ecf.frx":030A
   ScaleHeight     =   5490
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "lst_vendas_lmc_ecf.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza resumo do L.M.C."
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "lst_vendas_lmc_ecf.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime resumo do L.M.C."
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "lst_vendas_lmc_ecf.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4500
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.CheckBox chkCalculaNFCe 
         Caption         =   "Calcula NFCe"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkSomaCupomFiscal 
         Caption         =   "Soma Cupom Fiscal"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkCalculaNFe 
         Caption         =   "Calcula NFe (Saída e Devolução)"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkSomenteResumo 
         Caption         =   "Somente Resumo"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CheckBox chkImprimirDescontoUnitario 
         Caption         =   "Imprimir Desconto Unitário"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CheckBox chkSemDesconto 
         Caption         =   "Desconsidera Desconto"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkGeraDesconto 
         Caption         =   "Gera desconto automático"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_vendas_lmc_ecf.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_vendas_lmc_ecf.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_vendas_lmc_ecf.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_vendas_lmc_ecf"
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
'Dim l_sql As String
'Dim l_tipo_combustivel As String
'Dim l_total_recebido As Currency
'Dim l_total_vendas As Currency
'Dim l_total_valor_vendas As Currency
'Dim l_total_afericao As Currency
'Dim l_perdas_sobras As Currency
'Dim lCustoMedio As Currency
'Dim lSubCustoMedio As Currency
'Dim lTotalCusto As Currency
'Dim lSQl As String
'Dim lNomeTabelaAfericao As String
Dim lQtdVendaBombaA As Currency
Dim lQtdVendaBombaAA As Currency
Dim lQtdVendaBombaD As Currency
Dim lQtdVendaBombaDA As Currency
Dim lQtdVendaBombaG As Currency
Dim lQtdVendaBombaGA As Currency
Dim lQtdVendaBombaGE As Currency
Dim lValorVendaBombaA As Currency
Dim lValorVendaBombaAA As Currency
Dim lValorVendaBombaD As Currency
Dim lValorVendaBombaDA As Currency
Dim lValorVendaBombaG As Currency
Dim lValorVendaBombaGA As Currency
Dim lValorVendaBombaGE As Currency
Dim lQtdVendaECFA As Currency
Dim lQtdVendaECFAA As Currency
Dim lQtdVendaECFD As Currency
Dim lQtdVendaECFDA As Currency
Dim lQtdVendaECFG As Currency
Dim lQtdVendaECFGA As Currency
Dim lQtdVendaECFGE As Currency
Dim lValorVendaECFA As Currency
Dim lValorVendaECFAA As Currency
Dim lValorVendaECFD As Currency
Dim lValorVendaECFDA As Currency
Dim lValorVendaECFG As Currency
Dim lValorVendaECFGA As Currency
Dim lValorVendaECFGE As Currency
Dim lQtdDiferencaA As Currency
Dim lQtdDiferencaAA As Currency
Dim lQtdDiferencaD As Currency
Dim lQtdDiferencaDA As Currency
Dim lQtdDiferencaG As Currency
Dim lQtdDiferencaGA As Currency
Dim lQtdDiferencaGE As Currency
Dim lValorDiferencaA As Currency
Dim lValorDiferencaAA As Currency
Dim lValorDiferencaD As Currency
Dim lValorDiferencaDA As Currency
Dim lValorDiferencaG As Currency
Dim lValorDiferencaGA As Currency
Dim lValorDiferencaGE As Currency

Dim lSubQtdVendaBomba As Currency
Dim lSubValorVendaBomba As Currency
Dim lSubQtdVendaECF As Currency
Dim lSubValorVendaECF As Currency
Dim lSubQtdDiferenca As Currency
Dim lSubValorDiferenca As Currency
Dim lTotalDescontoA As Currency
Dim lTotalDescontoAA As Currency
Dim lTotalDescontoD As Currency
Dim lTotalDescontoDA As Currency
Dim lTotalDescontoG As Currency
Dim lTotalDescontoGA As Currency
Dim lTotalDescontoGE As Currency

Dim lSQL As String
Dim rstAfericao As New adodb.Recordset
Dim rstCombustivel As New adodb.Recordset
Dim rstMedicaoCombustivel As New adodb.Recordset
Dim rstNotaFiscalSaidaItem As New adodb.Recordset
Dim rstNotaFiscalSaidaItemDevolucao As New adodb.Recordset
Dim rstVendaCombustivel As New adodb.Recordset
Dim rstVendaCupomFiscal As New adodb.Recordset
Dim rstMovimentoDocumentoEletronicoItem As New adodb.Recordset
Dim lNomeTabelaMedicaoCombustivel As String
Dim lNomeTabelaMovimentoAfericao As String
Dim lNomeTabelaMovimentoBomba As String

Dim MedicaoCombustivel As New cMedicaoCombustivel
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub BuscaDadosAfericao()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade) AS TotalQuantidade,"
    lSQL = lSQL & " SUM([Valor Total]) AS TotalVenda"
    lSQL = lSQL & "  FROM " & lNomeTabelaMovimentoAfericao
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstAfericao = New adodb.Recordset
    Set rstAfericao = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosCombustivel()
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & " FROM Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Codigo, Nome"
    Set rstCombustivel = New adodb.Recordset
    Set rstCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosMedicaoCombustivel()
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    xDataInicial = CDate(msk_data_i.Text) + 1
    xDataFinal = CDate(msk_data_f.Text) + 1
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM([Desconto Dia Anterior]) AS TotalDesconto"
    lSQL = lSQL & "  FROM " & lNomeTabelaMedicaoCombustivel
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(xDataInicial)
    lSQL = lSQL & "   AND Data <= " & preparaData(xDataFinal)
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstMedicaoCombustivel = New adodb.Recordset
    Set rstMedicaoCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosNotaFiscalSaidaItem()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade) AS TotalQuantidade,"
    lSQL = lSQL & " SUM(Total) AS TotalVenda"
    lSQL = lSQL & "  FROM MovimentoNotaFiscalSaidaItem"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND Cancelado = " & preparaBooleano(False)
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("5929")
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("6929")
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("5664")
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("6664")
    'Não soma cfop 1??? e 2???
    lSQL = lSQL & "   AND CFOP NOT LIKE " & preparaTexto("1%")
    lSQL = lSQL & "   AND CFOP NOT LIKE " & preparaTexto("2%")
    ' Não Trazer Devolução 5661, 6661
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("5661")
    lSQL = lSQL & "   AND CFOP <> " & preparaTexto("6661")
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstNotaFiscalSaidaItem = New adodb.Recordset
    Set rstNotaFiscalSaidaItem = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosNotaFiscalSaidaItem_Devolucao()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade) AS TotalQuantidade,"
    lSQL = lSQL & " SUM(Total) AS TotalVenda"
    lSQL = lSQL & "  FROM MovimentoNotaFiscalSaidaItem"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND Cancelado = " & preparaBooleano(False)
    lSQL = lSQL & "   AND (CFOP = " & preparaTexto("1662")
    lSQL = lSQL & "        OR CFOP = " & preparaTexto("2662")
    lSQL = lSQL & "        OR CFOP = " & preparaTexto("5661")
    lSQL = lSQL & "        OR CFOP = " & preparaTexto("6661")
    lSQL = lSQL & "        OR CFOP = " & preparaTexto("1949")
    lSQL = lSQL & "        )"
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstNotaFiscalSaidaItemDevolucao = New adodb.Recordset
    Set rstNotaFiscalSaidaItemDevolucao = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosVendaCombustivel()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM([Quantidade da Saida]) AS TotalQuantidade,"
    lSQL = lSQL & " ROUND(SUM([Quantidade da Saida] * [Preco de Venda] + [Total Acrescimo] - [Total Desconto]) ,2) AS TotalVenda"
    lSQL = lSQL & "  FROM " & lNomeTabelaMovimentoBomba
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstVendaCombustivel = New adodb.Recordset
    Set rstVendaCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosVendaCupomFiscal()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", Data"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade) AS TotalQuantidade,"
    lSQL = lSQL & " SUM([Valor Total]) AS TotalVenda,"
    lSQL = lSQL & " SUM([Valor do Desconto]) AS TotalDesconto"
    lSQL = lSQL & "  FROM Movimento_Cupom_Fiscal"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Tipo de Combustivel] <> '  '"
    lSQL = lSQL & "   AND [Codigo da Ecf] < 900"
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstVendaCupomFiscal = New adodb.Recordset
    Set rstVendaCupomFiscal = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosVendaNFCe()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "TipoCombustivel_MovDEItem"
    If chkSomenteResumo.Value = 0 Then
        xTextoAgrupar = xTextoAgrupar + ", DataEmissao_MovDEItem"
    End If
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade_MovDEItem) AS TotalQuantidade,"
    lSQL = lSQL & " SUM(ValorTotalLiquido_MovDEItem) AS TotalVendaLiquida,"
    lSQL = lSQL & " SUM(ValorDesconto_MovDEItem) AS TotalDesconto"
    lSQL = lSQL & "  FROM MovimentoDocumentoEletronicoItem"
    lSQL = lSQL & " WHERE IdEstabelecimento_MovDEItem = " & g_empresa
    lSQL = lSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(False)
    lSQL = lSQL & "   AND Saida_MovDEItem = " & preparaBooleano(True)
    lSQL = lSQL & "   AND Modelo_MovDEItem = " & preparaTexto("65")
    lSQL = lSQL & "   AND DataEmissao_MovDEItem >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND DataEmissao_MovDEItem <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND Cancelado_MovDEItem = " & preparaBooleano(False)
    lSQL = lSQL & "   AND TipoCombustivel_MovDEItem <> '  '"
    lSQL = lSQL & "   AND EtapaConcluida_MovDEItem = 9"
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstMovimentoDocumentoEletronicoItem = New adodb.Recordset
    Set rstMovimentoDocumentoEletronicoItem = Conectar.RsConexao(lSQL)
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MedicaoCombustivel = Nothing
    
    Set rstAfericao = Nothing
    Set rstCombustivel = Nothing
    Set rstMedicaoCombustivel = Nothing
    Set rstNotaFiscalSaidaItem = Nothing
    Set rstNotaFiscalSaidaItemDevolucao = Nothing
    Set rstVendaCombustivel = Nothing
    Set rstVendaCupomFiscal = Nothing
    Set rstMovimentoDocumentoEletronicoItem = Nothing
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim xString As String
    
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                                                                           Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| VENDAS DO L.M.C. x DOCUMENTOS FISCAIS                                                                             Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 3, 40) = "VENDAS DO L.M.C. x DOCUMENTOS FISCAIS  "
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        Mid(xLinha, 3, 40) = "VENDAS DA PISTA x DOCUMENTOS FISCAIS   "
    Else
        Mid(xLinha, 3, 40) = "VENDAS MOV.BOMBA x DOCUMENTOS FISCAIS  "
    End If
    If chkSemDesconto.Value = 1 Then
        Mid(xLinha, 60, 25) = "DESCONTO DESCONSIDERADO  "
    End If
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____  SOMATORIA FISCAL:                                                                              |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    xString = ""
    If chkSomaCupomFiscal.Value = vbChecked Then
        xString = "(CUPOM FISCAL)"
    End If
    If chkCalculaNFe.Value = vbChecked Then
        If xString <> "" Then
            xString = xString & " +/- "
        End If
        xString = xString & "(NF-e)"
    End If
    If chkCalculaNFCe.Value = vbChecked Then
        If xString <> "" Then
            xString = xString & " + "
        End If
        xString = xString & "(NFC-e)"
    End If
    Mid(xLinha, 60, 40) = xString
    BioImprime "@Printer.Print " & xLinha
    'xLinha = "| PRODUTO.....:                                                                                                                         |"
    'Mid(xLinha, 17, 30) = Mid(cbo_combustivel, 6, Len(cbo_combustivel))
    'BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  DATA  DO  | COMBUSTIVEL         |VENDA  BOMBAS|VENDA  BOMBAS|  VENDA ECF  |  VENDA ECF  |  DIFERENÇA  |  DIFERENÇA  |PRECO   VENDA|  |"
    If chkImprimirDescontoUnitario.Value = 1 Then
        Mid(xLinha, 121, 13) = "DESCONTO  ECF"
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  MOVIMENTO |                     |  EM LITROS  | EM  VALORES |  EM LITROS  | EM  VALORES |  EM LITROS  | EM VALORES  |   UNITARIO  |  |"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDet(ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String, ByVal pNomeCombustivel As String)
    Dim xLinha As String
    Dim i As Integer
    Dim xQtdVendaBomba As Currency
    Dim xValorVendaBomba As Currency
    Dim xQtdVendaECF As Currency
    Dim xValorVendaECF As Currency
    Dim xQtdDiferenca As Currency
    Dim xValorDiferenca As Currency
    Dim xValorDescontoECF As Currency
    Dim xCodigoCombustivel(0 To 5) As Integer
    Dim xValor As Currency
    Dim Fim As Boolean
    Dim xValorDescontoLMC As Currency
    Dim xValorDescontoUnitario As Currency
    
    Fim = False
    
    'xQtdVendaBomba = MovimentoBomba.QuantidadeVendaData(g_empresa, pDataInicial, pDataFinal, pTipoCombustivel, 0)
    'xValorVendaBomba = MovimentoBomba.ValorVendaPeriodo(g_empresa, pDataInicial, pDataFinal, pTipoCombustivel, 1, 9)
    If BuscaRegistroVendaCombustivel(pTipoCombustivel, pDataInicial) Then
        xQtdVendaBomba = rstVendaCombustivel!TotalQuantidade
        xValorVendaBomba = rstVendaCombustivel!TotalVenda
    End If
    'xQtdVendaBomba = xQtdVendaBomba - MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, pDataInicial, pDataFinal, 1, 9, pTipoCombustivel, "")
    'xValorVendaBomba = xValorVendaBomba - MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, pDataInicial, pDataFinal, 1, 9, pTipoCombustivel, "")
    If BuscaRegistroAfericao(pTipoCombustivel, pDataInicial) Then
        xQtdVendaBomba = xQtdVendaBomba - rstAfericao!TotalQuantidade
        xValorVendaBomba = xValorVendaBomba - rstAfericao!TotalVenda
    End If
    xValorDescontoLMC = 0
    xValorDescontoUnitario = 0
    If chkImprimirDescontoUnitario.Value = 1 Then
        If xQtdVendaBomba > 0 Then
            xValorDescontoUnitario = xValorVendaBomba / xQtdVendaBomba
        End If
    End If
    
    If chkSemDesconto.Value = 0 Then
        'xValorDescontoLMC = MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, pDataInicial + 1, pDataFinal + 1, pTipoCombustivel)
        If BuscaRegistroMedicaoCombustivel(pTipoCombustivel, pDataInicial + 1) Then
            xValorDescontoLMC = rstMedicaoCombustivel!TotalDesconto
        End If
        xValorVendaBomba = xValorVendaBomba - xValorDescontoLMC
    End If
            
    xQtdVendaECF = 0
    xValorVendaECF = 0
    xValorDescontoECF = 0
    If chkSomaCupomFiscal.Value = vbChecked Then
        'xQtdVendaECF = MovimentoCupomFiscal.QuantidadeCombustivelVendaData(g_empresa, pTipoCombustivel, pDataInicial, pDataFinal, 1, 9)
        'xValorVendaECF = MovimentoCupomFiscal.ValorCombustivelVendaData(g_empresa, pTipoCombustivel, pDataInicial, pDataFinal, 1, 9)
        'xValorDescontoECF = MovimentoCupomFiscal.DescontoCombustivelVendaData(g_empresa, pTipoCombustivel, pDataInicial, pDataFinal, 1, 9)
        If BuscaRegistroVendaCupomFiscal(pTipoCombustivel, pDataInicial) Then
            xQtdVendaECF = xQtdVendaECF + rstVendaCupomFiscal!TotalQuantidade
            xValorVendaECF = xValorVendaECF + rstVendaCupomFiscal!TotalVenda
            xValorDescontoECF = xValorDescontoECF + rstVendaCupomFiscal!TotalDesconto
        End If
    End If
    If chkCalculaNFCe.Value = vbChecked Then
        If BuscaRegistroNFCe(pTipoCombustivel, pDataInicial) Then
            xQtdVendaECF = xQtdVendaECF + rstMovimentoDocumentoEletronicoItem!TotalQuantidade
            xValorVendaECF = xValorVendaECF + rstMovimentoDocumentoEletronicoItem!TotalVendaLiquida
            'xValorDescontoECF = xValorDescontoECF + rstVendaCupomFiscal!TotalDesconto
        End If
    End If
    If chkCalculaNFe.Value = vbChecked Then
        'xQtdVendaECF = xQtdVendaECF + MovNotaFiscalSaidaItem.QuantidadeCombustivelVendaData(g_empresa, pTipoCombustivel, pDataInicial, pDataFinal, False, False)
        'xValorVendaECF = xValorVendaECF + MovNotaFiscalSaidaItem.ValorCombustivelVendaData(g_empresa, pTipoCombustivel, pDataInicial, pDataFinal, False, False)
        If BuscaRegistroNotaFiscalSaidaItem(pTipoCombustivel, pDataInicial) Then
            xQtdVendaECF = xQtdVendaECF + rstNotaFiscalSaidaItem!TotalQuantidade
            xValorVendaECF = xValorVendaECF + rstNotaFiscalSaidaItem!TotalVenda
        End If
        If BuscaRegistroNotaFiscalSaidaItem_Devolucao(pTipoCombustivel, pDataInicial) Then
            xQtdVendaECF = xQtdVendaECF - rstNotaFiscalSaidaItemDevolucao!TotalQuantidade
            xValorVendaECF = xValorVendaECF - rstNotaFiscalSaidaItemDevolucao!TotalVenda
        End If
    End If
    
    xValorVendaECF = xValorVendaECF - xValorDescontoECF
    xQtdDiferenca = xQtdVendaBomba - xQtdVendaECF
    xValorDiferenca = xValorVendaBomba - xValorVendaECF
    If chkImprimirDescontoUnitario.Value = 1 Then
        If xQtdVendaECF > 0 Then
            xValorDescontoUnitario = xValorDescontoUnitario - (xValorVendaECF / xQtdVendaECF)
        End If
    End If

    'Sub-Total
    lSubQtdVendaBomba = lSubQtdVendaBomba + xQtdVendaBomba
    lSubValorVendaBomba = lSubValorVendaBomba + xValorVendaBomba
    lSubQtdVendaECF = lSubQtdVendaECF + xQtdVendaECF
    lSubValorVendaECF = lSubValorVendaECF + xValorVendaECF
    lSubQtdDiferenca = lSubQtdDiferenca + xQtdDiferenca
    lSubValorDiferenca = lSubValorDiferenca + xValorDiferenca


    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
    Mid(xLinha, 3, 10) = Format(pDataInicial, "dd/mm/yyyy")
    Mid(xLinha, 16, 20) = pNomeCombustivel
    
    i = Len(Format(xQtdVendaBomba, "####,##0.00"))
    Mid(xLinha, 38 + 11 - i, i) = Format(xQtdVendaBomba, "####,##0.00")
    i = Len(Format(xValorVendaBomba, "####,##0.00"))
    Mid(xLinha, 52 + 11 - i, i) = Format(xValorVendaBomba, "####,##0.00")
    i = Len(Format(xQtdVendaECF, "####,##0.00"))
    Mid(xLinha, 66 + 11 - i, i) = Format(xQtdVendaECF, "####,##0.00")
    i = Len(Format(xValorVendaECF, "####,##0.00"))
    Mid(xLinha, 80 + 11 - i, i) = Format(xValorVendaECF, "####,##0.00")
    i = Len(Format(xQtdDiferenca, "####,##0.00"))
    Mid(xLinha, 94 + 11 - i, i) = Format(xQtdDiferenca, "####,##0.00")
    i = Len(Format(xValorDiferenca, "####,##0.00"))
    Mid(xLinha, 108 + 11 - i, i) = Format(xValorDiferenca, "####,##0.00")
        
    If chkGeraDesconto.Value = 1 And g_nome_usuario = "L.M.C." And Me.chkSomenteResumo.Value = 0 Then
        If xQtdDiferenca > -5 And xQtdDiferenca < 5 Then
            If xValorDiferenca > 0 Then
                If BuscaRegistroMedicaoCombustivel(pTipoCombustivel, pDataInicial + 1) Then
                    If rstMedicaoCombustivel!TotalDesconto <> xValorDiferenca Then
                        If MedicaoCombustivel.LocalizarDataCombustivel(g_empresa, CDate(pDataInicial + 1), pTipoCombustivel) Then
                            MedicaoCombustivel.DescontoDiaAnterior = xValorDiferenca
                            If Not MedicaoCombustivel.Alterar(g_empresa, CDate(pDataInicial + 1), MedicaoCombustivel.NumeroTanque) Then
                                MsgBox "Não foi possível alterar o desconto.", vbInformation, "Erro de Integridade!"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If pTipoCombustivel = "A " Then
        lQtdVendaBombaA = lQtdVendaBombaA + xQtdVendaBomba
        lValorVendaBombaA = lValorVendaBombaA + xValorVendaBomba
        lQtdVendaECFA = lQtdVendaECFA + xQtdVendaECF
        lValorVendaECFA = lValorVendaECFA + xValorVendaECF
        lQtdDiferencaA = lQtdDiferencaA + xQtdDiferenca
        lValorDiferencaA = lValorDiferencaA + xValorDiferenca
        lTotalDescontoA = lTotalDescontoA + xValorDescontoLMC
    ElseIf pTipoCombustivel = "AA" Then
        lQtdVendaBombaAA = lQtdVendaBombaAA + xQtdVendaBomba
        lValorVendaBombaAA = lValorVendaBombaAA + xValorVendaBomba
        lQtdVendaECFAA = lQtdVendaECFAA + xQtdVendaECF
        lValorVendaECFAA = lValorVendaECFAA + xValorVendaECF
        lQtdDiferencaAA = lQtdDiferencaAA + xQtdDiferenca
        lValorDiferencaAA = lValorDiferencaAA + xValorDiferenca
        lTotalDescontoAA = lTotalDescontoAA + xValorDescontoLMC
    ElseIf pTipoCombustivel = "D " Then
        lQtdVendaBombaD = lQtdVendaBombaD + xQtdVendaBomba
        lValorVendaBombaD = lValorVendaBombaD + xValorVendaBomba
        lQtdVendaECFD = lQtdVendaECFD + xQtdVendaECF
        lValorVendaECFD = lValorVendaECFD + xValorVendaECF
        lQtdDiferencaD = lQtdDiferencaD + xQtdDiferenca
        lValorDiferencaD = lValorDiferencaD + xValorDiferenca
        lTotalDescontoD = lTotalDescontoD + xValorDescontoLMC
    ElseIf pTipoCombustivel = "DA" Then
        lQtdVendaBombaDA = lQtdVendaBombaDA + xQtdVendaBomba
        lValorVendaBombaDA = lValorVendaBombaDA + xValorVendaBomba
        lQtdVendaECFDA = lQtdVendaECFDA + xQtdVendaECF
        lValorVendaECFDA = lValorVendaECFDA + xValorVendaECF
        lQtdDiferencaDA = lQtdDiferencaDA + xQtdDiferenca
        lValorDiferencaDA = lValorDiferencaDA + xValorDiferenca
        lTotalDescontoDA = lTotalDescontoDA + xValorDescontoLMC
    ElseIf pTipoCombustivel = "G " Then
        lQtdVendaBombaG = lQtdVendaBombaG + xQtdVendaBomba
        lValorVendaBombaG = lValorVendaBombaG + xValorVendaBomba
        lQtdVendaECFG = lQtdVendaECFG + xQtdVendaECF
        lValorVendaECFG = lValorVendaECFG + xValorVendaECF
        lQtdDiferencaG = lQtdDiferencaG + xQtdDiferenca
        lValorDiferencaG = lValorDiferencaG + xValorDiferenca
        lTotalDescontoG = lTotalDescontoG + xValorDescontoLMC
    ElseIf pTipoCombustivel = "GA" Then
        lQtdVendaBombaGA = lQtdVendaBombaGA + xQtdVendaBomba
        lValorVendaBombaGA = lValorVendaBombaGA + xValorVendaBomba
        lQtdVendaECFGA = lQtdVendaECFGA + xQtdVendaECF
        lValorVendaECFGA = lValorVendaECFGA + xValorVendaECF
        lQtdDiferencaGA = lQtdDiferencaGA + xQtdDiferenca
        lValorDiferencaGA = lValorDiferencaGA + xValorDiferenca
        lTotalDescontoGA = lTotalDescontoGA + xValorDescontoLMC
    ElseIf pTipoCombustivel = "GE" Then
        lQtdVendaBombaGE = lQtdVendaBombaGE + xQtdVendaBomba
        lValorVendaBombaGE = lValorVendaBombaGE + xValorVendaBomba
        lQtdVendaECFGE = lQtdVendaECFGE + xQtdVendaECF
        lValorVendaECFGE = lValorVendaECFGE + xValorVendaECF
        lQtdDiferencaGE = lQtdDiferencaGE + xQtdDiferenca
        lValorDiferencaGE = lValorDiferencaGE + xValorDiferenca
        lTotalDescontoGE = lTotalDescontoGE + xValorDescontoLMC
    End If
    
    If xQtdVendaBomba > 0 Then
        xValor = Format(xValorVendaBomba / xQtdVendaBomba, "0000000.0000")
    Else
        xValor = 0
    End If
    i = Len(Format(xValor, "##,##0.0000"))
    Mid(xLinha, 122 + 11 - i, i) = Format(xValor, "##,##0.0000")
    
    If chkImprimirDescontoUnitario.Value = 1 Then
        Mid(xLinha, 122, 11) = "           "
        If xQtdVendaECF > 0 And xQtdVendaBomba > 0 Then
            i = Len(Format(xValorDescontoUnitario, "##,##0.0000"))
            Mid(xLinha, 122 + 11 - i, i) = Format(xValorDescontoUnitario, "##,##0.0000")
        End If
    End If
    If g_lmc = 2 And xQtdDiferenca > 0 And xValorVendaECF > 0 And xValorDiferenca <> 0 Then
        xValor = xQtdDiferenca * (Format(xValorVendaECF / xQtdVendaECF, "000000.0000"))
        If xValor <> xValorDiferenca Then
            xValor = xValorDiferenca - xValor
            xLinha = xLinha & " " & Format(xValor, "####,##0.00")
        End If
    End If
    
    'i = Len(Format(xValorDescontoECF, "####,##0.00"))
    'Mid(xLinha, 122 + 11 - i, i) = Format(xValorDescontoECF, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "|         ** |                     |             |             |             |             |             |             |             |  |"
    Mid(xLinha, 16, 20) = "  SUB-TOTAL"
    i = Len(Format(lSubQtdVendaBomba, "####,##0.00"))
    Mid(xLinha, 38 + 11 - i, i) = Format(lSubQtdVendaBomba, "####,##0.00")
    i = Len(Format(lSubValorVendaBomba, "####,##0.00"))
    Mid(xLinha, 52 + 11 - i, i) = Format(lSubValorVendaBomba, "####,##0.00")
    i = Len(Format(lSubQtdVendaECF, "####,##0.00"))
    Mid(xLinha, 66 + 11 - i, i) = Format(lSubQtdVendaECF, "####,##0.00")
    i = Len(Format(lSubValorVendaECF, "####,##0.00"))
    Mid(xLinha, 80 + 11 - i, i) = Format(lSubValorVendaECF, "####,##0.00")
    i = Len(Format(lSubQtdDiferenca, "####,##0.00"))
    Mid(xLinha, 94 + 11 - i, i) = Format(lSubQtdDiferenca, "####,##0.00")
    i = Len(Format(lSubValorDiferenca, "####,##0.00"))
    Mid(xLinha, 108 + 11 - i, i) = Format(lSubValorDiferenca, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    If lLinha >= 62 Then
        xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    Else
        xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    If lQtdVendaBombaA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("A ")
        i = Len(Format(lQtdVendaBombaA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaA, "####,##0.00")
        i = Len(Format(lValorVendaBombaA, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaA, "####,##0.00")
        i = Len(Format(lQtdVendaECFA, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFA, "####,##0.00")
        i = Len(Format(lValorVendaECFA, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFA, "####,##0.00")
        i = Len(Format(lQtdDiferencaA, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaA, "####,##0.00")
        i = Len(Format(lValorDiferencaA, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaAA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("AA")
        i = Len(Format(lQtdVendaBombaAA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaAA, "####,##0.00")
        i = Len(Format(lValorVendaBombaAA, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaAA, "####,##0.00")
        i = Len(Format(lQtdVendaECFAA, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFAA, "####,##0.00")
        i = Len(Format(lValorVendaECFAA, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFAA, "####,##0.00")
        i = Len(Format(lQtdDiferencaAA, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaAA, "####,##0.00")
        i = Len(Format(lValorDiferencaAA, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaAA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaD > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("D ")
        i = Len(Format(lQtdVendaBombaD, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaD, "####,##0.00")
        i = Len(Format(lValorVendaBombaD, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaD, "####,##0.00")
        i = Len(Format(lQtdVendaECFD, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFD, "####,##0.00")
        i = Len(Format(lValorVendaECFD, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFD, "####,##0.00")
        i = Len(Format(lQtdDiferencaD, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaD, "####,##0.00")
        i = Len(Format(lValorDiferencaD, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaD, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaDA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("DA")
        i = Len(Format(lQtdVendaBombaDA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaDA, "####,##0.00")
        i = Len(Format(lValorVendaBombaDA, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaDA, "####,##0.00")
        i = Len(Format(lQtdVendaECFDA, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFDA, "####,##0.00")
        i = Len(Format(lValorVendaECFDA, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFDA, "####,##0.00")
        i = Len(Format(lQtdDiferencaDA, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaDA, "####,##0.00")
        i = Len(Format(lValorDiferencaDA, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaDA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaG > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("G ")
        i = Len(Format(lQtdVendaBombaG, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaG, "####,##0.00")
        i = Len(Format(lValorVendaBombaG, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaG, "####,##0.00")
        i = Len(Format(lQtdVendaECFG, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFG, "####,##0.00")
        i = Len(Format(lValorVendaECFG, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFG, "####,##0.00")
        i = Len(Format(lQtdDiferencaG, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaG, "####,##0.00")
        i = Len(Format(lValorDiferencaG, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaG, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaGA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("GA")
        i = Len(Format(lQtdVendaBombaGA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaGA, "####,##0.00")
        i = Len(Format(lValorVendaBombaGA, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaGA, "####,##0.00")
        i = Len(Format(lQtdVendaECFGA, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFGA, "####,##0.00")
        i = Len(Format(lValorVendaECFGA, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFGA, "####,##0.00")
        i = Len(Format(lQtdDiferencaGA, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaGA, "####,##0.00")
        i = Len(Format(lValorDiferencaGA, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaGA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaGE > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("GE")
        i = Len(Format(lQtdVendaBombaGE, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaGE, "####,##0.00")
        i = Len(Format(lValorVendaBombaGE, "####,##0.00"))
        Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaGE, "####,##0.00")
        i = Len(Format(lQtdVendaECFGE, "####,##0.00"))
        Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFGE, "####,##0.00")
        i = Len(Format(lValorVendaECFGE, "####,##0.00"))
        Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFGE, "####,##0.00")
        i = Len(Format(lQtdDiferencaGE, "####,##0.00"))
        Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaGE, "####,##0.00")
        i = Len(Format(lValorDiferencaGE, "####,##0.00"))
        Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaGE, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    xLinha = "|         ** |                     |             |             |             |             |             |             |             |  |"
    Mid(xLinha, 16, 20) = "TOTAL GERAL"
    lQtdVendaBombaA = lQtdVendaBombaA + lQtdVendaBombaAA + lQtdVendaBombaD + lQtdVendaBombaDA + lQtdVendaBombaG + lQtdVendaBombaGA + lQtdVendaBombaGE
    i = Len(Format(lQtdVendaBombaA, "####,##0.00"))
    Mid(xLinha, 38 + 11 - i, i) = Format(lQtdVendaBombaA, "####,##0.00")
    lValorVendaBombaA = lValorVendaBombaA + lValorVendaBombaAA + lValorVendaBombaD + lValorVendaBombaDA + lValorVendaBombaG + lValorVendaBombaGA + lValorVendaBombaGE
    i = Len(Format(lValorVendaBombaA, "####,##0.00"))
    Mid(xLinha, 52 + 11 - i, i) = Format(lValorVendaBombaA, "####,##0.00")
    lQtdVendaECFA = lQtdVendaECFA + lQtdVendaECFAA + lQtdVendaECFD + lQtdVendaECFDA + lQtdVendaECFG + lQtdVendaECFGA + lQtdVendaECFGE
    i = Len(Format(lQtdVendaECFA, "####,##0.00"))
    Mid(xLinha, 66 + 11 - i, i) = Format(lQtdVendaECFA, "####,##0.00")
    lValorVendaECFA = lValorVendaECFA + lValorVendaECFAA + lValorVendaECFD + lValorVendaECFDA + lValorVendaECFG + lValorVendaECFGA + lValorVendaECFGE
    i = Len(Format(lValorVendaECFA, "####,##0.00"))
    Mid(xLinha, 80 + 11 - i, i) = Format(lValorVendaECFA, "####,##0.00")
    lQtdDiferencaA = lQtdDiferencaA + lQtdDiferencaAA + lQtdDiferencaD + lQtdDiferencaDA + lQtdDiferencaG + lQtdDiferencaGA + lQtdDiferencaGE
    i = Len(Format(lQtdDiferencaA, "####,##0.00"))
    Mid(xLinha, 94 + 11 - i, i) = Format(lQtdDiferencaA, "####,##0.00")
    lValorDiferencaA = lValorDiferencaA + lValorDiferencaAA + lValorDiferencaD + lValorDiferencaDA + lValorDiferencaG + lValorDiferencaGA + lValorDiferencaGE
    i = Len(Format(lValorDiferencaA, "####,##0.00"))
    Mid(xLinha, 108 + 11 - i, i) = Format(lValorDiferencaA, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpTotalDesconto()
    Dim xLinha As String
    Dim i As Integer
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    If lLinha >= 62 Then
        xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " TOTAL DOS DESCONTOS CONCEDIDOS"
    BioImprime "@Printer.Print " & " "
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  DATA  DO  | COMBUSTIVEL         |VLR. DESCONTO|             |             |             |             |             |             |  |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  MOVIMENTO |                     |  CONCEDIDO  |             |             |             |             |             |             |  |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    BioImprime "@Printer.Print " & xLinha
    
    
    If lTotalDescontoA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("A ")
        i = Len(Format(lTotalDescontoA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoAA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("AA")
        i = Len(Format(lTotalDescontoAA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoAA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoD > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("D ")
        i = Len(Format(lTotalDescontoD, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoD, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoDA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("DA")
        i = Len(Format(lTotalDescontoDA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoDA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoG > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("G ")
        i = Len(Format(lTotalDescontoG, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoG, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoGA > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("GA")
        i = Len(Format(lTotalDescontoGA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoGA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lTotalDescontoGE > 0 Then
        xLinha = "|            |                     |             |             |             |             |             |             |             |  |"
        Mid(xLinha, 16, 20) = BuscaTipoCombustivel("GE")
        i = Len(Format(lTotalDescontoGA, "####,##0.00"))
        Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoGE, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    xLinha = "|         ** |                     |             |             |             |             |             |             |             |  |"
    Mid(xLinha, 16, 20) = "TOTAL GERAL"
    lTotalDescontoA = lTotalDescontoA + lTotalDescontoAA + lTotalDescontoD + lTotalDescontoDA + lTotalDescontoG + lTotalDescontoGA + lTotalDescontoGE
    i = Len(Format(lTotalDescontoA, "####,##0.00"))
    Mid(xLinha, 38 + 11 - i, i) = Format(lTotalDescontoA, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "

End Sub
Private Sub Relatorio()
    Dim xData As Date
    
    ZeraVariaveis
    BuscaDadosAfericao
    BuscaDadosCombustivel
    BuscaDadosMedicaoCombustivel
    BuscaDadosNotaFiscalSaidaItem
    BuscaDadosNotaFiscalSaidaItem_Devolucao
    BuscaDadosVendaCombustivel
    If chkSomaCupomFiscal.Value = vbChecked Then
        BuscaDadosVendaCupomFiscal
    End If
    If chkCalculaNFCe.Value = vbChecked Then
        BuscaDadosVendaNFCe
    End If
    xData = CDate(msk_data_i.Text)
    If chkSomenteResumo.Value = 1 Then
        Call LoopCombustivel(CDate(msk_data_i.Text), CDate(msk_data_f.Text))
    Else
        'Loop data
        Do Until xData > CDate(msk_data_f.Text)
            lSubQtdVendaBomba = 0
            lSubValorVendaBomba = 0
            lSubQtdVendaECF = 0
            lSubValorVendaECF = 0
            lSubQtdDiferenca = 0
            lSubValorDiferenca = 0
            Call LoopCombustivel(xData, xData)
            'Call ImpDet(xData)
            'If chkDetalhadaBico.Value = 1 Then
            '    Call ImpDetBico(xData)
            'End If
            xData = xData + 1
            ImpSubTotal
        Loop
    End If
    ImpTotal
    If chkImprimirDescontoUnitario.Value = 1 Then
        ImpTotalDesconto
    End If
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Vendas do LMC x DOCUMENTOS FISCAIS|@|"
    frm_preview.Show 1
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
    End If
    g_string = ""
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    
    lQtdVendaBombaA = 0
    lQtdVendaBombaAA = 0
    lQtdVendaBombaD = 0
    lQtdVendaBombaDA = 0
    lQtdVendaBombaG = 0
    lQtdVendaBombaGA = 0
    lValorVendaBombaA = 0
    lValorVendaBombaAA = 0
    lValorVendaBombaD = 0
    lValorVendaBombaDA = 0
    lValorVendaBombaG = 0
    lValorVendaBombaGA = 0
    lQtdVendaECFA = 0
    lQtdVendaECFAA = 0
    lQtdVendaECFD = 0
    lQtdVendaECFDA = 0
    lQtdVendaECFG = 0
    lQtdVendaECFGA = 0
    lValorVendaECFA = 0
    lValorVendaECFAA = 0
    lValorVendaECFD = 0
    lValorVendaECFDA = 0
    lValorVendaECFG = 0
    lValorVendaECFGA = 0
    lQtdDiferencaA = 0
    lQtdDiferencaAA = 0
    lQtdDiferencaD = 0
    lQtdDiferencaDA = 0
    lQtdDiferencaG = 0
    lQtdDiferencaGA = 0
    lValorDiferencaA = 0
    lValorDiferencaAA = 0
    lValorDiferencaD = 0
    lValorDiferencaDA = 0
    lValorDiferencaG = 0
    lValorDiferencaGA = 0
    
    lSubQtdVendaBomba = 0
    lSubValorVendaBomba = 0
    lSubQtdVendaECF = 0
    lSubValorVendaECF = 0
    lSubQtdDiferenca = 0
    lSubValorDiferenca = 0
    
    lTotalDescontoA = 0
    lTotalDescontoAA = 0
    lTotalDescontoD = 0
    lTotalDescontoDA = 0
    lTotalDescontoG = 0
    lTotalDescontoGA = 0

End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Dim xData As String
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        If g_nome_usuario = "L.M.C." Then
            msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
            msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
        Else
            xData = Format(Date, "dd/mm/yyyy")
            If Day(CDate(xData)) > 1 Then
                Mid(xData, 1, 2) = Format(Day(CDate(xData)) - 1, "00")
            End If
            msk_data_f.Text = xData
            Mid(xData, 1, 2) = "01"
            msk_data_i.Text = xData
        End If
        cmd_visualizar.SetFocus
    End If
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
    Screen.MousePointer = 1
    CentraForm Me

    If g_nome_usuario = "L.M.C." Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        lNomeTabelaMedicaoCombustivel = "MedicaoCombustivelLMC"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba_LMC"
        lNomeTabelaMovimentoAfericao = "Movimento_Afericao_LMC"
        Me.Caption = Me.Caption & " - LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        lNomeTabelaMedicaoCombustivel = "MedicaoCombustivel"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba_Cupom"
        lNomeTabelaMovimentoAfericao = "Movimento_Afericao"
        Me.Caption = Me.Caption & " - ECF"
    Else
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        lNomeTabelaMedicaoCombustivel = "MedicaoCombustivel"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba"
        lNomeTabelaMovimentoAfericao = "Movimento_Afericao"
    End If
    If g_lmc <> 2 Then
        chkGeraDesconto.Visible = False
        chkSemDesconto.Visible = False
    End If
End Sub
Private Function BuscaTipoCombustivel(ByVal pTipoCombustivel As String) As String
    BuscaTipoCombustivel = "** Inexistente: " & pTipoCombustivel & " **"
    rstCombustivel.MoveFirst
    rstCombustivel.Find "Codigo = " & preparaTexto(pTipoCombustivel)
    If Not rstCombustivel.EOF Then
        BuscaTipoCombustivel = rstCombustivel!Nome
    End If
End Function
Private Function BuscaRegistroAfericao(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroAfericao = False
    rstAfericao.Filter = ""
    If rstAfericao.RecordCount > 0 Then
        rstAfericao.MoveFirst
        rstAfericao.Filter = xCondicao
        If Not rstAfericao.EOF Then
            BuscaRegistroAfericao = True
        End If
    End If
End Function
Private Function BuscaRegistroMedicaoCombustivel(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroMedicaoCombustivel = False
    rstMedicaoCombustivel.Filter = ""
    If rstMedicaoCombustivel.RecordCount > 0 Then
        rstMedicaoCombustivel.MoveFirst
        rstMedicaoCombustivel.Filter = xCondicao
        If Not rstMedicaoCombustivel.EOF Then
            BuscaRegistroMedicaoCombustivel = True
        End If
    End If
End Function
Private Function BuscaRegistroNotaFiscalSaidaItem(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroNotaFiscalSaidaItem = False
    rstNotaFiscalSaidaItem.Filter = ""
    If rstNotaFiscalSaidaItem.RecordCount > 0 Then
        rstNotaFiscalSaidaItem.MoveFirst
        rstNotaFiscalSaidaItem.Filter = xCondicao
        If Not rstNotaFiscalSaidaItem.EOF Then
            BuscaRegistroNotaFiscalSaidaItem = True
        End If
    End If
End Function
Private Function BuscaRegistroNotaFiscalSaidaItem_Devolucao(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroNotaFiscalSaidaItem_Devolucao = False
    rstNotaFiscalSaidaItemDevolucao.Filter = ""
    If rstNotaFiscalSaidaItemDevolucao.RecordCount > 0 Then
        rstNotaFiscalSaidaItemDevolucao.MoveFirst
        rstNotaFiscalSaidaItemDevolucao.Filter = xCondicao
        If Not rstNotaFiscalSaidaItemDevolucao.EOF Then
            BuscaRegistroNotaFiscalSaidaItem_Devolucao = True
        End If
    End If
End Function
Private Function BuscaRegistroVendaCombustivel(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroVendaCombustivel = False
    rstVendaCombustivel.Filter = ""
    If rstVendaCombustivel.RecordCount > 0 Then
        rstVendaCombustivel.MoveFirst
        rstVendaCombustivel.Filter = xCondicao
        If Not rstVendaCombustivel.EOF Then
            BuscaRegistroVendaCombustivel = True
        End If
    End If
End Function
Private Function BuscaRegistroVendaCupomFiscal(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    End If
    
    BuscaRegistroVendaCupomFiscal = False
    rstVendaCupomFiscal.Filter = ""
    If rstVendaCupomFiscal.RecordCount > 0 Then
        rstVendaCupomFiscal.MoveFirst
        rstVendaCupomFiscal.Filter = xCondicao
        If Not rstVendaCupomFiscal.EOF Then
            BuscaRegistroVendaCupomFiscal = True
        End If
    End If
End Function
Private Function BuscaRegistroNFCe(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " TipoCombustivel_MovDEItem = " & preparaTexto(pTipoCombustivel)
    If chkSomenteResumo.Value = 0 Then
        xCondicao = xCondicao & " AND DataEmissao_MovDEItem = " & preparaData(pData)
    End If
    
    BuscaRegistroNFCe = False
    rstMovimentoDocumentoEletronicoItem.Filter = ""
    If rstMovimentoDocumentoEletronicoItem.RecordCount > 0 Then
        rstMovimentoDocumentoEletronicoItem.MoveFirst
        rstMovimentoDocumentoEletronicoItem.Filter = xCondicao
        If Not rstMovimentoDocumentoEletronicoItem.EOF Then
            BuscaRegistroNFCe = True
        End If
    End If
End Function
Private Sub LoopCombustivel(ByVal pDataInicial As Date, ByVal pDataFinal As Date)
    'Dim Fim As Boolean
    Dim xLinha As String
    
    'Fim = False
    rstCombustivel.MoveFirst
    'Combustivel.LocalizarPrimeiro (g_empresa)
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 62 Then
        xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "+------------+---------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
'    Do Until Fim = True
'        Call ImpDet(pDataInicial, pDataFinal)
'        'If chkDetalhadaBico.Value = 1 Then
'        '    Call ImpDetBico(x_data)
'        'End If
'        'Imprimiu = True
'        If Combustivel.LocalizarProximo = False Then
'            Fim = True
'        End If
'    Loop
    Do Until rstCombustivel.EOF
        Call ImpDet(pDataInicial, pDataFinal, rstCombustivel("Codigo").Value, rstCombustivel("Nome").Value)
        rstCombustivel.MoveNext
    Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 5
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
