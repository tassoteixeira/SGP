VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_analise_movimentacao_postos2 
   Caption         =   "Emissão de Análise da Movimentação dos Postos - CUSTO"
   ClientHeight    =   3885
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6675
   Icon            =   "lst_analise_movimentacao_postos2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3885
   ScaleWidth      =   6675
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4740
      Picture         =   "lst_analise_movimentacao_postos2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2940
      Picture         =   "lst_analise_movimentacao_postos2.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprime a análise da movimentação dos postos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_analise_movimentacao_postos2.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Visualiza a análise da movimentação dos postos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      Begin VB.CheckBox chkSomenteUmaEmpresa 
         Caption         =   "Somente a Empresa Selecionada"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkVendaProduto 
         Caption         =   "Venda de P&rodutos"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkVendaCombustivel 
         Caption         =   "Venda de &Combustíveis"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_movimentacao_postos2.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_movimentacao_postos2.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5820
         Picture         =   "lst_analise_movimentacao_postos2.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
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
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5220
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4680
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
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3660
         TabIndex        =   12
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3660
         TabIndex        =   7
         Top             =   720
         Width           =   915
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
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_movimentacao_postos2"
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

Dim lApuracaoCombustivel As Currency
Dim lApuracaoComLubrificante As Currency

'Fim de variáveis padrão para relatório
'Const lLIMITE As Integer = 40
Dim l_litro_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_1(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_2(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_3(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_1(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_2(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_venda_3(1 To gQUANTIDADE_MAXIMA_BICO) As Currency

'------------ DADOS ENTRADA DE COMBUSTIVEL ----------

Dim l_litro_Entrada_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_Entrada_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_Entrada_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_Entrada_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_Entrada_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_litro_Entrada_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency

Dim l_custo_Entrada_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_Entrada_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_Entrada_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_Entrada_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_Entrada_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_Entrada_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency

Dim l_custo_medio_a(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_medio_aa(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_medio_d(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_medio_da(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_medio_g(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim l_custo_medio_ga(1 To gQUANTIDADE_MAXIMA_BICO) As Currency


Dim lCodigoEmpresa(1 To 8) As Integer
Dim lNomeEmpresa(1 To 8) As String
Dim lQtdEmpresa As Integer
Dim lSQL As String

Dim rstEmpresa As ADODB.Recordset
Dim rstTabela As ADODB.Recordset

Private Sub CalculaAfericao(i As Integer)
    lSQL = "SELECT [Tipo de Combustivel],"
    lSQL = lSQL & " SUM(Quantidade) AS TotalQtd,"
    lSQL = lSQL & " SUM(" & preparaArredonda("Quantidade * [Preco de Custo]", 2) & ") AS TotalCusto,"
    lSQL = lSQL & " SUM(" & preparaArredonda("Quantidade * [Preco de Venda]", 2) & ") AS TotalVenda"
    lSQL = lSQL & " FROM Movimento_Afericao"
    lSQL = lSQL & " WHERE Empresa = " & i
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & " AND Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & " AND Transferencia = " & preparaBooleano(False)
    lSQL = lSQL & " GROUP BY [Tipo de Combustivel]"
    
    
    
    
    Set rstTabela = Conectar.RsConexao(lSQL)
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                Select Case Trim(![Tipo de Combustivel])
                    Case "A"
                        l_litro_a(i) = l_litro_a(i) - !TotalQtd
                        l_litro_a(gQUANTIDADE_MAXIMA_BICO) = l_litro_a(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_a(i) = l_custo_a(i) - !TotalCusto
                        l_custo_a(gQUANTIDADE_MAXIMA_BICO) = l_custo_a(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_a(i) = l_venda_a(i) - !TotalVenda
                        l_venda_a(gQUANTIDADE_MAXIMA_BICO) = l_venda_a(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                    Case "AA"
                        l_litro_aa(i) = l_litro_aa(i) - !TotalQtd
                        l_litro_aa(gQUANTIDADE_MAXIMA_BICO) = l_litro_aa(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_aa(i) = l_custo_aa(i) - !TotalCusto
                        l_custo_aa(gQUANTIDADE_MAXIMA_BICO) = l_custo_aa(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_aa(i) = l_venda_aa(i) - !TotalVenda
                        l_venda_aa(gQUANTIDADE_MAXIMA_BICO) = l_venda_aa(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                    Case "D"
                        l_litro_d(i) = l_litro_d(i) - !TotalQtd
                        l_litro_d(gQUANTIDADE_MAXIMA_BICO) = l_litro_d(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_d(i) = l_custo_d(i) - !TotalCusto
                        l_custo_d(gQUANTIDADE_MAXIMA_BICO) = l_custo_d(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_d(i) = l_venda_d(i) - !TotalVenda
                        l_venda_d(gQUANTIDADE_MAXIMA_BICO) = l_venda_d(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                    Case "DA"
                        l_litro_da(i) = l_litro_da(i) - !TotalQtd
                        l_litro_da(gQUANTIDADE_MAXIMA_BICO) = l_litro_da(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_da(i) = l_custo_da(i) - !TotalCusto
                        l_custo_da(gQUANTIDADE_MAXIMA_BICO) = l_custo_da(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_da(i) = l_venda_da(i) - !TotalVenda
                        l_venda_da(gQUANTIDADE_MAXIMA_BICO) = l_venda_da(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                    Case "G"
                        l_litro_g(i) = l_litro_g(i) - !TotalQtd
                        l_litro_g(gQUANTIDADE_MAXIMA_BICO) = l_litro_g(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_g(i) = l_custo_g(i) - !TotalCusto
                        l_custo_g(gQUANTIDADE_MAXIMA_BICO) = l_custo_g(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_g(i) = l_venda_g(i) - !TotalVenda
                        l_venda_g(gQUANTIDADE_MAXIMA_BICO) = l_venda_g(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                    Case "GA"
                        l_litro_ga(i) = l_litro_ga(i) - !TotalQtd
                        l_litro_ga(gQUANTIDADE_MAXIMA_BICO) = l_litro_ga(gQUANTIDADE_MAXIMA_BICO) - !TotalQtd
                        l_custo_ga(i) = l_custo_ga(i) - !TotalCusto
                        l_custo_ga(gQUANTIDADE_MAXIMA_BICO) = l_custo_ga(gQUANTIDADE_MAXIMA_BICO) - !TotalCusto
                        l_venda_ga(i) = l_venda_ga(i) - !TotalVenda
                        l_venda_ga(gQUANTIDADE_MAXIMA_BICO) = l_venda_ga(gQUANTIDADE_MAXIMA_BICO) - !TotalVenda
                End Select
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    lQtdEmpresa = 0
    For i = 1 To gQUANTIDADE_MAXIMA_BICO
        l_litro_a(i) = 0
        l_litro_aa(i) = 0
        l_litro_d(i) = 0
        l_litro_da(i) = 0
        l_litro_g(i) = 0
        l_litro_ga(i) = 0
        l_venda_a(i) = 0
        l_venda_aa(i) = 0
        l_venda_d(i) = 0
        l_venda_da(i) = 0
        l_venda_g(i) = 0
        l_venda_ga(i) = 0
        l_custo_a(i) = 0
        l_custo_aa(i) = 0
        l_custo_d(i) = 0
        l_custo_da(i) = 0
        l_custo_g(i) = 0
        l_custo_ga(i) = 0
        l_custo_1(i) = 0
        l_venda_1(i) = 0
        l_custo_2(i) = 0
        l_venda_2(i) = 0
        l_custo_3(i) = 0
        l_venda_3(i) = 0
        
        l_litro_Entrada_a(i) = 0
        l_litro_Entrada_aa(i) = 0
        l_litro_Entrada_d(i) = 0
        l_litro_Entrada_da(i) = 0
        l_litro_Entrada_g(i) = 0
        l_litro_Entrada_ga(i) = 0
        
        l_custo_Entrada_a(i) = 0
        l_custo_Entrada_aa(i) = 0
        l_custo_Entrada_d(i) = 0
        l_custo_Entrada_da(i) = 0
        l_custo_Entrada_g(i) = 0
        l_custo_Entrada_ga(i) = 0

        l_custo_medio_a(i) = 0
        l_custo_medio_aa(i) = 0
        l_custo_medio_d(i) = 0
        l_custo_medio_da(i) = 0
        l_custo_medio_g(i) = 0
        l_custo_medio_ga(i) = 0

    Next
    For i = 1 To 8
        lCodigoEmpresa(i) = 15
        lNomeEmpresa(i) = ""
    Next
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome Abreviado das Empresas", gArquivoIni)
    lSQL = "SELECT Codigo, Nome, Inativo FROM Empresas ORDER BY Codigo"
    Set rstEmpresa = Conectar.RsConexao(lSQL)
    With rstEmpresa
        .MoveFirst
        i = 0
        Do Until .EOF
            If Not !Inativo Then
                i = i + 1
                If i > 8 Then
                    Exit Do
                End If
                lCodigoEmpresa(i) = !Codigo
                lNomeEmpresa(i) = RetiraGString(i)
            End If
            .MoveNext
        Loop
    End With
    g_string = ""
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
Private Sub Relatorio()
    Dim i As Integer
    
    ZeraVariaveis
    If chkVendaCombustivel.Value = 1 Then
        ImpCab
        ImpCab2
        Call LoopEmpresaMovimentoBomba
        Call ImpSubTotal
        
        ImpCab4
        Call LoopEmpresaEntradaCombustivel
        Call ImpSubTotalEntrada
        
        Call ImpApuracaoResultadosCombustivel
        
        ''Call ImpResumoCombustivel
        If chkVendaProduto.Value = 1 Then
            If lQtdEmpresa > 1 Then
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
        End If
    End If
    If chkVendaProduto.Value = 1 And chkVendaCombustivel.Value = 0 Then
        ImpCab
    End If
    If chkVendaProduto.Value = 1 Then
        ImpCab3
        Call LoopEmpresaMovimentoLubrificante
        Call ImpTotalLubrificante
        Call ImpApuracaoResultadosGeral
    End If
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Análise da Movimentação dos Postos|@|"
    frm_preview.Show 1
    
    rstEmpresa.Close
    rstTabela.Close
    Set rstEmpresa = Nothing
    Set rstTabela = Nothing
    
    cmd_sair.SetFocus
End Sub
Private Sub ImpEntradaCombustivel(i As Integer)

    
    lSQL = " select [Tipo de Combustivel],"
    lSQL = lSQL & " sum(Quantidade) AS TotQtdEntrada, "
    lSQL = lSQL & " sum([Valor da Entrada]) AS TotalEntrada,"
    lSQL = lSQL & " (sum([Valor da Entrada]) / sum(Quantidade)) AS CustoMedio"
    lSQL = lSQL & " FROM Entrada_Combustivel"
    lSQL = lSQL & " Where Empresa = " & i
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " GROUP BY [Tipo de Combustivel]"
    lSQL = lSQL & " ORDER BY [Tipo de Combustivel]"

    Set rstTabela = Conectar.RsConexao(lSQL)
    
        With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                If Trim(![Tipo de Combustivel]) = "A" Then
                    l_litro_Entrada_a(i) = l_litro_Entrada_a(i) + !TotQtdEntrada
                    l_litro_Entrada_a(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_a(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    
                    l_custo_Entrada_a(i) = l_custo_Entrada_a(i) + !TotalEntrada
                    l_custo_Entrada_a(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_a(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_a(i) = l_custo_medio_a(i) + !CustoMedio
                    l_custo_medio_a(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_a(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio
                    
                ElseIf Trim(![Tipo de Combustivel]) = "AA" Then
                    l_litro_Entrada_aa(i) = l_litro_Entrada_aa(i) + !TotQtdEntrada
                    l_litro_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    
                    l_custo_Entrada_aa(i) = l_custo_Entrada_aa(i) + !TotalEntrada
                    l_custo_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_aa(i) = l_custo_medio_aa(i) + !CustoMedio
                    l_custo_medio_aa(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_aa(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio
                   
                ElseIf Trim(![Tipo de Combustivel]) = "D" Then
                    l_litro_Entrada_d(i) = l_litro_Entrada_d(i) + !TotQtdEntrada
                    l_litro_Entrada_d(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_d(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    l_custo_Entrada_d(i) = l_custo_Entrada_d(i) + !TotalEntrada
                    l_custo_Entrada_d(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_d(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_d(i) = l_custo_medio_d(i) + !CustoMedio
                    l_custo_medio_d(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_d(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio

                
                ElseIf Trim(![Tipo de Combustivel]) = "DA" Then
                    l_litro_Entrada_da(i) = l_litro_Entrada_da(i) + !TotQtdEntrada
                    l_litro_Entrada_da(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_da(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    l_custo_Entrada_da(i) = l_custo_Entrada_da(i) + !TotalEntrada
                    l_custo_Entrada_da(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_da(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_da(i) = l_custo_medio_da(i) + !CustoMedio
                    l_custo_medio_da(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_da(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio

                ElseIf Trim(![Tipo de Combustivel]) = "G" Then
                    l_litro_Entrada_g(i) = l_litro_Entrada_g(i) + !TotQtdEntrada
                    l_litro_Entrada_g(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_g(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    l_custo_Entrada_g(i) = l_custo_Entrada_g(i) + !TotalEntrada
                    l_custo_Entrada_g(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_g(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_g(i) = l_custo_medio_g(i) + !CustoMedio
                    l_custo_medio_g(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_g(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio

                ElseIf Trim(![Tipo de Combustivel]) = "GA" Then
                    l_litro_Entrada_ga(i) = l_litro_Entrada_ga(i) + !TotQtdEntrada
                    l_litro_Entrada_ga(gQUANTIDADE_MAXIMA_BICO) = l_litro_Entrada_ga(gQUANTIDADE_MAXIMA_BICO) + !TotQtdEntrada
                    l_custo_Entrada_ga(i) = l_custo_Entrada_ga(i) + !TotalEntrada
                    l_custo_Entrada_ga(gQUANTIDADE_MAXIMA_BICO) = l_custo_Entrada_ga(gQUANTIDADE_MAXIMA_BICO) + !TotalEntrada
                    
                    l_custo_medio_ga(i) = l_custo_medio_ga(i) + !CustoMedio
                    l_custo_medio_ga(gQUANTIDADE_MAXIMA_BICO) = l_custo_medio_ga(gQUANTIDADE_MAXIMA_BICO) + !CustoMedio

                Else
                    MsgBox "teste" & ![Tipo de Combustivel]
                End If
                .MoveNext
            Loop
        End If
    End With



End Sub
Private Sub ImpMovimentoBomba(i As Integer)
    
    lSQL = "SELECT [Tipo de Combustivel],"
    lSQL = lSQL & " SUM([Quantidade da Saida]) AS TotalQtd,"
    lSQL = lSQL & " SUM(" & preparaArredonda("[Quantidade da Saida] * [Preco de Custo]", 2) & ") AS TotalCusto,"
    lSQL = lSQL & " SUM(" & preparaArredonda("[Quantidade da Saida] * [Preco de Venda]", 2) & ") AS TotalVenda,"
    
    'TotalAcrescimo e TotalDesconto são referentes a utilização de automação com 2 preços
    
    lSQL = lSQL & " " & preparaArredonda("SUM([Total Desconto])", 2) & " AS TotalDesconto,"
    lSQL = lSQL & " " & preparaArredonda("SUM([Total Acrescimo])", 2) & " AS TotalAcrescimo"

    
    lSQL = lSQL & " FROM Movimento_Bomba"
    lSQL = lSQL & " WHERE Empresa = " & i
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & " AND Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & " GROUP BY [Tipo de Combustivel]"

    

    Set rstTabela = Conectar.RsConexao(lSQL)
    
    
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                If Trim(![Tipo de Combustivel]) = "A" Then
                    l_litro_a(i) = l_litro_a(i) + !TotalQtd
                    l_litro_a(gQUANTIDADE_MAXIMA_BICO) = l_litro_a(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_a(i) = l_custo_a(i) + !TotalCusto
                    l_custo_a(gQUANTIDADE_MAXIMA_BICO) = l_custo_a(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_a(i) = (l_venda_a(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_a(gQUANTIDADE_MAXIMA_BICO) = (l_venda_a(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    
                ElseIf Trim(![Tipo de Combustivel]) = "AA" Then
                    l_litro_aa(i) = l_litro_aa(i) + !TotalQtd
                    l_litro_aa(gQUANTIDADE_MAXIMA_BICO) = l_litro_aa(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_aa(i) = l_custo_aa(i) + !TotalCusto
                    l_custo_aa(gQUANTIDADE_MAXIMA_BICO) = l_custo_aa(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_aa(i) = (l_venda_aa(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_aa(gQUANTIDADE_MAXIMA_BICO) = (l_venda_aa(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    
                   
                ElseIf Trim(![Tipo de Combustivel]) = "D" Then
                    l_litro_d(i) = l_litro_d(i) + !TotalQtd
                    l_litro_d(gQUANTIDADE_MAXIMA_BICO) = l_litro_d(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_d(i) = l_custo_d(i) + !TotalCusto
                    l_custo_d(gQUANTIDADE_MAXIMA_BICO) = l_custo_d(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_d(i) = (l_venda_d(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_d(gQUANTIDADE_MAXIMA_BICO) = (l_venda_d(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    
               
                ElseIf Trim(![Tipo de Combustivel]) = "DA" Then
                    l_litro_da(i) = l_litro_da(i) + !TotalQtd
                    l_litro_da(gQUANTIDADE_MAXIMA_BICO) = l_litro_da(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_da(i) = l_custo_da(i) + !TotalCusto
                    l_custo_da(gQUANTIDADE_MAXIMA_BICO) = l_custo_da(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_da(i) = (l_venda_da(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_da(gQUANTIDADE_MAXIMA_BICO) = (l_venda_da(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                
                ElseIf Trim(![Tipo de Combustivel]) = "G" Then
                    l_litro_g(i) = l_litro_g(i) + !TotalQtd
                    l_litro_g(gQUANTIDADE_MAXIMA_BICO) = l_litro_g(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_g(i) = l_custo_g(i) + !TotalCusto
                    l_custo_g(gQUANTIDADE_MAXIMA_BICO) = l_custo_g(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_g(i) = (l_venda_g(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_g(gQUANTIDADE_MAXIMA_BICO) = (l_venda_g(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto

                ElseIf Trim(![Tipo de Combustivel]) = "GA" Then
                    l_litro_ga(i) = l_litro_ga(i) + !TotalQtd
                    l_litro_ga(gQUANTIDADE_MAXIMA_BICO) = l_litro_ga(gQUANTIDADE_MAXIMA_BICO) + !TotalQtd
                    l_custo_ga(i) = l_custo_ga(i) + !TotalCusto
                    l_custo_ga(gQUANTIDADE_MAXIMA_BICO) = l_custo_ga(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    
                    l_venda_ga(i) = (l_venda_ga(i) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    l_venda_ga(gQUANTIDADE_MAXIMA_BICO) = (l_venda_ga(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda + !TotalAcrescimo) - !TotalDesconto
                    
                Else
                    MsgBox "teste" & ![Tipo de Combustivel]
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpMovimentoLubrificante(i As Integer)
    
    lSQL = "SELECT Produto.[Codigo do Grupo],"
    lSQL = lSQL & " SUM(" & preparaArredonda("Quantidade * [Valor Custo]", 2) & ") AS TotalCusto,"
    lSQL = lSQL & " SUM([Valor Total]) AS TotalVenda"
    lSQL = lSQL & " FROM Movimento_Lubrificante, Produto"
    lSQL = lSQL & " WHERE Empresa = " & i
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & " AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & " AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
    lSQL = lSQL & " GROUP BY [Codigo do Grupo]"
    Set rstTabela = Conectar.RsConexao(lSQL)
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                If ![Codigo do Grupo] = 1 Then
                    l_custo_1(i) = l_custo_1(i) + !TotalCusto
                    l_custo_1(gQUANTIDADE_MAXIMA_BICO) = l_custo_1(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    l_venda_1(i) = l_venda_1(i) + !TotalVenda
                    l_venda_1(gQUANTIDADE_MAXIMA_BICO) = l_venda_1(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda
                ElseIf ![Codigo do Grupo] = 2 Then
                    l_custo_2(i) = l_custo_2(i) + !TotalCusto
                    l_custo_2(gQUANTIDADE_MAXIMA_BICO) = l_custo_2(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    l_venda_2(i) = l_venda_2(i) + !TotalVenda
                    l_venda_2(gQUANTIDADE_MAXIMA_BICO) = l_venda_2(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda
                ElseIf ![Codigo do Grupo] > 2 Then
                    l_custo_3(i) = l_custo_3(i) + !TotalCusto
                    l_custo_3(gQUANTIDADE_MAXIMA_BICO) = l_custo_3(gQUANTIDADE_MAXIMA_BICO) + !TotalCusto
                    l_venda_3(i) = l_venda_3(i) + !TotalVenda
                    l_venda_3(gQUANTIDADE_MAXIMA_BICO) = l_venda_3(gQUANTIDADE_MAXIMA_BICO) + !TotalVenda
                Else
                    MsgBox "teste" & ![Codigo do Grupo]
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresaEntradaCombustivel()
    Dim i As Integer
    
    Dim x_litro_Entrada As Currency 'Linha de totais
    Dim x_custo_Entrada As Currency 'Linha de totais
    Dim x_custo_medio As Currency 'Linha de totais
    
    lQtdEmpresa = 0
    With rstEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If chkSomenteUmaEmpresa.Value = 0 Or (chkSomenteUmaEmpresa.Value = 1 And !Codigo = g_empresa) Then
                    lQtdEmpresa = lQtdEmpresa + 1
                    i = !Codigo
                    
                    ImpEntradaCombustivel i
                    
                    x_litro_Entrada = l_litro_Entrada_a(i) + l_litro_Entrada_aa(i) + l_litro_Entrada_d(i) + l_litro_Entrada_da(i) + l_litro_Entrada_g(i) + l_litro_Entrada_ga(i)
                    x_custo_Entrada = l_custo_Entrada_a(i) + l_custo_Entrada_aa(i) + l_custo_Entrada_d(i) + l_custo_Entrada_da(i) + l_custo_Entrada_g(i) + l_custo_Entrada_ga(i)
                    x_custo_medio = l_custo_medio_a(i) + l_custo_medio_aa(i) + l_custo_medio_d(i) + l_custo_medio_da(i) + l_custo_medio_g(i) + l_custo_medio_ga(i)
                    
                    
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "A ", l_litro_Entrada_a(i), l_custo_Entrada_a(i), l_custo_medio_a(i), l_venda_a(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "AA", l_litro_Entrada_aa(i), l_custo_Entrada_aa(i), l_custo_medio_aa(i), l_venda_aa(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "D ", l_litro_Entrada_d(i), l_custo_Entrada_d(i), l_custo_medio_d(i), l_venda_d(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "DA", l_litro_Entrada_da(i), l_custo_Entrada_da(i), l_custo_medio_da(i), l_venda_da(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "G ", l_litro_Entrada_g(i), l_custo_Entrada_g(i), l_custo_medio_g(i), l_venda_g(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "GA", l_litro_Entrada_ga(i), l_custo_Entrada_ga(i), l_custo_medio_ga(i), l_venda_ga(i))
                    Call ImpDetEntrada(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "**", x_litro_Entrada, x_custo_Entrada, x_custo_medio, 0)
                    If x_litro_Entrada > 0 Then
                        BioImprime "@Printer.Print " & "+--------------------------------+----+-------------------+--------------+--------------+"
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresaMovimentoBomba()
    Dim i As Integer
    Dim x_litro As Currency
    Dim x_venda As Currency
    Dim x_custo As Currency
    
    lQtdEmpresa = 0
    With rstEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If chkSomenteUmaEmpresa.Value = 0 Or (chkSomenteUmaEmpresa.Value = 1 And !Codigo = g_empresa) Then
                    lQtdEmpresa = lQtdEmpresa + 1
                    i = !Codigo
    '                If !Codigo > 11 Then
    '                    i = 11
    '                End If
                    ImpMovimentoBomba i
                    CalculaAfericao i
                    x_litro = l_litro_a(i) + l_litro_aa(i) + l_litro_d(i) + l_litro_da(i) + l_litro_g(i) + l_litro_ga(i)
                    x_venda = l_venda_a(i) + l_venda_aa(i) + l_venda_d(i) + l_venda_da(i) + l_venda_g(i) + l_venda_ga(i)
                    x_custo = l_custo_a(i) + l_custo_aa(i) + l_custo_d(i) + l_custo_da(i) + l_custo_g(i) + l_custo_ga(i)
                    
                    
                    
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "A ", l_litro_a(i), l_venda_a(i), l_custo_a(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "AA", l_litro_aa(i), l_venda_aa(i), l_custo_aa(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "D ", l_litro_d(i), l_venda_d(i), l_custo_d(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "DA", l_litro_da(i), l_venda_da(i), l_custo_da(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "G ", l_litro_g(i), l_venda_g(i), l_custo_g(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "GA", l_litro_ga(i), l_venda_ga(i), l_custo_ga(i))
                    Call ImpDet(Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO), "**", x_litro, x_venda, x_custo)
                    If x_litro > 0 Then
                        BioImprime "@Printer.Print " & "+--------------------------------+----+-----------+--------------+--------------+--------------+--------+--------+----------------------+"
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresaMovimentoLubrificante()
    Dim i As Integer
    Dim x_venda As Currency
    Dim x_custo As Currency
    
    With rstEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                'If !Codigo > 11 Then
                '    Exit Do
                'End If
                If chkSomenteUmaEmpresa.Value = 0 Or (chkSomenteUmaEmpresa.Value = 1 And !Codigo = g_empresa) Then
                    i = !Codigo
                    ImpMovimentoLubrificante i
                    x_venda = l_venda_1(i) + l_venda_2(i) + l_venda_3(i)
                    x_custo = l_custo_1(i) + l_custo_2(i) + l_custo_3(i)
                    Call ImpDetLubrificante("ÓLEO/LUBRIFICANTES ", l_venda_1(i), l_custo_1(i), Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO))
                    Call ImpDetLubrificante("FILTROS            ", l_venda_2(i), l_custo_2(i), Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO))
                    Call ImpDetLubrificante("DIVERSOS           ", l_venda_3(i), l_custo_3(i), Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO))
                    Call ImpDetLubrificante("*** TOTAL          ", x_venda, x_custo, Mid(!Nome, 1, gQUANTIDADE_MAXIMA_BICO))
                    If x_venda > 0 Then
                        BioImprime "@Printer.Print " & "+----------------------+---------------+---------------+---------------+--------+--------+----------------------------------------------+"
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpSubTotal()
    Dim x_litro As Currency
    Dim x_venda As Currency
    Dim x_custo As Currency
    
    x_litro = l_litro_a(gQUANTIDADE_MAXIMA_BICO) + l_litro_aa(gQUANTIDADE_MAXIMA_BICO) + l_litro_d(gQUANTIDADE_MAXIMA_BICO) + l_litro_da(gQUANTIDADE_MAXIMA_BICO) + l_litro_g(gQUANTIDADE_MAXIMA_BICO) + l_litro_ga(gQUANTIDADE_MAXIMA_BICO)
    x_venda = l_venda_a(gQUANTIDADE_MAXIMA_BICO) + l_venda_aa(gQUANTIDADE_MAXIMA_BICO) + l_venda_d(gQUANTIDADE_MAXIMA_BICO) + l_venda_da(gQUANTIDADE_MAXIMA_BICO) + l_venda_g(gQUANTIDADE_MAXIMA_BICO) + l_venda_ga(gQUANTIDADE_MAXIMA_BICO)
    x_custo = l_custo_a(gQUANTIDADE_MAXIMA_BICO) + l_custo_aa(gQUANTIDADE_MAXIMA_BICO) + l_custo_d(gQUANTIDADE_MAXIMA_BICO) + l_custo_da(gQUANTIDADE_MAXIMA_BICO) + l_custo_g(gQUANTIDADE_MAXIMA_BICO) + l_custo_ga(gQUANTIDADE_MAXIMA_BICO)
    
    
    Call ImpDet("*** TOTAL DOS POSTOS", "**", x_litro, x_venda, x_custo)
    BioImprime "@Printer.Print " & "+--------------------------------+----+-----------+--------------+--------------+--------------+--------+--------+----------------------+"
    BioImprime "@Printer.Print " & " "
End Sub

Private Sub ImpApuracaoResultadosCombustivel()
    Dim xTotalVenda As Currency
    Dim xTotalCustoEntrada As Currency
    Dim xLinha As String
    Dim i As Integer
    
    BioImprime "@Printer.Print " & "+---------------------------------------------------------+-----------------------------------------------------+"
    xLinha = "| *** APURAÇÃO BRUTA DOS RESULTADOS (COMBUSTÍVEIS)        |                                                     |"
    
    xTotalVenda = l_venda_a(gQUANTIDADE_MAXIMA_BICO) + l_venda_aa(gQUANTIDADE_MAXIMA_BICO) + l_venda_d(gQUANTIDADE_MAXIMA_BICO) + l_venda_da(gQUANTIDADE_MAXIMA_BICO) + l_venda_g(gQUANTIDADE_MAXIMA_BICO) + l_venda_ga(gQUANTIDADE_MAXIMA_BICO)
    xTotalCustoEntrada = l_custo_Entrada_a(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_d(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_da(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_g(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_ga(gQUANTIDADE_MAXIMA_BICO)
    
    lApuracaoCombustivel = xTotalVenda - xTotalCustoEntrada
    
    i = Len(Format(lApuracaoCombustivel, "##,###,##0.00"))
    Mid(xLinha, 100 + 12 - i, i) = Format(lApuracaoCombustivel, "####,##0.0")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    lLinha = lLinha + 1
    
    BioImprime "@Printer.Print " & "+---------------------------------------------------------+-----------------------------------------------------+"
    BioImprime "@Printer.Print " & " "

End Sub
Private Sub ImpApuracaoResultadosGeral()
    Dim xTotalVenda As Currency
    Dim xTotalCustoEntrada As Currency
    Dim xLinha As String
    Dim i As Integer

    BioImprime "@Printer.Print " & "+---------------------------------------------------------+-----------------------------------------------------+"
    xLinha = "| *** APURAÇÃO BRUTA DOS RESULTADOS (C/LUBRIFICANTES)     |                                                     |"

    xTotalVenda = l_venda_1(gQUANTIDADE_MAXIMA_BICO) + l_venda_2(gQUANTIDADE_MAXIMA_BICO) + l_venda_3(gQUANTIDADE_MAXIMA_BICO)

    lApuracaoComLubrificante = xTotalVenda + lApuracaoCombustivel

    i = Len(Format(lApuracaoComLubrificante, "##,###,##0.00"))
    Mid(xLinha, 100 + 12 - i, i) = Format(lApuracaoComLubrificante, "####,##0.0")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    lLinha = lLinha + 1

    BioImprime "@Printer.Print " & "+---------------------------------------------------------+-----------------------------------------------------+"
    BioImprime "@Printer.Print " & " "

End Sub

Private Sub ImpSubTotalEntrada()
    Dim x_custo_medio As Currency
    Dim x_litro_Entrada As Currency
    Dim x_custo_Entrada As Currency
    
    
    x_litro_Entrada = l_litro_Entrada_a(gQUANTIDADE_MAXIMA_BICO) + l_litro_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) + l_litro_Entrada_d(gQUANTIDADE_MAXIMA_BICO) + l_litro_Entrada_da(gQUANTIDADE_MAXIMA_BICO) + l_litro_Entrada_g(gQUANTIDADE_MAXIMA_BICO) + l_litro_Entrada_ga(gQUANTIDADE_MAXIMA_BICO)
    x_custo_Entrada = l_custo_Entrada_a(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_aa(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_d(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_da(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_g(gQUANTIDADE_MAXIMA_BICO) + l_custo_Entrada_ga(gQUANTIDADE_MAXIMA_BICO)
    x_custo_medio = l_custo_medio_a(gQUANTIDADE_MAXIMA_BICO) + l_custo_medio_aa(gQUANTIDADE_MAXIMA_BICO) + l_custo_medio_d(gQUANTIDADE_MAXIMA_BICO) + l_custo_medio_da(gQUANTIDADE_MAXIMA_BICO) + l_custo_medio_g(gQUANTIDADE_MAXIMA_BICO) + l_custo_medio_ga(gQUANTIDADE_MAXIMA_BICO)

    
    Call ImpDetEntrada("*** TOTAL DOS POSTOS", "**", x_litro_Entrada, x_custo_Entrada, x_custo_medio, 0)
    BioImprime "@Printer.Print " & "+--------------------------------+----+-------------------+--------------+--------------+"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpTotalLubrificante()
    Dim x_venda As Currency
    Dim x_custo As Currency
    
    x_venda = l_venda_1(gQUANTIDADE_MAXIMA_BICO) + l_venda_2(gQUANTIDADE_MAXIMA_BICO) + l_venda_3(gQUANTIDADE_MAXIMA_BICO)
    x_custo = l_custo_1(gQUANTIDADE_MAXIMA_BICO) + l_custo_2(gQUANTIDADE_MAXIMA_BICO) + l_custo_3(gQUANTIDADE_MAXIMA_BICO)
    Call ImpDetLubrificante("*** TOTAL GERAL    ", x_venda, x_custo, "** TODOS OS POSTOS **")
    BioImprime "@Printer.Print " & "+----------------------+---------------+---------------+---------------+--------+--------+----------------------------------------------+"
    Printer.FontName = "Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpResumoCombustivel()
    Dim x_empresa As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Currency
    Dim xLinha As String
    Dim x_litro(1 To 20) As Currency
    Dim x_venda(1 To 20) As Currency
    
    For i = 1 To 20
        x_litro(i) = l_litro_a(i) + l_litro_aa(i) + l_litro_d(i) + l_litro_da(i) + l_litro_g(i) + l_litro_ga(i)
    Next
    x_venda(1) = l_venda_a(1) + l_venda_a(2) + l_venda_a(3) + l_venda_a(4) + l_venda_a(5) + l_venda_a(6) + l_venda_a(9)
    x_venda(2) = l_venda_aa(1) + l_venda_aa(2) + l_venda_aa(3) + l_venda_aa(4) + l_venda_aa(5) + l_venda_aa(6) + l_venda_aa(9)
    x_venda(3) = l_venda_d(1) + l_venda_d(2) + l_venda_d(3) + l_venda_d(4) + l_venda_d(5) + l_venda_d(6) + l_venda_d(9)
    x_venda(4) = l_venda_da(1) + l_venda_da(2) + l_venda_da(3) + l_venda_da(4) + l_venda_da(5) + l_venda_da(6) + l_venda_da(9)
    x_venda(5) = l_venda_g(1) + l_venda_g(2) + l_venda_g(3) + l_venda_g(4) + l_venda_g(5) + l_venda_g(6) + l_venda_g(9)
    x_venda(6) = l_venda_ga(1) + l_venda_ga(2) + l_venda_ga(3) + l_venda_ga(4) + l_venda_ga(5) + l_venda_ga(6) + l_venda_ga(9)
    x_venda(gQUANTIDADE_MAXIMA_BICO) = x_venda(1) + x_venda(2) + x_venda(3) + x_venda(4) + x_venda(5) + x_venda(6)
    BioImprime "@Printer.Print " & "+----+-----------+-----------+-----------+-----------+-----------+-----------+-----------+-------------+---------+----------------------+"
    xLinha = "|PROD|           |           |           |           |           |           |TOT. LITROS| TOTAL EM R$ |% S/TOTAL|                      |"
    '          1         2         3         4         5         6         7         8         9        10        11        12        13     13
    ' 12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '"|PROD|    0 1    |    0 2    |    0 3    |    0 4    |    0 5    |    0 6    |TOT. LITROS| TOTAL EM R$ |% S/TOTAL|                      |"
    
    '"| VENCIMENTO |            |            |            |            |            |            |            |            |     T O T A L    |"
    For x_empresa = 1 To 8
        i2 = Len(Trim(lNomeEmpresa(x_empresa)))
        i3 = (11 * x_empresa + x_empresa - 5) + ((11 - i2) / 2)
        If Mid(Format(i3, "000.0"), 5, 1) <> "0" Then
            i = Val(i3) + 1
        Else
            i = Val(i3)
        End If
        Mid(xLinha, i, i2) = Trim(lNomeEmpresa(x_empresa))
    Next
    BioImprime "@Printer.Print " & xLinha
    
    
    BioImprime "@Printer.Print " & "+----+-----------+-----------+-----------+-----------+-----------+-----------+-----------+-------------+---------+----------------------+"
    Call ImpDetResumo("A ", l_litro_a(1), l_litro_a(2), l_litro_a(3), l_litro_a(4), l_litro_a(5), l_litro_a(6), l_litro_a(gQUANTIDADE_MAXIMA_BICO), x_venda(1), x_venda(gQUANTIDADE_MAXIMA_BICO))
    Call ImpDetResumo("AA", l_litro_aa(1), l_litro_aa(2), l_litro_aa(3), l_litro_aa(4), l_litro_aa(5), l_litro_aa(6), l_litro_aa(gQUANTIDADE_MAXIMA_BICO), x_venda(2), x_venda(gQUANTIDADE_MAXIMA_BICO))
    Call ImpDetResumo("D ", l_litro_d(1), l_litro_d(2), l_litro_d(3), l_litro_d(4), l_litro_d(5), l_litro_d(6), l_litro_d(gQUANTIDADE_MAXIMA_BICO), x_venda(3), x_venda(gQUANTIDADE_MAXIMA_BICO))
    Call ImpDetResumo("DA", l_litro_da(1), l_litro_da(2), l_litro_da(3), l_litro_da(4), l_litro_da(5), l_litro_da(6), l_litro_da(gQUANTIDADE_MAXIMA_BICO), x_venda(4), x_venda(gQUANTIDADE_MAXIMA_BICO))
    Call ImpDetResumo("G ", l_litro_g(1), l_litro_g(2), l_litro_g(3), l_litro_g(4), l_litro_g(5), l_litro_g(6), l_litro_g(gQUANTIDADE_MAXIMA_BICO), x_venda(5), x_venda(gQUANTIDADE_MAXIMA_BICO))
    Call ImpDetResumo("GA", l_litro_ga(1), l_litro_ga(2), l_litro_ga(3), l_litro_ga(4), l_litro_ga(5), l_litro_ga(6), l_litro_ga(gQUANTIDADE_MAXIMA_BICO), x_venda(6), x_venda(gQUANTIDADE_MAXIMA_BICO))
    BioImprime "@Printer.Print " & "+----+-----------+-----------+-----------+-----------+-----------+-----------+-----------+-------------+---------+----------------------+"
    Call ImpDetResumo("**", x_litro(1), x_litro(2), x_litro(3), x_litro(4), x_litro(6), x_litro(9), x_litro(gQUANTIDADE_MAXIMA_BICO), x_venda(gQUANTIDADE_MAXIMA_BICO), x_venda(gQUANTIDADE_MAXIMA_BICO))
    BioImprime "@Printer.Print " & "+----+-----------+-----------+-----------+-----------+-----------+-----------+-----------+-------------+---------+----------------------+"
End Sub
Private Sub ImpDet(x_empresa As String, x_combustivel As String, x_litro As Currency, x_venda As Currency, x_custo As Currency)
    Dim x_lucro As Currency
    Dim x_porc_1 As Currency
    Dim x_porc_2 As Currency
    Dim xLucroMedio As Currency
    Dim xLinha As String
    Dim i As Integer
    
    If x_venda > 0 Then
        x_lucro = x_venda - x_custo
        x_porc_1 = x_lucro * 100 / x_custo
        x_porc_2 = x_lucro * 100 / x_venda
        xLucroMedio = x_lucro / x_litro
        xLinha = "|                                |    |           |              |              |              |        |        |                      |"
        Mid(xLinha, 3, 30) = x_empresa
        Mid(xLinha, 36, 2) = x_combustivel
        i = Len(Format(x_litro, "####,##0.0"))
        Mid(xLinha, 40 + 10 - i, i) = Format(x_litro, "####,##0.0")
        i = Len(Format(x_venda, "##,###,##0.00"))
        Mid(xLinha, 52 + 13 - i, i) = Format(x_venda, "##,###,##0.00")
        i = Len(Format(x_custo, "##,###,##0.00"))
        Mid(xLinha, 67 + 13 - i, i) = Format(x_custo, "##,###,##0.00")
        i = Len(Format(x_lucro, "##,###,##0.00"))
        Mid(xLinha, 82 + 13 - i, i) = Format(x_lucro, "##,###,##0.00")
        i = Len(Format(x_porc_1, "##0.00"))
        Mid(xLinha, 98 + 6 - i, i) = Format(x_porc_1, "##0.00")
        i = Len(Format(x_porc_2, "##0.00"))
        Mid(xLinha, 107 + 6 - i, i) = Format(x_porc_2, "##0.00")
        i = Len(Format(xLucroMedio, "##0.0000"))
        Mid(xLinha, 128 + 8 - i, i) = Format(xLucroMedio, "##0.0000")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If x_combustivel = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetEntrada(ByVal pNomeEmpresa As String, ByVal pTipoCombustivel As String, ByVal pQtdTotalEntrada As Currency, ByVal pValorTotalEntrada As Currency, ByVal pCustoMedio As Currency, ByVal pTotalVenda As Currency)
    Dim x_lucro As Currency
    Dim x_porc_1 As Currency
    Dim x_porc_2 As Currency
    Dim xLucroMedio As Currency
    Dim xLinha As String
    Dim i As Integer
    
    If pValorTotalEntrada > 0 Then
        x_lucro = pTotalVenda - pValorTotalEntrada
        x_porc_1 = x_lucro * 100 / pValorTotalEntrada

        xLinha = "|                                |    |                   |              |              |"
        Mid(xLinha, 3, 30) = pNomeEmpresa
        Mid(xLinha, 36, 2) = pTipoCombustivel
        i = Len(Format(pQtdTotalEntrada, "####,##0.0"))
        Mid(xLinha, 46 + 12 - i, i) = Format(pQtdTotalEntrada, "####,##0.0")
        i = Len(Format(pValorTotalEntrada, "##,###,##0.00"))
        Mid(xLinha, 63 + 10 - i, i) = Format(pValorTotalEntrada, "##,###,##0.00")
        i = Len(Format(pCustoMedio, "##,###,##0.00"))
        Mid(xLinha, 82 + 6 - i, i) = Format(pCustoMedio, "##,###,##0.00")
        
'        i = Len(Format(x_lucro, "##,###,##0.00"))
'        Mid(xLinha, 82 + 13 - i, i) = Format(x_lucro, "##,###,##0.00")
'        i = Len(Format(x_porc_1, "##0.00"))
'        Mid(xLinha, 98 + 6 - i, i) = Format(x_porc_1, "##0.00")
'        i = Len(Format(x_porc_2, "##0.00"))
'        Mid(xLinha, 107 + 6 - i, i) = Format(x_porc_2, "##0.00")
'        i = Len(Format(xLucroMedio, "##0.0000"))
'        Mid(xLinha, 128 + 8 - i, i) = Format(xLucroMedio, "##0.0000")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If pTipoCombustivel = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetLubrificante(x_historico As String, x_venda As Currency, x_custo As Currency, x_empresa As String)
    Dim x_lucro As Currency
    Dim x_porc_1 As Currency
    Dim x_porc_2 As Currency
    Dim xLinha As String
    Dim i As Integer
    
    If x_venda > 0 Then
        x_lucro = x_venda - x_custo
        x_porc_1 = x_lucro * 100 / x_custo
        x_porc_2 = x_lucro * 100 / x_venda
        xLinha = "|                      |               |               |               |        |        |                                              |"
        Mid(xLinha, 3, 20) = x_historico
        i = Len(Format(x_venda, "##,###,##0.00"))
        Mid(xLinha, 26 + 13 - i, i) = Format(x_venda, "##,###,##0.00")
        i = Len(Format(x_custo, "##,###,##0.00"))
        Mid(xLinha, 42 + 13 - i, i) = Format(x_custo, "##,###,##0.00")
        i = Len(Format(x_lucro, "##,###,##0.00"))
        Mid(xLinha, 58 + 13 - i, i) = Format(x_lucro, "##,###,##0.00")
        i = Len(Format(x_porc_1, "##0.00"))
        Mid(xLinha, 74 + 6 - i, i) = Format(x_porc_1, "##0.00")
        i = Len(Format(x_porc_2, "##0.00"))
        Mid(xLinha, 83 + 6 - i, i) = Format(x_porc_2, "##0.00")
        Mid(xLinha, 92, 40) = x_empresa
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If Mid(x_historico, 1, 3) = "***" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetResumo(x_combustivel As String, x_litro_ns As Currency, x_litro_pl As Currency, x_litro_87 As Currency, x_litro_tk As Currency, x_litro_68 As Currency, x_litro_co As Currency, x_litro_tt As Currency, x_total_venda As Currency, x_total_venda2 As Currency)
    Dim x_porc As Currency
    Dim xLinha As String
    Dim i As Integer
    
    If x_total_venda > 0 Then
        x_porc = x_total_venda * 100 / x_total_venda2
        xLinha = "|    |           |           |           |           |           |           |           |             |       % |                      |"
        Mid(xLinha, 3, 2) = x_combustivel
        If x_litro_ns > 0 Then
            i = Len(Format(x_litro_ns, "####,##0.0"))
            Mid(xLinha, 7 + 10 - i, i) = Format(x_litro_ns, "####,##0.0")
        End If
        If x_litro_pl > 0 Then
            i = Len(Format(x_litro_pl, "####,##0.0"))
            Mid(xLinha, 19 + 10 - i, i) = Format(x_litro_pl, "####,##0.0")
        End If
        If x_litro_87 > 0 Then
            i = Len(Format(x_litro_87, "####,##0.0"))
            Mid(xLinha, 31 + 10 - i, i) = Format(x_litro_87, "####,##0.0")
        End If
        If x_litro_tk > 0 Then
            i = Len(Format(x_litro_tk, "####,##0.0"))
            Mid(xLinha, 43 + 10 - i, i) = Format(x_litro_tk, "####,##0.0")
        End If
        If x_litro_68 > 0 Then
            i = Len(Format(x_litro_68, "####,##0.0"))
            Mid(xLinha, 55 + 10 - i, i) = Format(x_litro_68, "####,##0.0")
        End If
        If x_litro_co > 0 Then
            i = Len(Format(x_litro_co, "####,##0.0"))
            Mid(xLinha, 67 + 10 - i, i) = Format(x_litro_co, "####,##0.0")
        End If
        If x_litro_tt > 0 Then
            i = Len(Format(x_litro_tt, "####,##0.0"))
            Mid(xLinha, 79 + 10 - i, i) = Format(x_litro_tt, "####,##0.0")
        End If
        If x_total_venda > 0 Then
            i = Len(Format(x_total_venda, "##,###,##0.00"))
            Mid(xLinha, 91 + 13 - i, i) = Format(x_total_venda, "##,###,##0.00")
            i = Len(Format(x_porc, "##0.00"))
            Mid(xLinha, 106 + 6 - i, i) = Format(x_porc, "##0.00")
        End If
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If x_combustivel = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    xLinha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| ANÁLISE DA MOVIMENTAÇÃO GERAL DOS POSTOS                        , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____       PERÍODO.: _ AO _     |"
    Mid(xLinha, 29, 10) = msk_data_i
    Mid(xLinha, 42, 10) = msk_data_f
    Mid(xLinha, 69, 1) = cbo_periodo_i
    Mid(xLinha, 74, 1) = cbo_periodo_f
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpCab2()
    Dim xLinha As String
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
   
    BioImprime "@Printer.Print " & "+--------------------------------+----+-----------+--------------+--------------+--------------+--------+--------+----------------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                        |PROD|   LITROS  | TOTAL  VENDA | TOTAL  CUSTO | LUCRO BRUTO  |% S/CUST|% S/VEND| LUCRO MEDIO POR LITRO|"
    BioImprime "@Printer.Print " & "+--------------------------------+----+-----------+--------------+--------------+--------------+--------+--------+----------------------+"
End Sub
Private Sub ImpCab3()
    Dim xLinha As String
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+----------------------+---------------+---------------+---------------+--------+--------+----------------------------------------------+"
    BioImprime "@Printer.Print " & "| PRODUTOS             | TOTAL DE VENDA| TOTAL DE CUSTO| TOTAL DO LUCRO|% S/CUST|% S/VEND| EMPRESA                                      |"
    BioImprime "@Printer.Print " & "+----------------------+---------------+---------------+---------------+--------+--------+----------------------------------------------+"
End Sub

Private Sub ImpCab4()
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"

    BioImprime "@Printer.Print " & "+--------------------------------+----+-------------------+--------------+--------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                        |PROD|   LITROS ENTRADA  |TOTAL ENTRADA |  CUSTO MÉDIO |"
    BioImprime "@Printer.Print " & "+--------------------------------+----+-------------------+--------------+--------------+"

End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
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
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf chkVendaCombustivel.Value = 0 And chkVendaProduto.Value = 0 Then
        MsgBox "Escolha um tipo de venda.", vbInformation, "Atenção!"
        chkVendaCombustivel.SetFocus
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
        cbo_periodo_f.ListIndex = 3
        msk_data_i.SetFocus
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
        cbo_periodo_i.SetFocus
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
