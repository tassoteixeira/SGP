VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_analise_vendas_funcionarios 
   Caption         =   "Emissão da Análise de Vendas dos Funcionários"
   ClientHeight    =   3750
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "lst_analise_vendas_funcionarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_analise_vendas_funcionarios.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Visualiza análise de vendas dos funcionários."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "lst_analise_vendas_funcionarios.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprime análise de vendas dos funcionários."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "lst_analise_vendas_funcionarios.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo do Posto"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   2340
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo da Loja"
         Height          =   255
         Left            =   3660
         TabIndex        =   20
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_vendas_funcionarios.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_vendas_funcionarios.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6060
         Picture         =   "lst_analise_vendas_funcionarios.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   2175
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1920
         Width           =   4875
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4920
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
         Caption         =   "I&mprimir Produto"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1500
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
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   1035
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
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_vendas_funcionarios"
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
Dim l_vendas As Currency
Dim l_indice As Currency
Dim l_percentual As Currency
Dim l_premiacao As Currency
Dim l_sub_vendas As Currency
Dim l_sub_premiacao As Currency
Dim l_total_vendas As Currency
Dim l_total_premiacao As Currency
Dim lSQL As String

Dim Premiacao As New cPremiacao

Dim rstEmpresa As adodb.Recordset
Dim rstFuncionario As adodb.Recordset
Dim rstTabela As adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Premiacao = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    l_vendas = 0
    l_indice = 0
    l_percentual = 0
    l_premiacao = 0
    l_sub_vendas = 0
    l_sub_premiacao = 0
    l_total_vendas = 0
    l_total_premiacao = 0
End Sub
Private Sub PreencheCboGrupo()
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    
    lSQL = "SELECT Codigo, Nome FROM Grupo ORDER BY Nome"
    Set rstTabela = Conectar.RsConexao(lSQL)
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_grupo.AddItem !Nome
                cbo_grupo.ItemData(cbo_grupo.NewIndex) = !Codigo
                .MoveNext
            Loop
            rstTabela.Close
        End If
    End With
    Set rstTabela = Nothing
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
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "Todos os Caixas"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 0
    cbo_tipo_movimento.AddItem "Caixa de Óleo/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "Caixa da Borr./Lavador"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    Call LoopEmpresa
    rstEmpresa.Close
    Set rstEmpresa = Nothing
    If l_total_vendas > 0 Then
        rstFuncionario.Close
        Set rstFuncionario = Nothing
        Call ImpTotal(l_total_vendas, l_total_premiacao)
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Análise de Vendas dos Funcionários|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Private Sub LoopFuncionario(ByVal pEmpresa As Integer, ByVal pNomeEmpresa As String)
    Dim x_custo As Currency
    Dim x_venda As Currency
    
    lSQL = "SELECT Codigo, Nome, Cargo FROM Funcionario WHERE Empresa = " & pEmpresa & " ORDER BY Nome"
    Set rstFuncionario = Conectar.RsConexao(lSQL)
    With rstFuncionario
        If .RecordCount > 0 Then
            Do Until .EOF
                l_vendas = 0
                l_indice = 0
                l_percentual = 0
                l_premiacao = 0
                Call ImpMovimentoLubrificante(pEmpresa, !Codigo)
                If l_vendas > 0 Then
                    If Premiacao.LocalizarCodigo(g_empresa, CDate("01/" & Format(msk_data_i.Text, "mm") & "/" & Format(msk_data_i.Text, "yyyy"))) Then
                        l_indice = l_vendas * 100 / Premiacao.ValorBase
                        If l_indice >= Premiacao.PercentualBase1 Then
                            l_percentual = Premiacao.PercentualComissao1
                            l_premiacao = l_vendas * Premiacao.PercentualComissao1 / 100
                        ElseIf l_indice >= Premiacao.PercentualBase2 Then
                            l_percentual = Premiacao.PercentualComissao2
                            l_premiacao = l_vendas * Premiacao.PercentualComissao2 / 100
                        ElseIf l_indice >= Premiacao.PercentualBase3 Then
                            l_percentual = Premiacao.PercentualComissao3
                            l_premiacao = l_vendas * Premiacao.PercentualComissao3 / 100
                        End If
                    End If
                    If g_nome_empresa Like "*POSTO MOREIRA COSTA*" And UCase(Trim(!Cargo)) Like "TROCADOR DE OLEO" Then
                        Premiacao.ValorBase = 10000
                        Premiacao.PercentualBase1 = 100
                        Premiacao.PercentualComissao1 = 10
                        Premiacao.PercentualBase2 = 50.01
                        Premiacao.PercentualComissao2 = 7
                        Premiacao.PercentualBase3 = 1
                        Premiacao.PercentualComissao3 = 5
                        l_indice = l_vendas * 100 / Premiacao.ValorBase
                        If l_indice >= Premiacao.PercentualBase1 Then
                            l_percentual = Premiacao.PercentualComissao1
                            l_premiacao = l_vendas * Premiacao.PercentualComissao1 / 100
                        ElseIf l_indice >= Premiacao.PercentualBase2 Then
                            l_percentual = Premiacao.PercentualComissao2
                            l_premiacao = l_vendas * Premiacao.PercentualComissao2 / 100
                        ElseIf l_indice >= Premiacao.PercentualBase3 Then
                            l_percentual = Premiacao.PercentualComissao3
                            l_premiacao = l_vendas * Premiacao.PercentualComissao3 / 100
                        End If
                    End If
                    Call ImpDet(!Codigo, !Nome, l_vendas, l_indice, l_percentual, l_premiacao, pNomeEmpresa)
                    l_sub_premiacao = l_sub_premiacao + l_premiacao
                    l_total_premiacao = l_total_premiacao + l_premiacao
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpMovimentoLubrificante(ByVal pEmpresa As Integer, ByVal pFuncionario As Integer)
    
    lSQL = "SELECT Produto.[Codigo do Grupo],"
    lSQL = lSQL & " SUM([Valor Total]) AS TotalVenda"
    lSQL = lSQL & " FROM Movimento_Lubrificante, Produto"
    lSQL = lSQL & " WHERE Empresa = " & pEmpresa
    lSQL = lSQL & " AND [Codigo do Funcionario] = " & pFuncionario
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & " AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & " AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
    If chkExclusivoLoja.Value = 1 And chkExclusivoPosto.Value = 0 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
    If chkExclusivoLoja.Value = 0 And chkExclusivoPosto.Value = 1 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    lSQL = lSQL & " GROUP BY [Codigo do Grupo]"
    Set rstTabela = Conectar.RsConexao(lSQL)
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                If cbo_grupo.ItemData(cbo_grupo.ListIndex) = ![Codigo do Grupo] Or cbo_grupo.ItemData(cbo_grupo.ListIndex) = 0 Then
                    l_vendas = l_vendas + !TotalVenda
                    l_sub_vendas = l_sub_vendas + !TotalVenda
                    l_total_vendas = l_total_vendas + !TotalVenda
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresa()
    Dim i As Integer
    Dim x_litro As Currency
    Dim x_venda As Currency
    Dim x_custo As Currency
    
    lSQL = "SELECT Codigo, Nome, Inativo FROM Empresas ORDER BY Codigo"
    Set rstEmpresa = Conectar.RsConexao(lSQL)
    With rstEmpresa
        .MoveFirst
        Do Until .EOF
            If !Codigo > 11 Then
                Exit Do
            End If
            l_sub_vendas = 0
            l_sub_premiacao = 0
            Call LoopFuncionario(!Codigo, !Nome)
            If l_sub_vendas > 0 Then
                Call ImpSubTotal(l_sub_vendas, l_sub_premiacao, !Nome)
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpSubTotal(x_vendas As Currency, x_premiacao As Currency, x_nome_empresa As String)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|     | ** TOTAL DA EMPRESA                      |                  |         |         |                  |                            |"
    i = Len(Format(x_vendas, "##,###,##0.00"))
    Mid(x_linha, 55 + 13 - i, i) = Format(x_vendas, "##,###,##0.00")
    i = Len(Format(x_premiacao, "##,###,##0.00"))
    Mid(x_linha, 94 + 13 - i, i) = Format(x_premiacao, "##,###,##0.00")
    Mid(x_linha, 110, 26) = x_nome_empresa
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-----+------------------------------------------+------------------+---------+---------+------------------+----------------------------+"
    lLinha = lLinha + 2
End Sub
Private Sub ImpTotal(x_vendas As Currency, x_premiacao As Currency)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|     | ** TOTAL GERAL DAS EMPRESAS              |                  |                   |                  |                            |"
    i = Len(Format(x_vendas, "##,###,##0.00"))
    Mid(x_linha, 55 + 13 - i, i) = Format(x_vendas, "##,###,##0.00")
    i = Len(Format(x_premiacao, "##,###,##0.00"))
    Mid(x_linha, 94 + 13 - i, i) = Format(x_premiacao, "##,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-----+------------------------------------------+------------------+-------------------+------------------+----------------------------+"
    Mid(x_linha, 11, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 2
End Sub
Private Sub ImpDet(x_codigo As Integer, x_nome As String, x_vendas As Currency, x_indice As Currency, x_percentual As Currency, x_premiacao As Currency, x_nome_empresa As String)
    Dim x_linha As String
    Dim i As Integer
    If lPagina = 0 Then
        Call ImpCab
    End If
    If lLinha >= 55 Then
        x_linha = "+-----+------------------------------------------+------------------+---------+---------+------------------+----------------------------+"
        Mid(x_linha, 11, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    x_linha = "|     |                                          |                  |       % |       % |                  |                            |"
    Mid(x_linha, 3, 3) = Format(x_codigo, "000")
    Mid(x_linha, 9, 40) = x_nome
    i = Len(Format(x_vendas, "##,###,##0.00"))
    Mid(x_linha, 55 + 13 - i, i) = Format(x_vendas, "##,###,##0.00")
    i = Len(Format(x_indice, "##0.00"))
    Mid(x_linha, 71 + 6 - i, i) = Format(x_indice, "##0.00")
    i = Len(Format(x_percentual, "##0.00"))
    Mid(x_linha, 81 + 6 - i, i) = Format(x_percentual, "##0.00")
    i = Len(Format(x_premiacao, "##,###,##0.00"))
    Mid(x_linha, 94 + 13 - i, i) = Format(x_premiacao, "##,###,##0.00")
    Mid(x_linha, 110, 26) = x_nome_empresa
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
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
    Printer.CurrentY = 0
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| ANÁLISE DAS VENDAS DOS FUNCIONÁRIOS                             , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____       PERÍODO.: _ AO _     |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    Mid(x_linha, 69, 1) = cbo_periodo_i
    Mid(x_linha, 74, 1) = cbo_periodo_f
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| TIPO DO MOVIMENTO.......:                                                    |"
    Mid(x_linha, 29, 30) = cbo_tipo_movimento
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| GRUPO...................:                                                    |"
    Mid(x_linha, 29, 3) = Format(cbo_grupo.ItemData(cbo_grupo.ListIndex), "000")
    Mid(x_linha, 33, 30) = cbo_grupo
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-----+------------------------------------------+------------------+---------+---------+------------------+----------------------------+"
    BioImprime "@Printer.Print " & "| COD.| NOME DO FUNCIONÁRIO                      | TOTAL DAS VENDAS |  ÍNDICE | PERCENT.|VALOR DA PREMIAÇÃO| EMPRESA                    |"
    BioImprime "@Printer.Print " & "|     |                                          |                  |ALCANCADO|PREMIAÇÃO|                  |                            |"
    BioImprime "@Printer.Print " & "+-----+------------------------------------------+------------------+---------+---------+------------------+----------------------------+"
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoPosto.SetFocus
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
Private Sub chkExclusivoLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chkExclusivoPosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoLoja.SetFocus
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
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Selecione o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", 64, "Atenção!"
        cbo_grupo.SetFocus
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
        cbo_tipo_movimento.ListIndex = 1
        cbo_grupo.ListIndex = 0
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
    PreencheCboTipoMovimento
    PreencheCboGrupo
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
