VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lstPrecoMedioSemanal 
   Caption         =   "Preço Médio Semanal"
   ClientHeight    =   2940
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "lstPrecoMedioSemanal.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lstPrecoMedioSemanal.frx":030A
   ScaleHeight     =   2940
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "lstPrecoMedioSemanal.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza Preço Médio Semanal."
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "lstPrecoMedioSemanal.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime Preço Médio Semanal."
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "lstPrecoMedioSemanal.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1980
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lstPrecoMedioSemanal.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lstPrecoMedioSemanal.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lstPrecoMedioSemanal.frx":6CBA
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
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lstPrecoMedioSemanal"
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
Dim lValorVendaBombaA As Currency
Dim lValorVendaBombaAA As Currency
Dim lValorVendaBombaD As Currency
Dim lValorVendaBombaDA As Currency
Dim lValorVendaBombaG As Currency
Dim lValorVendaBombaGA As Currency
Dim lQtdVendaECFA As Currency
Dim lQtdVendaECFAA As Currency
Dim lQtdVendaECFD As Currency
Dim lQtdVendaECFDA As Currency
Dim lQtdVendaECFG As Currency
Dim lQtdVendaECFGA As Currency
Dim lValorVendaECFA As Currency
Dim lValorVendaECFAA As Currency
Dim lValorVendaECFD As Currency
Dim lValorVendaECFDA As Currency
Dim lValorVendaECFG As Currency
Dim lValorVendaECFGA As Currency
Dim lQtdDiferencaA As Currency
Dim lQtdDiferencaAA As Currency
Dim lQtdDiferencaD As Currency
Dim lQtdDiferencaDA As Currency
Dim lQtdDiferencaG As Currency
Dim lQtdDiferencaGA As Currency
Dim lValorDiferencaA As Currency
Dim lValorDiferencaAA As Currency
Dim lValorDiferencaD As Currency
Dim lValorDiferencaDA As Currency
Dim lValorDiferencaG As Currency
Dim lValorDiferencaGA As Currency

Dim lSubQtdEntradaComb As Currency
Dim lSubValorEntradaComb As Currency
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

Dim lSQL As String
Dim rstCombustivel As New adodb.Recordset
Dim rstEntradaCombustivel As New adodb.Recordset
Dim rstVendaCombustivel As New adodb.Recordset
Dim lNomeTabelaEntradaCombustivel As String
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
Private Sub BuscaDadosCombustivel()
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & " FROM Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Codigo, Nome"
    Set rstCombustivel = New adodb.Recordset
    Set rstCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosEntradaCombustivel()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    xTextoAgrupar = xTextoAgrupar + ", Data"
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM(Quantidade) AS TotalQuantidade,"
    lSQL = lSQL & " ROUND(SUM([Valor da Entrada]) ,2) AS TotalEntrada"
    lSQL = lSQL & "  FROM " & lNomeTabelaEntradaCombustivel
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstEntradaCombustivel = New adodb.Recordset
    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosVendaCombustivel()
    Dim xTextoAgrupar As String
    
    xTextoAgrupar = "[Tipo de Combustivel]"
    xTextoAgrupar = xTextoAgrupar + ", Data"
    lSQL = ""
    lSQL = lSQL & "SELECT " & xTextoAgrupar & ", SUM([Quantidade da Saida]) AS TotalQuantidade,"
    lSQL = lSQL & " ROUND(SUM([Quantidade da Saida] * [Preco de Venda]) ,2) AS TotalVenda"
    lSQL = lSQL & "  FROM " & lNomeTabelaMovimentoBomba
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY " & xTextoAgrupar
    lSQL = lSQL & " ORDER BY " & xTextoAgrupar
    Set rstVendaCombustivel = New adodb.Recordset
    Set rstVendaCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MedicaoCombustivel = Nothing
    
    Set rstCombustivel = Nothing
    Set rstEntradaCombustivel = Nothing
    Set rstVendaCombustivel = Nothing
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
    xLinha = "| PREÇO MÉDIO POR SEMANA                                                                                            Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 3, 40) = "PREÇO MÉDIO POR SEMANA (L.M.C.)        "
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        Mid(xLinha, 3, 40) = "PREÇO MÉDIO POR SEMANA (PISTA)         "
    Else
        Mid(xLinha, 3, 40) = "PREÇO MÉDIO POR SEMANA                 "
    End If
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____  SOMATORIA FISCAL:                                                                              |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    xString = ""
    Mid(xLinha, 60, 40) = xString
    BioImprime "@Printer.Print " & xLinha
    'xLinha = "| PRODUTO.....:                                                                                                                         |"
    'Mid(xLinha, 17, 30) = Mid(cbo_combustivel, 6, Len(cbo_combustivel))
    'BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DATA SEMANA| DATA SEMANA| COMBUSTIVEL         |VENDA  BOMBAS|VENDA  BOMBAS|   PREÇO  |   PREÇO  |   PREÇO  |   ENTRADAS  |   ENTRADAS  |   PREÇO  |   PREÇO  |   PREÇO  |           |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  INICIAL   |    FINAL   |                     |  EM LITROS  | EM  VALORES |  MÍNIMO  |  MÁXIMO  |   MÉDIO  |  EM LITROS  | EM  VALORES |  MÍNIMO  |  MÁXIMO  |   MÉDIO  |           |"
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
    Dim xPrecoMinimo As Currency
    Dim xPrecoMaximo As Currency
    Dim xPrecoMedio As Currency
    Dim xPrecoMinimoEnt As Currency
    Dim xPrecoMaximoEnt As Currency
    Dim xPrecoMedioEnt As Currency
    Dim xQtdEntradaComb As Currency
    Dim xValorEntradaComb As Currency
    Dim xLucro As Currency
    
    Fim = False
    xPrecoMinimo = 0
    xPrecoMaximo = 0
    xPrecoMedio = 0
    
    'xQtdVendaBomba = MovimentoBomba.QuantidadeVendaData(g_empresa, pDataInicial, pDataFinal, pTipoCombustivel, 0)
    'xValorVendaBomba = MovimentoBomba.ValorVendaPeriodo(g_empresa, pDataInicial, pDataFinal, pTipoCombustivel, 1, 9)
    xQtdVendaBomba = 0
    xValorVendaBomba = 0
    If BuscaRegistroVendaCombustivel(pTipoCombustivel, pDataInicial, pDataFinal) Then
        If rstVendaCombustivel.RecordCount > 0 Then
            xPrecoMinimo = 0
            xPrecoMaximo = 0
            xPrecoMedio = 0
            Do Until rstVendaCombustivel.EOF
                If xPrecoMinimo = 0 And rstVendaCombustivel!TotalQuantidade > 0 Then
                    xPrecoMinimo = Round(rstVendaCombustivel!TotalVenda / rstVendaCombustivel!TotalQuantidade, 4)
                    xPrecoMaximo = xPrecoMinimo
                    xPrecoMedio = xPrecoMinimo
                End If
                xQtdVendaBomba = xQtdVendaBomba + rstVendaCombustivel!TotalQuantidade
                xValorVendaBomba = xValorVendaBomba + rstVendaCombustivel!TotalVenda
                If rstVendaCombustivel!TotalQuantidade > 0 Then
                    If xPrecoMaximo < Round(rstVendaCombustivel!TotalVenda / rstVendaCombustivel!TotalQuantidade, 4) Then
                        xPrecoMaximo = Round(rstVendaCombustivel!TotalVenda / rstVendaCombustivel!TotalQuantidade, 4)
                    End If
                    xPrecoMedio = Round(xValorVendaBomba / xQtdVendaBomba, 4)
                End If
                rstVendaCombustivel.MoveNext
            Loop
        End If
    End If
    'xQtdVendaBomba = xQtdVendaBomba - MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, pDataInicial, pDataFinal, 1, 9, pTipoCombustivel, "")
    'xValorVendaBomba = xValorVendaBomba - MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, pDataInicial, pDataFinal, 1, 9, pTipoCombustivel, "")
'    If BuscaRegistroAfericao(pTipoCombustivel, pDataInicial) Then
'        xQtdVendaBomba = xQtdVendaBomba - rstAfericao!TotalQuantidade
'        xValorVendaBomba = xValorVendaBomba - rstAfericao!TotalVenda
'    End If
    xValorDescontoLMC = 0
    xValorDescontoUnitario = 0
    
'    If chkSemDesconto.Value = 0 Then
'        'xValorDescontoLMC = MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, pDataInicial + 1, pDataFinal + 1, pTipoCombustivel)
'        If BuscaRegistroMedicaoCombustivel(pTipoCombustivel, pDataInicial + 1) Then
'            xValorDescontoLMC = rstMedicaoCombustivel!TotalDesconto
'        End If
'        xValorVendaBomba = xValorVendaBomba - xValorDescontoLMC
'    End If
            
    xQtdVendaECF = 0
    xValorVendaECF = 0
    xValorDescontoECF = 0
    
    xValorVendaECF = xValorVendaECF - xValorDescontoECF
    xQtdDiferenca = xQtdVendaBomba - xQtdVendaECF
    xValorDiferenca = xValorVendaBomba - xValorVendaECF

    'Sub-Total
    lSubQtdVendaBomba = lSubQtdVendaBomba + xQtdVendaBomba
    lSubValorVendaBomba = lSubValorVendaBomba + xValorVendaBomba
    lSubQtdVendaECF = lSubQtdVendaECF + xQtdVendaECF
    lSubValorVendaECF = lSubValorVendaECF + xValorVendaECF
    lSubQtdDiferenca = lSubQtdDiferenca + xQtdDiferenca
    lSubValorDiferenca = lSubValorDiferenca + xValorDiferenca


    '                  1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20        21
    '         1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901
    xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
    Mid(xLinha, 3, 10) = Format(pDataInicial, "dd/mm/yyyy")
    Mid(xLinha, 16, 10) = Format(pDataFinal, "dd/mm/yyyy")
    Mid(xLinha, 29, 20) = pNomeCombustivel
    
    i = Len(Format(xQtdVendaBomba, "####,##0.00"))
    Mid(xLinha, 51 + 11 - i, i) = Format(xQtdVendaBomba, "####,##0.00")
    i = Len(Format(xValorVendaBomba, "####,##0.00"))
    Mid(xLinha, 65 + 11 - i, i) = Format(xValorVendaBomba, "####,##0.00")
    
    i = Len(Format(xPrecoMinimo, "##0.0000"))
    Mid(xLinha, 79 + 8 - i, i) = Format(xPrecoMinimo, "##0.0000")
    
    i = Len(Format(xPrecoMaximo, "##0.0000"))
    Mid(xLinha, 90 + 8 - i, i) = Format(xPrecoMaximo, "##0.0000")
    
    i = Len(Format(xPrecoMedio, "##0.0000"))
    Mid(xLinha, 101 + 8 - i, i) = Format(xPrecoMedio, "##0.0000")
    
    
    
    xQtdEntradaComb = 0
    xValorEntradaComb = 0
    If BuscaRegistroEntradaCombustivel(pTipoCombustivel, pDataInicial, pDataFinal) Then
        If rstEntradaCombustivel.RecordCount > 0 Then
            xPrecoMinimoEnt = Round(rstEntradaCombustivel!TotalEntrada / rstEntradaCombustivel!TotalQuantidade, 4)
            xPrecoMaximoEnt = xPrecoMinimoEnt
            xPrecoMedioEnt = xPrecoMinimoEnt
            Do Until rstEntradaCombustivel.EOF
                xQtdEntradaComb = xQtdEntradaComb + rstEntradaCombustivel!TotalQuantidade
                xValorEntradaComb = xValorEntradaComb + rstEntradaCombustivel!TotalEntrada
                If xPrecoMaximoEnt < Round(rstEntradaCombustivel!TotalEntrada / rstEntradaCombustivel!TotalQuantidade, 4) Then
                    xPrecoMaximoEnt = Round(rstEntradaCombustivel!TotalEntrada / rstEntradaCombustivel!TotalQuantidade, 4)
                End If
                xPrecoMedioEnt = Round(xValorEntradaComb / xQtdEntradaComb, 4)
                rstEntradaCombustivel.MoveNext
            Loop
        End If
    End If
    i = Len(Format(xQtdEntradaComb, "####,##0.00"))
    Mid(xLinha, 112 + 11 - i, i) = Format(xQtdEntradaComb, "####,##0.00")
    i = Len(Format(xValorEntradaComb, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(xValorEntradaComb, "####,##0.00")
    
    
    
    i = Len(Format(xPrecoMinimoEnt, "##0.0000"))
    Mid(xLinha, 140 + 8 - i, i) = Format(xPrecoMinimoEnt, "##0.0000")
    
    i = Len(Format(xPrecoMaximoEnt, "##0.0000"))
    Mid(xLinha, 151 + 8 - i, i) = Format(xPrecoMaximoEnt, "##0.0000")
    
    i = Len(Format(xPrecoMedioEnt, "##0.0000"))
    Mid(xLinha, 162 + 8 - i, i) = Format(xPrecoMedioEnt, "##0.0000")
    
    xLucro = xPrecoMedio - xPrecoMedioEnt
    i = Len(Format(xLucro, "##0.0000"))
    Mid(xLinha, 174 + 8 - i, i) = Format(xLucro, "##0.0000")
    
    'Sub-Total Entradas
    lSubQtdEntradaComb = lSubQtdEntradaComb + xQtdEntradaComb
    lSubValorEntradaComb = lSubValorEntradaComb + xValorEntradaComb
    
    
'    i = Len(Format(xQtdVendaECF, "####,##0.00"))
'    Mid(xLinha, 79 + 11 - i, i) = Format(xQtdVendaECF, "####,##0.00")
'    i = Len(Format(xValorVendaECF, "####,##0.00"))
'    Mid(xLinha, 93 + 11 - i, i) = Format(xValorVendaECF, "####,##0.00")
'    i = Len(Format(xQtdDiferenca, "####,##0.00"))
'    Mid(xLinha, 107 + 11 - i, i) = Format(xQtdDiferenca, "####,##0.00")
'    i = Len(Format(xValorDiferenca, "####,##0.00"))
'    Mid(xLinha, 121 + 11 - i, i) = Format(xValorDiferenca, "####,##0.00")
        
    
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
    End If
    
    If xQtdVendaBomba > 0 Then
        xValor = Format(xValorVendaBomba / xQtdVendaBomba, "0000000.0000")
    Else
        xValor = 0
    End If
'    i = Len(Format(xValor, "##,##0.0000"))
'    Mid(xLinha, 135 + 11 - i, i) = Format(xValor, "##,##0.0000")
'
'    If chkImprimirDescontoUnitario.Value = 1 Then
'        Mid(xLinha, 135, 11) = "           "
'        If xQtdVendaECF > 0 And xQtdVendaBomba > 0 Then
'            i = Len(Format(xValorDescontoUnitario, "##,##0.0000"))
'            Mid(xLinha, 122 + 11 - i, i) = Format(xValorDescontoUnitario, "##,##0.0000")
'        End If
'    End If
'    If g_lmc = 2 And xQtdDiferenca > 0 And xValorVendaECF > 0 And xValorDiferenca <> 0 Then
'        xValor = xQtdDiferenca * (Format(xValorVendaECF / xQtdVendaECF, "000000.0000"))
'        If xValor <> xValorDiferenca Then
'            xValor = xValorDiferenca - xValor
'            xLinha = xLinha & " " & Format(xValor, "####,##0.00")
'        End If
'    End If
    
    'i = Len(Format(xValorDescontoECF, "####,##0.00"))
    'Mid(xLinha, 122 + 11 - i, i) = Format(xValorDescontoECF, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20        21
    '         1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901
    xLinha = "|                                               |             |             |          |          |          |             |             |          |          |          |           |"
    Mid(xLinha, 30, 13) = "**  SUB-TOTAL"
    i = Len(Format(lSubQtdVendaBomba, "####,##0.00"))
    Mid(xLinha, 51 + 11 - i, i) = Format(lSubQtdVendaBomba, "####,##0.00")
    i = Len(Format(lSubValorVendaBomba, "####,##0.00"))
    Mid(xLinha, 65 + 11 - i, i) = Format(lSubValorVendaBomba, "####,##0.00")
    
    i = Len(Format(lSubQtdEntradaComb, "####,##0.00"))
    Mid(xLinha, 112 + 11 - i, i) = Format(lSubQtdEntradaComb, "####,##0.00")
    i = Len(Format(lSubValorEntradaComb, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(lSubValorEntradaComb, "####,##0.00")
    
    
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    If lLinha >= 62 Then
        xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    Else
        xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    If lQtdVendaBombaA > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("A ")
        i = Len(Format(lQtdVendaBombaA, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaA, "####,##0.00")
        i = Len(Format(lValorVendaBombaA, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaAA > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("AA")
        i = Len(Format(lQtdVendaBombaAA, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaAA, "####,##0.00")
        i = Len(Format(lValorVendaBombaAA, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaAA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaD > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("D ")
        i = Len(Format(lQtdVendaBombaD, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaD, "####,##0.00")
        i = Len(Format(lValorVendaBombaD, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaD, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaDA > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("DA")
        i = Len(Format(lQtdVendaBombaDA, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaDA, "####,##0.00")
        i = Len(Format(lValorVendaBombaDA, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaDA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaG > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("G ")
        i = Len(Format(lQtdVendaBombaG, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaG, "####,##0.00")
        i = Len(Format(lValorVendaBombaG, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaG, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    If lQtdVendaBombaGA > 0 Then
        xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
        Mid(xLinha, 29, 20) = BuscaTipoCombustivel("GA")
        i = Len(Format(lQtdVendaBombaGA, "####,##0.00"))
        Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaGA, "####,##0.00")
        i = Len(Format(lValorVendaBombaGA, "####,##0.00"))
        Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaGA, "####,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    xLinha = "|            |            |                     |             |             |          |          |          |             |             |          |          |          |           |"
    Mid(xLinha, 29, 20) = "TOTAL GERAL"
    lQtdVendaBombaA = lQtdVendaBombaA + lQtdVendaBombaAA + lQtdVendaBombaD + lQtdVendaBombaDA + lQtdVendaBombaG + lQtdVendaBombaGA
    i = Len(Format(lQtdVendaBombaA, "####,##0.00"))
    Mid(xLinha, 51 + 11 - i, i) = Format(lQtdVendaBombaA, "####,##0.00")
    lValorVendaBombaA = lValorVendaBombaA + lValorVendaBombaAA + lValorVendaBombaD + lValorVendaBombaDA + lValorVendaBombaG + lValorVendaBombaGA
    i = Len(Format(lValorVendaBombaA, "####,##0.00"))
    Mid(xLinha, 65 + 11 - i, i) = Format(lValorVendaBombaA, "####,##0.00")
    
    BioImprime "@Printer.Print " & xLinha
    
    
    xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
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
    
    
    xLinha = "|         ** |                     |             |             |             |             |             |             |             |  |"
    Mid(xLinha, 16, 20) = "TOTAL GERAL"
    lTotalDescontoA = lTotalDescontoA + lTotalDescontoAA + lTotalDescontoD + lTotalDescontoDA + lTotalDescontoG + lTotalDescontoGA
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
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xDiaDaSemana As Integer
    
    ZeraVariaveis
    BuscaDadosCombustivel
    BuscaDadosVendaCombustivel
    BuscaDadosEntradaCombustivel
    xData = CDate(msk_data_i.Text)
    'Loop data
    Do Until xData > CDate(msk_data_f.Text)
        lSubQtdEntradaComb = 0
        lSubValorEntradaComb = 0
        lSubQtdVendaBomba = 0
        lSubValorVendaBomba = 0
        lSubQtdVendaECF = 0
        lSubValorVendaECF = 0
        lSubQtdDiferenca = 0
        lSubValorDiferenca = 0
        xDiaDaSemana = Weekday(xData)
        xDataInicial = xData
        xDataFinal = xData + (7 - xDiaDaSemana)
        If xDataFinal > CDate(msk_data_f.Text) Then
            xDataFinal = CDate(msk_data_f.Text)
        End If
        Call LoopCombustivel(xDataInicial, xDataFinal)
        'Call ImpDet(xData)
        'If chkDetalhadaBico.Value = 1 Then
        '    Call ImpDetBico(xData)
        'End If
        xData = xDataFinal + 1
        ImpSubTotal
    Loop
    ImpTotal
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Preço Médio Semanal.|@|"
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
    
    lSubQtdEntradaComb = 0
    lSubValorEntradaComb = 0
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
        lNomeTabelaEntradaCombustivel = "Entrada_Combustivel_LMC"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba_LMC"
        Me.Caption = Me.Caption & " - LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        lNomeTabelaEntradaCombustivel = "Entrada_Combustivel"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba_Cupom"
        Me.Caption = Me.Caption & " - ECF"
    Else
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        lNomeTabelaEntradaCombustivel = "Entrada_Combustivel"
        lNomeTabelaMovimentoBomba = "Movimento_Bomba"
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
Private Function BuscaRegistroEntradaCombustivel(ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND Data >= " & preparaData(pDataInicial)
    xCondicao = xCondicao & " AND Data <= " & preparaData(pDataFinal)
    
    BuscaRegistroEntradaCombustivel = False
    rstEntradaCombustivel.Filter = ""
    If rstEntradaCombustivel.RecordCount > 0 Then
        rstEntradaCombustivel.MoveFirst
        rstEntradaCombustivel.Filter = xCondicao
        If Not rstEntradaCombustivel.EOF Then
            BuscaRegistroEntradaCombustivel = True
        End If
    End If
End Function
Private Function BuscaRegistroVendaCombustivel(ByVal pTipoCombustivel As String, ByVal pDataInicial As Date, ByVal pDataFinal As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND Data >= " & preparaData(pDataInicial)
    xCondicao = xCondicao & " AND Data <= " & preparaData(pDataFinal)
    
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
        xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "+------------+------------+---------------------+-------------+-------------+----------+----------+----------+-------------+-------------+----------+----------+----------+-----------+"
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
