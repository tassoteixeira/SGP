VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_resumo_lmc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo do L.M.C."
   ClientHeight    =   6405
   ClientLeft      =   2775
   ClientTop       =   3795
   ClientWidth     =   5475
   Icon            =   "lst_resumo_lmc.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_resumo_lmc.frx":030A
   ScaleHeight     =   6405
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "lst_resumo_lmc.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Visualiza resumo do L.M.C."
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "lst_resumo_lmc.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime resumo do L.M.C."
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "lst_resumo_lmc.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5460
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.TextBox txtFornecedor 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   24
         Top             =   1920
         Width           =   3435
      End
      Begin VB.CheckBox chkDetalhadaTanque 
         Caption         =   "Venda Detalhada por Tanque"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   2580
         Width           =   2595
      End
      Begin VB.CheckBox chkRescunhoLmc 
         Caption         =   "Rascunho do LMC"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   3300
         Width           =   2295
      End
      Begin VB.TextBox txtPaginaInicial 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "001"
         Top             =   3900
         Width           =   675
      End
      Begin VB.CheckBox chkImprimeResumo 
         Caption         =   "Imprime Resumo do LMC"
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   4200
         Width           =   3135
      End
      Begin VB.CheckBox chkImprimeEntrada 
         Caption         =   "Imprime Entradas de Combustiveis"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   4560
         Width           =   3135
      End
      Begin VB.CheckBox chkCustoDoEstoque 
         Caption         =   "Totaliza Custo do Estoque"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CheckBox chkDetalhadaBico 
         Caption         =   "Venda Detalhada por Bico"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   2220
         Width           =   2295
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_resumo_lmc.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_resumo_lmc.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   3435
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
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_resumo_lmc.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "&Fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5220
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label4 
         Caption         =   "&Página Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   3900
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Combustível"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
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
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_resumo_lmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lPaginaOld As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de variáveis padrão para relatório
Dim lSQL As String
Dim l_tipo_combustivel As String
Dim l_total_recebido As Currency
Dim l_total_vendas As Currency
Dim l_total_valor_vendas As Currency
Dim l_total_afericao As Currency
Dim l_perdas_sobras As Currency
Dim lCustoMedio As Currency
Dim lSubCustoMedio As Currency
Dim lTotalCusto As Currency
Dim lParcialEntrada As Currency
Dim lParcialVenda As Currency
Dim lParcialAfericao As Currency
Dim lParcialPerdasSobras As Currency
Dim lParcialMes As Integer
'Dim lNumeroTanque As Integer
Dim lTodosCombustiveis As Boolean
Dim lObservacao(0 To 30) As String
Dim iObs As Integer
Dim rstMovimentoBomba As New adodb.Recordset
Dim rstMovimentoBombaNormal As New adodb.Recordset
Dim rstEntradaCombustivel As New adodb.Recordset
Dim rstNFeDestinada As New adodb.Recordset
Dim rstMovimentoAfericao As New adodb.Recordset
Dim rstMovimentoNFeDevolucao As New adodb.Recordset
Dim rstMedicaoCombustivel As New adodb.Recordset
Dim rstTanque As New adodb.Recordset

Private Combustivel As New cCombustivel
Private EntradaCombustivel As New cEntradaCombustivel
Private Fornecedor As New cFornecedor
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoBombaNormal As New cMovimentoBomba
Private lChaveAcessoNFeDestinadas As New Dictionary
Private lChaveAcessoNFeDestinadasNaoCadastradas As New Dictionary


Private Sub BuscaEntradaCombustivel()
    Dim xDadosParaTrazer As String
    
    xDadosParaTrazer = "[Data], [Tipo de Combustivel]"
    If chkRescunhoLmc.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero da Nota]"
    End If
    lSQL = ""
    lSQL = lSQL & "   SELECT " & xDadosParaTrazer
    lSQL = lSQL & "          , Sum(Quantidade) AS QTD"
    lSQL = lSQL & "     FROM " & EntradaCombustivel.NomeTabela
    lSQL = lSQL & "    WHERE empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    'tarefa redmine 567
    lSQL = lSQL & "      AND [Data] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND [Data] <= " & preparaData(CDate(msk_data_f.Text))
'    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
'    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " GROUP BY " & xDadosParaTrazer
    lSQL = lSQL & " ORDER BY " & xDadosParaTrazer
    'Abre RecordSet
    Set rstEntradaCombustivel = New adodb.Recordset
    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaMedicaoCombustivel()
    Dim xDadosParaTrazer As String
    
    xDadosParaTrazer = "Data, [Tipo de Combustivel]"
    If chkRescunhoLmc.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero do Tanque]"
    End If
'    xSQL = xSQL & " SELECT Sum([Desconto Dia Anterior]) AS TotalDesconto"
    
    lSQL = ""
    lSQL = lSQL & "   SELECT " & xDadosParaTrazer
    lSQL = lSQL & "          , Sum(Quantidade) AS QTD"
    lSQL = lSQL & "          , Sum([Desconto Dia Anterior]) AS TotalDesconto" 'new 05/08
    lSQL = lSQL & "     FROM " & MedicaoCombustivel.NomeTabela
    lSQL = lSQL & "    WHERE empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text) + 1)
    lSQL = lSQL & " GROUP BY " & xDadosParaTrazer
    lSQL = lSQL & " ORDER BY " & xDadosParaTrazer
    'Abre RecordSet
    Set rstMedicaoCombustivel = New adodb.Recordset
    Set rstMedicaoCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaMovimentoAfericao()
    Dim xDadosParaTrazer As String

'SUM([Valor Total]) as TotalValor
    xDadosParaTrazer = "Data, [Tipo de Combustivel], Periodo"
    If chkRescunhoLmc.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Codigo da Bomba]"
    End If
    

    lSQL = ""
    lSQL = lSQL & "   SELECT " & xDadosParaTrazer
    lSQL = lSQL & "          , Sum(Quantidade) AS QTD, SUM([Valor Total]) AS ValorTotal"
    lSQL = lSQL & "          , Sum([Preco de Custo] * Quantidade) as Total" 'new 05/08
    lSQL = lSQL & "     FROM " & MovimentoAfericao.NomeTabela
    lSQL = lSQL & "    WHERE empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " GROUP BY " & xDadosParaTrazer
    lSQL = lSQL & " ORDER BY " & xDadosParaTrazer
    'Abre RecordSet
    Set rstMovimentoAfericao = New adodb.Recordset
    Set rstMovimentoAfericao = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaMovimentoBomba()
    Dim xDadosParaTrazer As String
    
    xDadosParaTrazer = "Data, [Tipo de Combustivel]"
    If chkDetalhadaBico.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", Periodo, [Codigo da Bomba]"
    End If
    If chkDetalhadaTanque.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero do Tanque]"
    End If
    If chkRescunhoLmc.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero do Tanque], [Codigo da Bomba], Periodo, Abertura, Encerrante, [Quantidade da Saida], [Preco de Venda]"
    End If
    
    lSQL = ""
    lSQL = lSQL & "   SELECT " & xDadosParaTrazer
    lSQL = lSQL & "          , Sum([Quantidade da Saida]) AS QTD"
    lSQL = lSQL & "          , Sum([Quantidade da Saida] * [Preco de Venda]) AS ValorVenda" 'new 31/07
    lSQL = lSQL & "          , Sum([Quantidade da Saida] * [Preco de Custo]) AS TotalCustoVenda" 'new 05/08
    lSQL = lSQL & "     FROM " & MovimentoBomba.NomeTabela
    lSQL = lSQL & "    WHERE empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    'lSQL = lSQL & "      AND [Numero do Tanque] = " & preparaTexto(xNumeroTanque)
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " GROUP BY " & xDadosParaTrazer
    lSQL = lSQL & " ORDER BY " & xDadosParaTrazer
    'Abre RecordSet
    Set rstMovimentoBomba = New adodb.Recordset
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaMovimentoBombaNormal()
    Dim xDadosParaTrazer As String
    
    xDadosParaTrazer = "Data, [Tipo de Combustivel]"
    If chkDetalhadaBico.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", Periodo, [Codigo da Bomba]"
    End If
    If chkDetalhadaTanque.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero do Tanque]"
    End If
    If chkRescunhoLmc.Value = 1 Then
        xDadosParaTrazer = xDadosParaTrazer & ", [Numero do Tanque], [Codigo da Bomba], Periodo, Abertura, Encerrante, [Quantidade da Saida], [Preco de Venda]"
    End If
    
    lSQL = ""
    lSQL = lSQL & "   SELECT " & xDadosParaTrazer
    lSQL = lSQL & "          , Sum([Quantidade da Saida]) AS QTD"
    lSQL = lSQL & "     FROM Movimento_Bomba"
    lSQL = lSQL & "    WHERE empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    'lSQL = lSQL & "      AND [Numero do Tanque] = " & preparaTexto(xNumeroTanque)
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " GROUP BY " & xDadosParaTrazer
    lSQL = lSQL & " ORDER BY " & xDadosParaTrazer
    'Abre RecordSet
    Set rstMovimentoBombaNormal = New adodb.Recordset
    Set rstMovimentoBombaNormal = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaMovimentoNFeDevolucao()
    lSQL = ""
    lSQL = lSQL & "   SELECT Data, Serie, Numero, Quantidade, [Tipo de Combustivel]"
    lSQL = lSQL & "     FROM MovimentoNotaFiscalSaidaItem"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND (CFOP = " & preparaTexto("5661")
    lSQL = lSQL & "            OR CFOP = " & preparaTexto("6661") & ")"
    lSQL = lSQL & "      AND Cancelado = " & preparaBooleano(False)
    lSQL = lSQL & " ORDER BY Data, Serie, Numero"
    Set rstMovimentoNFeDevolucao = New adodb.Recordset
    Set rstMovimentoNFeDevolucao = Conectar.RsConexao(lSQL)
End Sub

Private Function TotalMedidaCombustivel(ByVal pData As Date, ByVal pTipoCombustivel As String) As Currency
    TotalMedidaCombustivel = 0
    
    If rstMedicaoCombustivel.RecordCount > 0 Then
        rstMedicaoCombustivel.MoveFirst
        rstMedicaoCombustivel.Find "Data = " & preparaData(pData)
        If Not rstMedicaoCombustivel.EOF Then
            Do Until rstMedicaoCombustivel.EOF
                If rstMedicaoCombustivel!Data = pData Then
                    TotalMedidaCombustivel = TotalMedidaCombustivel + rstMedicaoCombustivel!qtd
                Else
                    Exit Do
                End If
                rstMedicaoCombustivel.MoveNext
            Loop
        End If
    End If
End Function
Private Function TotalDescontoCombustivel_MedicaoCombustivel(ByVal pData As Date, ByVal pTipoCombustivel As String) As Currency
    TotalDescontoCombustivel_MedicaoCombustivel = 0
    
    If rstMedicaoCombustivel.RecordCount > 0 Then
        rstMedicaoCombustivel.MoveFirst
        rstMedicaoCombustivel.Find "Data = " & preparaData(pData)
        If Not rstMedicaoCombustivel.EOF Then
            Do Until rstMedicaoCombustivel.EOF
                If rstMedicaoCombustivel!Data = pData Then
                    TotalDescontoCombustivel_MedicaoCombustivel = TotalDescontoCombustivel_MedicaoCombustivel + rstMedicaoCombustivel!TotalDesconto
                Else
                    Exit Do
                End If
                rstMedicaoCombustivel.MoveNext
            Loop
        End If
    End If
End Function
Private Function TotalEntradaPeriodo(ByVal pData As Date, ByVal pTipoCombustivel As String) As Currency
    TotalEntradaPeriodo = 0
    
    If rstEntradaCombustivel.RecordCount > 0 Then
        rstEntradaCombustivel.MoveFirst
        'tarefa redmine 567
        rstEntradaCombustivel.Find "[Data] = " & preparaData(pData)
        'rstEntradaCombustivel.Find "Data = " & preparaData(pData)
        If Not rstEntradaCombustivel.EOF Then
            Do Until rstEntradaCombustivel.EOF
                If rstEntradaCombustivel![Data] = pData Then
                'If rstEntradaCombustivel!Data = pData Then
                    TotalEntradaPeriodo = TotalEntradaPeriodo + rstEntradaCombustivel!qtd
                Else
                    Exit Do
                End If
                rstEntradaCombustivel.MoveNext
            Loop
        End If
    End If
End Function
Private Function TotalQtdPeriodoCombustivel(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pPeriodo As Integer) As Currency
    TotalQtdPeriodoCombustivel = 0
    
    If rstMovimentoAfericao.RecordCount > 0 Then
        rstMovimentoAfericao.MoveFirst
        rstMovimentoAfericao.Find "Data = " & preparaData(pData)
        If Not rstMovimentoAfericao.EOF Then
            Do Until rstMovimentoAfericao.EOF
                If rstMovimentoAfericao!Data = pData Then
                    If pPeriodo = 0 Or rstMovimentoAfericao!Periodo = pPeriodo Then
                        TotalQtdPeriodoCombustivel = TotalQtdPeriodoCombustivel + rstMovimentoAfericao!qtd
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoAfericao.MoveNext
            Loop
        End If
    End If
End Function
Private Function TotalCustoCombustivel_MovimentoAfericao(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pPeriodo As Integer) As Currency
    TotalCustoCombustivel_MovimentoAfericao = 0
    
    If rstMovimentoAfericao.RecordCount > 0 Then
        rstMovimentoAfericao.MoveFirst
        rstMovimentoAfericao.Find "Data = " & preparaData(pData)
        If Not rstMovimentoAfericao.EOF Then
            Do Until rstMovimentoAfericao.EOF
                If rstMovimentoAfericao!Data = pData Then
                    If pPeriodo = 0 Or rstMovimentoAfericao!Periodo = pPeriodo Then
                        TotalCustoCombustivel_MovimentoAfericao = TotalCustoCombustivel_MovimentoAfericao + rstMovimentoAfericao!total
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoAfericao.MoveNext
            Loop
        End If
    End If
End Function
Private Function ValorTotalCombustivel_MovimentoAfericao(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pPeriodo As Integer) As Currency
    ValorTotalCombustivel_MovimentoAfericao = 0
    
    If rstMovimentoAfericao.RecordCount > 0 Then
        rstMovimentoAfericao.MoveFirst
        rstMovimentoAfericao.Find "Data = " & preparaData(pData)
        If Not rstMovimentoAfericao.EOF Then
            Do Until rstMovimentoAfericao.EOF
                If rstMovimentoAfericao!Data = pData Then
                    If pPeriodo = 0 Or rstMovimentoAfericao!Periodo = pPeriodo Then
                        ValorTotalCombustivel_MovimentoAfericao = ValorTotalCombustivel_MovimentoAfericao + rstMovimentoAfericao!ValorTotal
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoAfericao.MoveNext
            Loop
        End If
    End If
End Function

Private Function TotalQuantidadeVenda_MovimentoBomba(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroTanque As Integer) As Currency
    Dim xCondicao As String
    
    TotalQuantidadeVenda_MovimentoBomba = 0
    xCondicao = "Data = " & preparaData(pData)
'    If pNumeroTanque > 0 Then
'        xCondicao = xCondicao & " AND [Numero do Tanque] = " & pNumeroTanque
'    End If
    
    If rstMovimentoBomba.RecordCount > 0 Then
        rstMovimentoBomba.MoveFirst
        rstMovimentoBomba.Find xCondicao
        If Not rstMovimentoBomba.EOF Then
            Do Until rstMovimentoBomba.EOF
                If rstMovimentoBomba!Data = pData Then
                    If pNumeroTanque > 0 Then
                        If pNumeroTanque = 0 Or rstMovimentoBomba![Numero do Tanque] = pNumeroTanque Then
                            TotalQuantidadeVenda_MovimentoBomba = TotalQuantidadeVenda_MovimentoBomba + rstMovimentoBomba!qtd
                        End If
                    Else
                        TotalQuantidadeVenda_MovimentoBomba = TotalQuantidadeVenda_MovimentoBomba + rstMovimentoBomba!qtd
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoBomba.MoveNext
            Loop
        End If
    End If
End Function
Private Function TotalQuantidadeVendaNormal_MovimentoBomba(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroTanque As Integer) As Currency
    Dim xCondicao As String
    
    TotalQuantidadeVendaNormal_MovimentoBomba = 0
    xCondicao = "Data = " & preparaData(pData)
'    If pNumeroTanque > 0 Then
'        xCondicao = xCondicao & " AND [Numero do Tanque] = " & pNumeroTanque
'    End If
    
    If rstMovimentoBombaNormal.RecordCount > 0 Then
        rstMovimentoBombaNormal.MoveFirst
        rstMovimentoBombaNormal.Find xCondicao
        If Not rstMovimentoBombaNormal.EOF Then
            Do Until rstMovimentoBombaNormal.EOF
                If rstMovimentoBombaNormal!Data = pData Then
                    If pNumeroTanque > 0 Then
                        If pNumeroTanque = 0 Or rstMovimentoBombaNormal![Numero do Tanque] = pNumeroTanque Then
                            TotalQuantidadeVendaNormal_MovimentoBomba = TotalQuantidadeVendaNormal_MovimentoBomba + rstMovimentoBombaNormal!qtd
                        End If
                    Else
                        TotalQuantidadeVendaNormal_MovimentoBomba = TotalQuantidadeVendaNormal_MovimentoBomba + rstMovimentoBombaNormal!qtd
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoBombaNormal.MoveNext
            Loop
        End If
    End If
End Function

Private Function TotalValorVenda_MovimentoBomba(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroTanque As Integer) As Currency
    Dim xCondicao As String
    
    TotalValorVenda_MovimentoBomba = 0
    xCondicao = "Data = " & preparaData(pData)
'    If pNumeroTanque > 0 Then
'        xCondicao = xCondicao & " AND [Numero do Tanque] = " & pNumeroTanque
'    End If
'    lSQL = lSQL & "          , Sum([Quantidade da Saida] * [Preco de Venda]) AS ValorVenda" 'new 31/07
    
    If rstMovimentoBomba.RecordCount > 0 Then
        rstMovimentoBomba.MoveFirst
        rstMovimentoBomba.Find xCondicao
        If Not rstMovimentoBomba.EOF Then
            Do Until rstMovimentoBomba.EOF
                If rstMovimentoBomba!Data = pData Then
                    If pTipoCombustivel = "" Or rstMovimentoBomba![Tipo de Combustivel] = pTipoCombustivel Then
                        TotalValorVenda_MovimentoBomba = TotalValorVenda_MovimentoBomba + rstMovimentoBomba!ValorVenda
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoBomba.MoveNext
            Loop
        End If
    End If
End Function

Private Function TotalCustoMedioVenda_MovimentoBomba(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pNumeroTanque As Integer) As Currency
    Dim xCondicao As String
    
    TotalCustoMedioVenda_MovimentoBomba = 0
    xCondicao = "Data = " & preparaData(pData)
    If rstMovimentoBomba.RecordCount > 0 Then
        rstMovimentoBomba.MoveFirst
        rstMovimentoBomba.Find xCondicao
        If Not rstMovimentoBomba.EOF Then
            Do Until rstMovimentoBomba.EOF
                If rstMovimentoBomba!Data = pData Then
                    If pTipoCombustivel = "" Or rstMovimentoBomba![Tipo de Combustivel] = pTipoCombustivel Then
                        TotalCustoMedioVenda_MovimentoBomba = TotalCustoMedioVenda_MovimentoBomba + rstMovimentoBomba!TotalCustoVenda
                    End If
                Else
                    Exit Do
                End If
                rstMovimentoBomba.MoveNext
            Loop
        End If
    End If
End Function

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Combustivel = Nothing
    Set EntradaCombustivel = Nothing
    Set Fornecedor = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoBombaNormal = Nothing
    Set rstMovimentoAfericao = Nothing
    Set rstMovimentoNFeDevolucao = Nothing
    Set rstEntradaCombustivel = Nothing
    Set rstMedicaoCombustivel = Nothing
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    If lPagina = 0 Then
        lPagina = Val(txtPaginaInicial.Text) - 1
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
    xLinha = "|                                                                                                                            FOLHA: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| RESUMO DO L.M.C.                                                                                                  Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 3, 40) = "RESUMO DO L.M.C.                       "
    Else
        Mid(xLinha, 3, 40) = "RESUMO DA MOVIMENTAÇÃO DE COMBUSTÍVEL  "
    End If
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____                                                                                                 |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| COMBUSTIVEL.:                                                                                                                         |"
    Mid(xLinha, 17, 30) = Mid(cbo_combustivel.Text, 6, Len(cbo_combustivel.Text))
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|   DATA   |ESTOQUE  DE|  TOTAL  | TOTAL DAS |TOTAL DAS|  ESTOQUE |  ESTOQUE | -  PERDAS |   CUSTO  |VALOR TOTAL|   VENDA   |VALOR TOTAL|"
    If chkCustoDoEstoque.Value = 1 Then
        Mid(xLinha, 102, 11) = "TOTAL CUSTO"
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|          |  ABERTURA | RECEBIDO|   VENDAS  |AFERICOES|ESCRITURAL|FECHAMENTO| +  SOBRAS |  UNITARIO| DO  CUSTO |  UNITARIO | DA  VENDA |"
    If chkCustoDoEstoque.Value = 1 Then
        Mid(xLinha, 102, 11) = "DO  ESTOQUE"
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpCabEntrada()
    Dim xLinha As String
    
    If lPagina = 0 Then
        lPagina = Val(txtPaginaInicial.Text) - 1
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
    If chkImprimeResumo.Value = 0 Then
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                   FOLHA: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| ENTRADA DE COMBUSTIVEL                                   Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 26, 10) = "- L.M.C.  "
    Else
        Mid(xLinha, 26, 10) = "          "
    End If
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| COMBUSTIVEL.:                                                                |"
    Mid(xLinha, 17, 30) = Mid(cbo_combustivel.Text, 6, Len(cbo_combustivel.Text))
    BioImprime "@Printer.Print " & xLinha
    'BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
'             12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "+----------+----------+-----------+-----------+-----------+--------------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DATA  DA |  NUMERO  |   VALOR   | QUANTIDADE|   VALOR   |                    |"
    If lTodosCombustiveis = True Then
        Mid(xLinha, 60, 20) = "COMBUSTIVEL         "
    Else
        Mid(xLinha, 60, 20) = "FORNECEDOR          "
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  ENTRADA | DA  NOTA |  UNITARIO | DA ENTRADA|   TOTAL   |                    |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+----------+-----------+-----------+-----------+--------------------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDadosEntrada()
    Dim xLinha As String
    
    Do Until rstEntradaCombustivel.EOF
        If Trim(txtFornecedor.Text) <> "" Then
            If Fornecedor.LocalizarCodigo(g_empresa, rstEntradaCombustivel("Codigo do Fornecedor").Value) Then
                If UCase(Fornecedor.Nome) Like "*" & UCase(Trim(txtFornecedor.Text)) & "*" Then
                    ImpDetEntrada
                    lParcialEntrada = lParcialEntrada + rstEntradaCombustivel("Quantidade").Value
                    lParcialVenda = lParcialVenda + rstEntradaCombustivel("Valor da Entrada").Value
                End If
            End If
        Else
            ImpDetEntrada
            lParcialEntrada = lParcialEntrada + rstEntradaCombustivel("Quantidade").Value
            lParcialVenda = lParcialVenda + rstEntradaCombustivel("Valor da Entrada").Value
        End If
        rstEntradaCombustivel.MoveNext
    Loop
    ImpTotalEntrada
    
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Entrada de Combustíveis|@|"
    frm_preview.Show 1
End Sub
Private Sub ImprimeConsistenciaNFeDestinadas()
    Dim xLinha As String
    Dim i As Integer
    Dim xChaveAcesso
       
    If lChaveAcessoNFeDestinadasNaoCadastradas.Count = 0 Then
        Exit Sub
    End If
    
    If lLinha >= 95 Then
        xLinha = "+---------------------------------+-----------+-----------+--------------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCabEntrada
    End If
        '              1         2         3         4         5         6         7         8
        '     12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "                                                                                "
    Mid(xLinha, 2, 52) = "**NFe Destinadas não cadastradas/Importadas no SGP**"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    For Each xChaveAcesso In lChaveAcessoNFeDestinadasNaoCadastradas
        Dim xData As String
        xData = "Não Informada"
        If IsDate(lChaveAcessoNFeDestinadasNaoCadastradas(xChaveAcesso)) Then
            xData = FormatDateTime(CDate(lChaveAcessoNFeDestinadasNaoCadastradas(xChaveAcesso)), vbShortDate)
        End If
        Mid(xLinha, 2, 80) = "Chave: " & RetiraString(1, xChaveAcesso) & " Data: " & xData & " - " & RetiraString(3, xChaveAcesso)
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    Next
    
End Sub


Private Sub ImprimeDevolucaoNFe(ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pAbertura As Currency, ByVal pFechamento As Currency)
    Dim xLinha As String
    Dim i As Integer
    Dim xTotalDevolucao As Currency
    Dim xEstoqueFechamento As Currency
    Dim xPerdasSobras As Currency
    
    ' Subtrai as NFe de Devoluções 5661, 6661
    If rstMovimentoNFeDevolucao.RecordCount > 0 Then
        rstMovimentoNFeDevolucao.MoveFirst
        rstMovimentoNFeDevolucao.Find "Data = " & preparaData(pData)
        pFechamento = pAbertura
        If Not rstMovimentoNFeDevolucao.EOF Then
            Do Until rstMovimentoNFeDevolucao.EOF
                If rstMovimentoNFeDevolucao!Data = pData Then
                    xTotalDevolucao = xTotalDevolucao - rstMovimentoNFeDevolucao!Quantidade
                    xLinha = "|          |           |         |           |         |          |          |           |          |  *** DEVOLUÇÃO DE COMBUSTÍVEL *** |"
                    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
                    
                    i = Len(Format(pAbertura, "####,##0.0"))
                    Mid(xLinha, 14 + 10 - i, i) = Format(pAbertura, "####,##0.0")
                    
                    i = Len(Format(xTotalDevolucao, "#,###,##0"))
                    Mid(xLinha, 25 + 9 - i, i) = Format(xTotalDevolucao, "#,###,##0")
                    
                    
                    Mid(xLinha, 36, 9) = "  N. NFe "
                    i = Len(Format(rstMovimentoNFeDevolucao!numero, "###,##0"))
                    Mid(xLinha, 48 + 7 - i, i) = Format(rstMovimentoNFeDevolucao!numero, "###,##0")
                    
                    
                    xEstoqueFechamento = pAbertura - rstMovimentoNFeDevolucao!Quantidade
                    
                    i = Len(Format(xEstoqueFechamento, "###,##0.00"))
                    Mid(xLinha, 57 + 10 - i, i) = Format(xEstoqueFechamento, "###,##0.00")
                    
                    i = Len(Format(pFechamento, "###,##0.00"))
                    Mid(xLinha, 68 + 10 - i, i) = Format(pFechamento, "###,##0.00")
                    
                    
                    xPerdasSobras = pFechamento - pAbertura + rstMovimentoNFeDevolucao!Quantidade
                    i = Len(Format(xPerdasSobras, "#,###,##0.00"))
                    Mid(xLinha, 80 + 10 - i, i) = Format(xPerdasSobras, "#,###,##0.00")
                    
                    l_perdas_sobras = l_perdas_sobras + xPerdasSobras
                    
                    
                    BioImprime "@Printer.Print " & xLinha
                    lLinha = lLinha + 1
                Else
                    Exit Do
                End If
                rstMovimentoNFeDevolucao.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub ImpDet(ByVal x_data As Date)
    Dim xLinha As String
    Dim i As Integer
    Dim x_estoque_abertura As Currency
    Dim x_total_recebido As Currency
    Dim x_total_vendas As Currency
    Dim x_total_valor_vendas As Currency
    Dim x_total_afericao As Currency
    Dim x_estoque_escritural As Currency
    Dim x_estoque_fechamento As Currency
    Dim x_perdas_sobras As Currency
    Dim xVendaUnitario As Currency
    Dim xLitrosVendaNormal As Currency
    
    'x_estoque_abertura = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, x_data, l_tipo_combustivel, 0)
    x_estoque_abertura = TotalMedidaCombustivel(x_data, l_tipo_combustivel)
    'x_total_recebido = EntradaCombustivel.TotalEntradaPeriodo(g_empresa, x_data, x_data, l_tipo_combustivel, 0)
    x_total_recebido = TotalEntradaPeriodo(x_data, l_tipo_combustivel)
    'x_total_vendas = MovimentoBomba.QuantidadeVendaData(g_empresa, x_data, x_data, l_tipo_combustivel, 0)
    x_total_vendas = TotalQuantidadeVenda_MovimentoBomba(x_data, l_tipo_combustivel, 0)
    
    'x_total_valor_vendas = MovimentoBomba.ValorVendaPeriodo(g_empresa, x_data, x_data, l_tipo_combustivel, 1, 9)
    x_total_valor_vendas = TotalValorVenda_MovimentoBomba(x_data, l_tipo_combustivel, 0)
    'lSubCustoMedio = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, x_data, x_data, l_tipo_combustivel, 1, 9, "")
    lSubCustoMedio = TotalCustoMedioVenda_MovimentoBomba(x_data, l_tipo_combustivel, 0)
    'lSubCustoMedio = lSubCustoMedio - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, x_data, x_data, 1, 9, l_tipo_combustivel, "")
    lSubCustoMedio = lSubCustoMedio - TotalCustoCombustivel_MovimentoAfericao(x_data, l_tipo_combustivel, 0)
    
    'diminui nas vendas do mês os descontos do mês
    If g_nome_usuario = "L.M.C." Then
        'x_total_valor_vendas = x_total_valor_vendas - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, x_data + 1, x_data + 1, Mid(cbo_combustivel.Text, 1, 2))
        x_total_valor_vendas = x_total_valor_vendas - TotalDescontoCombustivel_MedicaoCombustivel(x_data + 1, l_tipo_combustivel)
    End If
    'x_total_afericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, x_data, x_data, 1, 9, l_tipo_combustivel, "")
    x_total_afericao = TotalQtdPeriodoCombustivel(x_data, l_tipo_combustivel, 0)
    'x_total_valor_vendas = x_total_valor_vendas - MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, x_data, x_data, 1, 9, l_tipo_combustivel, "")
    x_total_valor_vendas = x_total_valor_vendas - ValorTotalCombustivel_MovimentoAfericao(x_data, l_tipo_combustivel, 0)
    x_estoque_escritural = x_estoque_abertura + x_total_recebido + x_total_afericao - x_total_vendas
    'x_estoque_fechamento = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(x_data + 1), l_tipo_combustivel, 0)
    x_estoque_fechamento = TotalMedidaCombustivel(x_data + 1, l_tipo_combustivel)
    x_perdas_sobras = x_estoque_fechamento - x_estoque_escritural
    l_total_recebido = l_total_recebido + x_total_recebido
    l_total_vendas = l_total_vendas + x_total_vendas
    l_perdas_sobras = l_perdas_sobras + x_perdas_sobras
    l_total_valor_vendas = l_total_valor_vendas + Format(x_total_valor_vendas, "0000000000.00")
    l_total_afericao = l_total_afericao + x_total_afericao
    lTotalCusto = lTotalCusto + lSubCustoMedio
    If (x_total_vendas - x_total_afericao) > 0 Then
        xVendaUnitario = Format(x_total_valor_vendas / (x_total_vendas - x_total_afericao), "00000000.0000")
        lCustoMedio = Format(lSubCustoMedio / (x_total_vendas - x_total_afericao), "00000000.0000")
    Else
        xVendaUnitario = 0
        lCustoMedio = 0
    End If

    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 100 Then
        xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
        Mid(xLinha, 102, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    If lParcialMes = 0 Then
        lParcialMes = Month(x_data)
    End If
    If lParcialMes <> Month(x_data) Then
        ImpResumoMes
        lParcialEntrada = 0
        lParcialVenda = 0
        lParcialAfericao = 0
        lParcialPerdasSobras = 0
        lParcialMes = Month(x_data)
    End If
    xLinha = "|          |           |         |           |         |          |          |           |          |           |           |           |"
    Mid(xLinha, 2, 10) = Format(x_data, "dd/mm/yyyy")
    i = Len(Format(x_estoque_abertura, "####,##0.0"))
    Mid(xLinha, 14 + 10 - i, i) = Format(x_estoque_abertura, "####,##0.0")
    i = Len(Format(x_total_recebido, "#,###,##0"))
    Mid(xLinha, 25 + 9 - i, i) = Format(x_total_recebido, "#,###,##0")
    i = Len(Format(x_total_vendas, "####,##0.00"))
    Mid(xLinha, 35 + 11 - i, i) = Format(x_total_vendas, "####,##0.00")
    If g_nome_usuario = "L.M.C." Then
        'xLitrosVendaNormal = MovimentoBombaNormal.QuantidadeVendaData(g_empresa, x_data, x_data, l_tipo_combustivel, 0)
        xLitrosVendaNormal = TotalQuantidadeVendaNormal_MovimentoBomba(x_data, l_tipo_combustivel, 0)
        If xLitrosVendaNormal < x_total_vendas Then
            Mid(xLinha, 47, 1) = "+"
        ElseIf xLitrosVendaNormal > x_total_vendas Then
            Mid(xLinha, 47, 1) = "-"
        End If
    End If
    i = Len(Format(x_total_afericao, "##,##0.00"))
    Mid(xLinha, 47 + 9 - i, i) = Format(x_total_afericao, "##,##0.00")
    i = Len(Format(x_estoque_escritural, "###,##0.00"))
    Mid(xLinha, 57 + 10 - i, i) = Format(x_estoque_escritural, "###,##0.00")
    i = Len(Format(x_estoque_fechamento, "###,##0.00"))
    Mid(xLinha, 68 + 10 - i, i) = Format(x_estoque_fechamento, "###,##0.00")
    i = Len(Format(x_perdas_sobras, "#,###,##0.00"))
    Mid(xLinha, 80 + 10 - i, i) = Format(x_perdas_sobras, "#,###,##0.00")
    i = Len(Format(lCustoMedio, "####0.0000"))
    Mid(xLinha, 91 + 10 - i, i) = Format(lCustoMedio, "####0.0000")
    If chkCustoDoEstoque.Value = 1 Then
        lSubCustoMedio = Format(lCustoMedio * x_estoque_fechamento, "0000000000.00")
    End If
    i = Len(Format(lSubCustoMedio, "####,##0.00"))
    Mid(xLinha, 102 + 11 - i, i) = Format(lSubCustoMedio, "####,##0.00")
    i = Len(Format(xVendaUnitario, "##,##0.0000"))
    Mid(xLinha, 114 + 11 - i, i) = Format(xVendaUnitario, "##,##0.0000")
    i = Len(Format(x_total_valor_vendas, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(x_total_valor_vendas, "####,##0.00")
    
    Call ImprimeDevolucaoNFe(x_data, l_tipo_combustivel, x_estoque_abertura, x_estoque_fechamento)
    
    If chkRescunhoLmc.Value = 0 Then
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    lParcialEntrada = lParcialEntrada + x_total_recebido
    lParcialVenda = lParcialVenda + x_total_vendas
    lParcialAfericao = lParcialAfericao + x_total_afericao
    lParcialPerdasSobras = lParcialPerdasSobras + x_perdas_sobras

    If x_estoque_escritural < 0 Then
        AdcionaMensagem (Format(x_data, "dd/mm/yyyy") & ": ESTOQUE ESCRITURAL NEGATIVO " & x_estoque_escritural)
    ElseIf x_estoque_fechamento < 0 Then
        AdcionaMensagem (Format(x_data, "dd/mm/yyyy") & ": ESTOQUE DE FECHAMENTO NEGATIVO " & x_estoque_fechamento)
    End If
    If x_perdas_sobras < -300 Or x_perdas_sobras > 300 Then
        AdcionaMensagem (Format(x_data, "dd/mm/yyyy") & ": PERDAS/SOBRAS ESTÁ ACIMA DE UMA NORMALIDADE PADRÃO. " & x_perdas_sobras)
    End If
    
End Sub
Private Sub ImpDetBico(ByVal pData As Date)
    Dim xLinha As String
    Dim xTotal As Currency
    Dim i As Integer
   
    xTotal = 0
    
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Codigo da Bomba], [Quantidade da Saida], Periodo"
'    lSQL = lSQL & "  FROM " & MovimentoBomba.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(pData)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    lSQL = lSQL & " ORDER BY [Codigo da Bomba] ASC, Periodo ASC"
'
'    'cbo_combustivel.Clear
'    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    'loop RecordSet
    With rstMovimentoBomba
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "Data = " & preparaData(pData)
        End If
        If Not .BOF Or Not .EOF Then
            '.MoveFirst
            Do Until .EOF
                If !Data <> pData Then
                    Exit Do
                End If
                'xTotal = xTotal + ![Quantidade da Saida]
                xTotal = xTotal + !qtd
                If lLinha >= 100 Then
                    xLinha = "+------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+----------+"
                    Mid(xLinha, 102, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                xLinha = "|          | BICO:     |P-       |           |         |          |          |           |          |           |           |           |"
                i = Len(Format(![Codigo da Bomba], "#0"))
                Mid(xLinha, 19 + 2 - i, i) = Format(![Codigo da Bomba], "#0")
                Mid(xLinha, 27, 1) = !Periodo
                'i = Len(Format(![Quantidade da Saida], "#,###,##0.00"))
                'Mid(xLinha, 33 + 12 - i, i) = Format(![Quantidade da Saida], "#,###,##0.00")
                i = Len(Format(!qtd, "#,###,##0.00"))
                Mid(xLinha, 33 + 12 - i, i) = Format(!qtd, "#,###,##0.00")
                i = Len(Format(xTotal, "#,###,##0.00"))
                Mid(xLinha, 55 + 12 - i, i) = Format(xTotal, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
                .MoveNext
            Loop
        End If
        '.Close
    End With
    'Set rstMovimentoBomba = Nothing
End Sub
Private Sub ImpDetTanque(ByVal pData As Date)
    Dim xLinha As String
    Dim i As Integer
    Dim xEstoqueAbertura As Currency
    Dim xTotalRecebido As Currency
    Dim xTotalVendas As Currency
    Dim xTotalAfericao As Currency
    Dim xEstoqueEscritural As Currency
    Dim xEstoqueFechamento As Currency
    Dim xPerdasSobras As Currency
   
    lSQL = ""
    lSQL = lSQL & "SELECT [Numero do Tanque]"
    lSQL = lSQL & "  FROM Tanque_Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & " ORDER BY [Numero do Tanque] ASC"
    Set rstTanque = Conectar.RsConexao(lSQL)
    
    'loop RecordSet
    With rstTanque
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                xEstoqueAbertura = 0
                xTotalRecebido = 0
                xTotalVendas = 0
                xTotalAfericao = 0
                xEstoqueEscritural = 0
                xEstoqueFechamento = 0
                xPerdasSobras = 0
                xEstoqueAbertura = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, pData, l_tipo_combustivel, ![Numero do Tanque])
                xTotalRecebido = EntradaCombustivel.TotalEntradaPeriodo(g_empresa, pData, pData, l_tipo_combustivel, ![Numero do Tanque])
                'xTotalVendas = MovimentoBomba.QuantidadeVendaData(g_empresa, pData, pData, l_tipo_combustivel, ![Numero do Tanque])
                xTotalVendas = TotalQuantidadeVenda_MovimentoBomba(pData, l_tipo_combustivel, ![Numero do Tanque])
                xTotalAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, pData, pData, 1, 9, l_tipo_combustivel, "")
                xEstoqueEscritural = xEstoqueAbertura + xTotalRecebido + xTotalAfericao - xTotalVendas
                xEstoqueFechamento = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(pData + 1), l_tipo_combustivel, ![Numero do Tanque])
                xPerdasSobras = xEstoqueFechamento - xEstoqueEscritural
                'aquiaqui
                If lLinha >= 100 Then
                    xLinha = "+------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+----------+"
                    Mid(xLinha, 102, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                        ' 1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                xLinha = "|          |           |         |           |         |          |          |           | TQ:      |           |           |           |"
                Mid(xLinha, 6, 2) = Format(pData, "dd")
                i = Len(Format(xEstoqueAbertura, "####,##0.0"))
                Mid(xLinha, 14 + 10 - i, i) = Format(xEstoqueAbertura, "####,##0.0")
                i = Len(Format(xTotalRecebido, "#,###,##0"))
                Mid(xLinha, 25 + 9 - i, i) = Format(xTotalRecebido, "#,###,##0")
                i = Len(Format(xTotalVendas, "####,##0.00"))
                Mid(xLinha, 35 + 11 - i, i) = Format(xTotalVendas, "####,##0.00")
                i = Len(Format(xTotalAfericao, "##,##0.00"))
                Mid(xLinha, 47 + 9 - i, i) = Format(xTotalAfericao, "##,##0.00")
                i = Len(Format(xEstoqueEscritural, "###,##0.00"))
                Mid(xLinha, 57 + 10 - i, i) = Format(xEstoqueEscritural, "###,##0.00")
                i = Len(Format(xEstoqueFechamento, "###,##0.00"))
                Mid(xLinha, 68 + 10 - i, i) = Format(xEstoqueFechamento, "###,##0.00")
                i = Len(Format(xPerdasSobras, "#,###,##0.00"))
                Mid(xLinha, 80 + 10 - i, i) = Format(xPerdasSobras, "#,###,##0.00")
                i = Len(Format(![Numero do Tanque], "#0"))
                Mid(xLinha, 96 + 2 - i, i) = Format(![Numero do Tanque], "#0")
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
                .MoveNext
            Loop
        End If
        '.Close
    End With
    'Set rstMovimentoBomba = Nothing
End Sub
Private Sub ImpDetEntrada()
    Dim xLinha As String
    Dim i As Integer
    
    If lLinha >= 95 Then
        xLinha = "+---------------------------------+-----------+-----------+--------------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCabEntrada
    End If
        '              1         2         3         4         5         6         7         8
        '     12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|          |          |           |           |           |                    |"
    'tarefa redmine 567
    Mid(xLinha, 2, 10) = Format(rstEntradaCombustivel("Data").Value, "dd/mm/yyyy")
    'Mid(xLinha, 2, 10) = Format(rstEntradaCombustivel("Data").Value, "dd/mm/yyyy")
    Mid(xLinha, 13, 10) = rstEntradaCombustivel("Numero da Nota").Value
    i = Len(Format(rstEntradaCombustivel("Valor do Litro").Value, "##,##0.0000"))
    Mid(xLinha, 23 + 11 - i, i) = Format(rstEntradaCombustivel("Valor do Litro").Value, "##,##0.0000")
    i = Len(Format(rstEntradaCombustivel("Quantidade").Value, "####,##0.00"))
    Mid(xLinha, 36 + 11 - i, i) = Format(rstEntradaCombustivel("Quantidade").Value, "####,##0.00")
    i = Len(Format(rstEntradaCombustivel("Valor da Entrada").Value, "####,##0.00"))
    Mid(xLinha, 48 + 11 - i, i) = Format(rstEntradaCombustivel("Valor da Entrada").Value, "####,##0.00")
    If lTodosCombustiveis = True Then
        Mid(xLinha, 60, 2) = rstEntradaCombustivel("Tipo de Combustivel").Value
        If Fornecedor.LocalizarCodigo(g_empresa, rstEntradaCombustivel("Codigo do Fornecedor").Value) Then
            Mid(xLinha, 63, 17) = Fornecedor.Nome
        Else
            Mid(xLinha, 63, 17) = "** INEXISTENTE **"
        End If
    Else
        If Fornecedor.LocalizarCodigo(g_empresa, rstEntradaCombustivel("Codigo do Fornecedor").Value) Then
            Mid(xLinha, 60, 20) = Fornecedor.Nome
        Else
            Mid(xLinha, 60, 20) = "** INEXISTENTE **"
        End If
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpLinhaRascunhoLmc(ByVal pLinha As String)
    Dim xLinha As String
                
    If lLinha >= 100 Then
        xLinha = "+------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+----------+"
        Mid(xLinha, 102, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    BioImprime "@Printer.Print " & pLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpRascunhoLmc(ByVal pData As Date)
    Dim xLinha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xQtdVenda As Currency
    Dim xTotalQtdVenda As Currency
    Dim xTotalValorVenda As Currency
    Dim xTotalValorVendaMes As Currency
    Dim xTotalQtdAfericao As Currency
    Dim xTotalValorAfericao As Currency
    Dim xQtdAfericao As Currency
    Dim xEstoqueEscritural As Currency
    Dim xQtdEntrada As Currency
    Dim xVolumeDisponivel As Currency
    Dim xNumeroTanque As Integer
    Dim xString As String
    Dim xEstoqueFinal As Currency
    Dim xDataInicio As String
   
    xTotalQtdAfericao = 0
    xTotalValorAfericao = 0
    xTotalQtdVenda = 0
    xTotalValorVenda = 0
    xTotalValorVendaMes = 0
    xEstoqueEscritural = 0
    xQtdEntrada = 0
    xVolumeDisponivel = 0
    xEstoqueFinal = 0
    xNumeroTanque = 1
    'lNumeroTanque = 1
    
    
    If lPagina = 0 Then
        ImpCab
    End If
    
    xLinha = "|                                                                                                                                       |"
    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
    Mid(xLinha, 127, 10) = Format(pData, "dd/mm/yyyy")
    ImpLinhaRascunhoLmc (xLinha)
    
    
    'Início Estoque Abertura
    xLinha = "|     TQ-                   TQ-                   TQ-                   TQ-                    TQ-                EST.ABERT.            |"
    i2 = 0
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Numero do Tanque], Quantidade"
'    lSQL = lSQL & "  FROM " & MedicaoCombustivel.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(pData)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    lSQL = lSQL & " ORDER BY [Numero do Tanque]"
'    Set rstMedicaoCombustivel = Conectar.RsConexao(lSQL)
    If rstMedicaoCombustivel.RecordCount > 0 Then
        rstMedicaoCombustivel.MoveFirst
        rstMedicaoCombustivel.Find "Data = " & preparaData(pData)
        Do Until rstMedicaoCombustivel.EOF
            If rstMedicaoCombustivel!Data <> pData Then
                Exit Do
            End If
            i2 = i2 + 1
            xEstoqueEscritural = xEstoqueEscritural + rstMedicaoCombustivel!qtd
            xVolumeDisponivel = xVolumeDisponivel + rstMedicaoCombustivel!qtd
            xNumeroTanque = rstMedicaoCombustivel![Numero do Tanque]
            'lNumeroTanque = rstMedicaoCombustivel![Numero do Tanque]
            If i2 = 1 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xLinha, 10 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xLinha, 10 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xLinha, 15 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 2 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xLinha, 32 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xLinha, 32 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xLinha, 37 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 3 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xLinha, 54 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xLinha, 54 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xLinha, 58 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 4 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xLinha, 76 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xLinha, 76 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xLinha, 81 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 5 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xLinha, 99 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xLinha, 99 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xLinha, 104 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            End If
            rstMedicaoCombustivel.MoveNext
        Loop
    End If
    'rstMedicaoCombustivel.Clone
    i = Len(Format(xEstoqueEscritural, "###,##0.0"))
    Mid(xLinha, 128 + 9 - i, i) = Format(xEstoqueEscritural, "###,##0.0")
    ImpLinhaRascunhoLmc (xLinha)
    'Fim Estoque Abertura
    
    

    
    'Início Busca Entradas
    xQtdAfericao = 0
'    lSQL = ""
'    lSQL = lSQL & "SELECT Data, [Numero da Nota], Quantidade"
'    lSQL = lSQL & "  FROM " & EntradaCombustivel.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(pData)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    lSQL = lSQL & " ORDER BY [Numero da Nota]"
'    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
    If rstEntradaCombustivel.RecordCount > 0 Then
        rstEntradaCombustivel.MoveFirst
        rstEntradaCombustivel.Find "Data = " & preparaData(pData)
        Do Until rstEntradaCombustivel.EOF
            If rstEntradaCombustivel!Data <> pData Then
                Exit Do
            End If
            xQtdEntrada = xQtdEntrada + rstEntradaCombustivel!qtd
            xVolumeDisponivel = xVolumeDisponivel + rstEntradaCombustivel!qtd
            xEstoqueEscritural = xEstoqueEscritural + rstEntradaCombustivel!qtd
            xLinha = "|                                                N.NOTA                  DATA               N.TANQUE              VOL.RECEB.            |"
            Mid(xLinha, 57, 10) = rstEntradaCombustivel![Numero da Nota]
            Mid(xLinha, 80, 10) = Format(rstEntradaCombustivel!Data, "dd/mm/yyyy")
            i = Len(Format(xNumeroTanque, "#0"))
            'i = Len(Format(lNumeroTanque, "#0"))
            Mid(xLinha, 107 + 2 - i, i) = Format(xNumeroTanque, "#0")
            'Mid(xLinha, 107 + 2 - i, i) = Format(lNumeroTanque, "#0")
            i = Len(Format(rstEntradaCombustivel!qtd, "###,##0.0"))
            Mid(xLinha, 128 + 9 - i, i) = Format(rstEntradaCombustivel!qtd, "###,##0.0")
            ImpLinhaRascunhoLmc (xLinha)
            rstEntradaCombustivel.MoveNext
        Loop
    End If
    'rstEntradaCombustivel.Clone
    xLinha = "|                                                                                                         Volume Recebido               |"
    i = Len(Format(xQtdEntrada, "###,##0.0"))
    Mid(xLinha, 128 + 9 - i, i) = Format(xQtdEntrada, "###,##0.0")
    ImpLinhaRascunhoLmc (xLinha)
    xLinha = "|                                                                                                         Volume Disponivel             |"
    i = Len(Format(xVolumeDisponivel, "###,##0.0"))
    Mid(xLinha, 128 + 9 - i, i) = Format(xVolumeDisponivel, "###,##0.0")
    ImpLinhaRascunhoLmc (xLinha)
    'Fim Busca Entradas

    
    
    'aquiaquiauqiaqui29/07 rascunholmc29
    
    'Busca Movimento de Bomba
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Numero do Tanque], [Codigo da Bomba], Periodo, Abertura, Encerrante, [Quantidade da Saida], [Preco de Venda]"
'    lSQL = lSQL & "  FROM " & MovimentoBomba.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(pData)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    lSQL = lSQL & " ORDER BY [Codigo da Bomba] ASC, Periodo ASC"
'
'    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    'loop RecordSet
    'With rstMovimentoBomba
        'If Not rstMovimentoBomba.BOF Or Not rstMovimentoBomba.EOF Then
        If rstMovimentoBomba.RecordCount > 0 Then
            rstMovimentoBomba.MoveFirst
            rstMovimentoBomba.Find "Data = " & preparaData(pData)
            Do Until rstMovimentoBomba.EOF
                If rstMovimentoBomba!Data <> pData Then
                    Exit Do
                End If
                xQtdVenda = rstMovimentoBomba!qtd
                xTotalQtdVenda = xTotalQtdVenda + rstMovimentoBomba!qtd
                xTotalValorVenda = xTotalValorVenda + Round(rstMovimentoBomba!qtd * rstMovimentoBomba![Preco de Venda], 2)
                xEstoqueEscritural = xEstoqueEscritural - rstMovimentoBomba!qtd

                'Início Busca Afericao
'                xQtdAfericao = 0
'                lSQL = ""
'                lSQL = lSQL & "SELECT SUM(Quantidade) as TotalQuantidade, SUM([Valor Total]) as TotalValor"
'                lSQL = lSQL & "  FROM " & MovimentoAfericao.NomeTabela
'                lSQL = lSQL & " WHERE Empresa = " & g_empresa
'                lSQL = lSQL & "   AND Data = " & preparaData(pData)
'                lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'                lSQL = lSQL & "   AND [Codigo da Bomba] = " & rstMovimentoBomba![Codigo da Bomba]
'                Set rstMovimentoAfericao = Conectar.RsConexao(lSQL)
          
                'With rstMovimentoAfericao
                If rstMovimentoAfericao.RecordCount > 0 Then
                    rstMovimentoAfericao.MoveFirst
                    rstMovimentoAfericao.Find "Data = " & preparaData(pData)
                    Do Until rstMovimentoAfericao.EOF
                        If rstMovimentoAfericao!Data <> pData Then
                            Exit Do
                        End If
                        If rstMovimentoAfericao![Codigo da Bomba] = rstMovimentoBomba![Codigo da Bomba] Then
                            xQtdAfericao = rstMovimentoAfericao!qtd
                            xTotalQtdAfericao = xTotalQtdAfericao + rstMovimentoAfericao!qtd
                            xTotalValorAfericao = xTotalValorAfericao + rstMovimentoAfericao!ValorTotal
                            xQtdVenda = xQtdVenda - rstMovimentoAfericao!qtd
                            xTotalQtdVenda = xTotalQtdVenda - rstMovimentoAfericao!qtd
                            xTotalValorVenda = xTotalValorVenda - rstMovimentoAfericao!ValorTotal
                            xEstoqueEscritural = xEstoqueEscritural + rstMovimentoAfericao!qtd
                        End If
                        rstMovimentoAfericao.MoveNext
                    Loop
                End If
                'rstMovimentoAfericao.Clone
                'End With
                'Fim Busca Afericao

'                                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
'                         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
                xLinha = "|                       TANQUE:       BICO:     FECHAM.                ABERT.               AFERICAO              VENDA BICO            |"
                i = Len(Format(rstMovimentoBomba![Numero do Tanque], "#0"))
                Mid(xLinha, 32 + 2 - i, i) = Format(rstMovimentoBomba![Numero do Tanque], "#0")
                i = Len(Format(rstMovimentoBomba![Codigo da Bomba], "#0"))
                Mid(xLinha, 44 + 2 - i, i) = Format(rstMovimentoBomba![Codigo da Bomba], "#0")
                i = Len(Format(rstMovimentoBomba!Encerrante, "####,##0.00"))
                Mid(xLinha, 56 + 11 - i, i) = Format(rstMovimentoBomba!Encerrante, "####,##0.00")
                i = Len(Format(rstMovimentoBomba!Abertura, "####,##0.00"))
                Mid(xLinha, 79 + 11 - i, i) = Format(rstMovimentoBomba!Abertura, "####,##0.00")
                i = Len(Format(xQtdAfericao, "####,##0.00"))
                Mid(xLinha, 103 + 10 - i, i) = Format(xQtdAfericao, "####,##0.00")
                i = Len(Format(xQtdVenda, "####,##0.00"))
                Mid(xLinha, 126 + 11 - i, i) = Format(xQtdVenda, "####,##0.00")
                ImpLinhaRascunhoLmc (xLinha)
                rstMovimentoBomba.MoveNext
            Loop
        End If
        'rstMovimentoBomba.Close
        'rstMovimentoAfericao.Close
    'End With
    'Set rstMovimentoBomba = Nothing
    
'                      1         2         3         4         5         6         7         8         9        10        11        12        13     13
'             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|                                                                                                         Vendas no dia                 |"
    i = Len(Format(xTotalQtdVenda, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(xTotalQtdVenda, "####,##0.00")
    ImpLinhaRascunhoLmc (xLinha)
    
'                      1         2         3         4         5         6         7         8         9        10        11        12        13     13
'             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|                                                              Preco das vendas do dia                    Estoque Escritural            |"
    xTotalValorVenda = xTotalValorVenda - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, (CDate(pData) + 1), CDate(pData + 1), l_tipo_combustivel)
    i = Len(Format(xTotalValorVenda, "####,##0.00"))
    Mid(xLinha, 90 + 11 - i, i) = Format(xTotalValorVenda, "####,##0.00")
    i = Len(Format(xEstoqueEscritural, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(xEstoqueEscritural, "####,##0.00")
    ImpLinhaRascunhoLmc (xLinha)
    
    
    'Início Estoque Fechamento
    xString = "|     TQ-                   TQ-                   TQ-                   TQ-                    TQ-                EST. TOTAL            |"
    i2 = 0
'    lSQL = ""
'    lSQL = lSQL & "SELECT [Numero do Tanque], Quantidade"
'    lSQL = lSQL & "  FROM " & MedicaoCombustivel.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(CDate(pData) + 1)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    lSQL = lSQL & " ORDER BY [Numero do Tanque]"
'    Set rstMedicaoCombustivel = Conectar.RsConexao(lSQL)
    If rstMedicaoCombustivel.RecordCount > 0 Then
        rstMedicaoCombustivel.MoveFirst
        rstMedicaoCombustivel.Find "Data = " & preparaData(pData + 1)
        Do Until rstMedicaoCombustivel.EOF
            If rstMedicaoCombustivel!Data <> (pData + 1) Then
                Exit Do
            End If
            i2 = i2 + 1
            xEstoqueFinal = xEstoqueFinal + rstMedicaoCombustivel!qtd
            xNumeroTanque = rstMedicaoCombustivel![Numero do Tanque]
            'lNumeroTanque = rstMedicaoCombustivel![Numero do Tanque]
            If i2 = 1 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xString, 10 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xString, 10 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xString, 15 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 2 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xString, 32 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xString, 32 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xString, 37 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 3 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xString, 54 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xString, 54 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xString, 58 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 4 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xString, 76 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xString, 76 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xString, 81 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            ElseIf i2 = 5 Then
                i = Len(Format(xNumeroTanque, "#0"))
                'i = Len(Format(lNumeroTanque, "#0"))
                Mid(xString, 99 + 2 - i, i) = Format(xNumeroTanque, "#0")
                'Mid(xString, 99 + 2 - i, i) = Format(lNumeroTanque, "#0")
                i = Len(Format(rstMedicaoCombustivel!qtd, "###,##0.0"))
                Mid(xString, 104 + 9 - i, i) = Format(rstMedicaoCombustivel!qtd, "###,##0.0")
            End If
            rstMedicaoCombustivel.MoveNext
        Loop
    End If
    'rstMedicaoCombustivel.Clone
    i = Len(Format(xEstoqueFinal, "###,##0.0"))
    Mid(xString, 128 + 9 - i, i) = Format(xEstoqueFinal, "###,##0.0")
    'Fim Estoque Fechamento
    
    
    
    'Calcula Aferições do Mês
    xDataInicio = Format(pData, "dd/mm/yyyy")
    Mid(xDataInicio, 1, 2) = "01"
'    lSQL = ""
'    lSQL = lSQL & "SELECT SUM([Valor Total]) as TotalValor"
'    lSQL = lSQL & "  FROM " & MovimentoAfericao.NomeTabela
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(xDataInicio))
'    lSQL = lSQL & "   AND Data <= " & preparaData(pData)
'    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
'    Set rstMovimentoAfericao = Conectar.RsConexao(lSQL)
    If rstMovimentoAfericao.RecordCount > 0 Then  '
    'If Not IsNull(rstMovimentoAfericao!TotalValor) Then
        rstMovimentoAfericao.MoveFirst '
            rstMovimentoAfericao.Find "Data = " & preparaData(pData) '
            Do Until rstMovimentoAfericao.EOF '
                If rstMovimentoAfericao!Data <> pData Then '
                    Exit Do '
                End If '
                If rstMovimentoAfericao![Codigo da Bomba] = rstMovimentoBomba![Codigo da Bomba] Then '
                    xTotalValorVendaMes = xTotalValorVendaMes - rstMovimentoAfericao!ValorTotal
                End If '
                rstMovimentoAfericao.MoveNext '
            Loop '
    'End If
    End If
    'rstMovimentoAfericao.Close
    'xTotalValorVendaMes = xTotalValorVendaMes + MovimentoBomba.ValorVendaPeriodo(g_empresa, CDate(xDataInicio), pData, l_tipo_combustivel, 1, 9)
    xTotalValorVendaMes = xTotalValorVendaMes + TotalValorVenda_MovimentoBomba(pData, l_tipo_combustivel, 0)
    xTotalValorVendaMes = xTotalValorVendaMes - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, (CDate(xDataInicio) + 1), CDate(pData + 1), l_tipo_combustivel)
    xLinha = "|                                                              Valor Acumulado do mes                     Estoque Fechamento            |"
    i = Len(Format(xTotalValorVendaMes, "####,##0.00"))
    Mid(xLinha, 90 + 11 - i, i) = Format(xTotalValorVendaMes, "####,##0.00")
    i = Len(Format(xEstoqueFinal, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(xEstoqueFinal, "####,##0.00")
    ImpLinhaRascunhoLmc (xLinha)
    
    xLinha = "|                                                                                                         -Perdas +Sobras(*)            |"
    xEstoqueEscritural = xEstoqueFinal - xEstoqueEscritural
    i = Len(Format(xEstoqueEscritural, "####,##0.00;(####,##0.00)"))
    Mid(xLinha, 126 + 11 - i, i) = Format(xEstoqueEscritural, "####,##0.00;(####,##0.00)")
    ImpLinhaRascunhoLmc (xLinha)
    
    
    ImpLinhaRascunhoLmc (xString)
    xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
    ImpLinhaRascunhoLmc (xLinha)
End Sub
Private Sub ImpResumoMes()
    Dim xLinha As String
    Dim i As Integer
    
    
    If Month(CDate(msk_data_i.Text)) <> Month(CDate(msk_data_f.Text)) Or Year(CDate(msk_data_i.Text)) <> Year(CDate(msk_data_f.Text)) Then
        xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|       PARCIAL DO MES |         |           |         |          |          |           |          |           |           |           |"
        i = Len(Format(lParcialEntrada, "#,###,##0"))
        Mid(xLinha, 25 + 9 - i, i) = Format(lParcialEntrada, "#,###,##0")
        i = Len(Format(lParcialVenda, "#,###,##0.00"))
        Mid(xLinha, 34 + 12 - i, i) = Format(lParcialVenda, "#,###,##0.00")
        i = Len(Format(lParcialAfericao, "#,###,##0.00"))
        Mid(xLinha, 44 + 12 - i, i) = Format(lParcialAfericao, "#,###,##0.00")
        i = Len(Format(lParcialPerdasSobras, "#,###,##0.00"))
        Mid(xLinha, 78 + 12 - i, i) = Format(lParcialPerdasSobras, "#,###,##0.00")
    '    i = Len(Format(lTotalCusto, "####,##0.00"))
    '    Mid(xLinha, 102 + 11 - i, i) = Format(lTotalCusto, "####,##0.00")
    '    i = Len(Format(l_total_valor_vendas, "#,###,##0.00"))
    '    Mid(xLinha, 125 + 12 - i, i) = Format(l_total_valor_vendas, "#,###,##0.00")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+----------------------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 3
    End If
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    Dim xPrecoMedio As Currency
    
    xLinha = "+----------+-----------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|            *** TOTAL |         |           |         |          |          |           |          |           |           |           |"
    i = Len(Format(l_total_recebido, "#,###,##0"))
    Mid(xLinha, 25 + 9 - i, i) = Format(l_total_recebido, "#,###,##0")
    i = Len(Format(l_total_vendas, "#,###,##0.00"))
    Mid(xLinha, 34 + 12 - i, i) = Format(l_total_vendas, "#,###,##0.00")
    i = Len(Format(l_total_afericao, "#,###,##0.00"))
    Mid(xLinha, 44 + 12 - i, i) = Format(l_total_afericao, "#,###,##0.00")
    i = Len(Format(l_perdas_sobras, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(l_perdas_sobras, "#,###,##0.00")
    xPrecoMedio = 0
    If l_total_vendas > 0 And lTotalCusto > 0 Then
        'xPrecoMedio = lTotalCusto / (l_total_vendas + l_total_afericao)
        xPrecoMedio = lTotalCusto / (l_total_vendas - l_total_afericao) 'foi alterado para subtração pois em casos em q havia aferição o preço de custo ficava incorreto
        i = Len(Format(xPrecoMedio, "####0.0000"))
        Mid(xLinha, 91 + 10 - i, i) = Format(xPrecoMedio, "####0.0000")
    End If
    i = Len(Format(lTotalCusto, "####,##0.00"))
    Mid(xLinha, 102 + 11 - i, i) = Format(lTotalCusto, "####,##0.00")
    xPrecoMedio = 0
    If l_total_vendas > 0 And l_total_valor_vendas > 0 Then
        'xPrecoMedio = l_total_valor_vendas / (l_total_vendas + l_total_afericao)
        xPrecoMedio = l_total_valor_vendas / (l_total_vendas - l_total_afericao) 'foi alterado para subtração pois em casos em q havia aferição o preço unitário ficava incorreto
        i = Len(Format(xPrecoMedio, "####0.0000"))
        Mid(xLinha, 114 + 11 - i, i) = Format(xPrecoMedio, "##,##0.0000")
    End If
    i = Len(Format(l_total_valor_vendas, "#,###,##0.00"))
    Mid(xLinha, 125 + 12 - i, i) = Format(l_total_valor_vendas, "#,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                      |         |           |         |          |          |           |          |           |LUCRO BRUTO|           |"
    l_total_valor_vendas = l_total_valor_vendas - lTotalCusto
    i = Len(Format(l_total_valor_vendas, "#,###,##0.00"))
    Mid(xLinha, 125 + 12 - i, i) = Format(l_total_valor_vendas, "#,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------------------+---------+-----------+---------+----------+----------+-----------+----------+-----------+-----------+-----------+"
    Mid(xLinha, 102, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    If l_perdas_sobras < -300 Or l_perdas_sobras > 300 Then
        AdcionaMensagem ("O TOTAL DAS PERDAS/SOBRAS ESTÁ ACIMA DE UMA NORMALIDADE PADRÃO.")
    End If
    ImprimeMensagem
'    l_total_valor_vendas = (l_total_valor_vendas - lTotalCusto) / l_total_vendas
'    xLinha = "  LUCRO MEDIO POR LITRO                                                                                                                  "
'    i = Len(Format(l_total_valor_vendas, "#,###,##0.00"))
'    Mid(xLinha, 40 + 12 - i, i) = Format(l_total_valor_vendas, "#,###,##0.00")
'    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpTotalEntrada()
    Dim xLinha As String
    Dim i As Integer
    Dim xPrecoMedio As Currency
    
    xLinha = "+----------+----------+-----------+-----------+-----------+--------------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                       *** TOTAL |           |           |                    |"
    i = Len(Format(lParcialEntrada, "####,##0.00"))
    Mid(xLinha, 36 + 11 - i, i) = Format(lParcialEntrada, "####,##0.00")
    i = Len(Format(lParcialVenda, "####,##0.00"))
    Mid(xLinha, 48 + 11 - i, i) = Format(lParcialVenda, "####,##0.00")
    xPrecoMedio = 0
    If lParcialEntrada > 0 And lParcialVenda > 0 Then
        xPrecoMedio = lParcialVenda / lParcialEntrada
        i = Len(Format(xPrecoMedio, "##,##0.0000"))
        Mid(xLinha, 68 + 11 - i, i) = Format(xPrecoMedio, "##,##0.0000")
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+---------------------------------+-----------+-----------+--------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImprimeMensagem()
    Dim xLinha As String
    Dim i As Integer

    For i = 0 To iObs - 1
        xLinha = lObservacao(i)
        BioImprime "@Printer.Print " & xLinha
    Next
End Sub
Private Sub AdcionaMensagem(ByVal pMensagem As String)
    If iObs < 30 Then
        lObservacao(iObs) = pMensagem
        iObs = iObs + 1
    ElseIf iObs = 30 Then
        If lObservacao(30) = "" Then
            lObservacao(30) = pMensagem
        Else
            lObservacao(30) = "Estourou "
        End If
    End If
End Sub
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
Private Sub Relatorio()
    Dim x_data As Date
    Dim Imprimiu As Boolean
    Dim Tempo As Date
    ''Tempo = Time
    
    Imprimiu = False
    ZeraVariaveis
    
    BuscaEntradaCombustivel
    BuscaMedicaoCombustivel
    BuscaMovimentoAfericao
    BuscaMovimentoBomba
    BuscaMovimentoNFeDevolucao
    If g_nome_usuario = "L.M.C." Then
        BuscaMovimentoBombaNormal 'new 05/08
    End If
    If chkImprimeResumo.Value = 1 Then
        x_data = CDate(msk_data_i.Text)
        'Loop data
        Do Until x_data > CDate(msk_data_f.Text)
            Call ImpDet(x_data)
            If chkRescunhoLmc.Value = 1 Then
                ImpRascunhoLmc (x_data)
            Else
                If chkDetalhadaBico.Value = 1 Then
                    Call ImpDetBico(x_data)
                End If
                If chkDetalhadaTanque.Value = 1 Then
                    Call ImpDetTanque(x_data)
                End If
            End If
            Imprimiu = True
            x_data = x_data + 1
        Loop
        If Imprimiu = True Then
            ImpResumoMes
            ImpTotal
            
            If ObtenhaChaveNotaDestinada Then
                 Dim xChaveAcessoNFeNaoEncontradas As New Dictionary
                 Set lChaveAcessoNFeDestinadasNaoCadastradas = EntradaCombustivel.IdentidicarNotasInexistentes(g_empresa, lChaveAcessoNFeDestinadas)
                
                 If lChaveAcessoNFeDestinadasNaoCadastradas.Count > 0 Then
                     Dim xEntradaProduto As New CadastroDLL.cEntradaProduto
                     
                     Set xChaveAcessoNFeNaoEncontradas = xEntradaProduto.IdentidicarNotasInexistentes(g_empresa, lChaveAcessoNFeDestinadasNaoCadastradas)
                     
                     If xChaveAcessoNFeNaoEncontradas.Count <> lChaveAcessoNFeDestinadasNaoCadastradas.Count Then
                          Set lChaveAcessoNFeDestinadasNaoCadastradas = xChaveAcessoNFeNaoEncontradas
                     End If
                 
                 End If
                 
                 Call ImprimeConsistenciaNFeDestinadas
            End If
            
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Relatório do Resumo do LMC|@|"
            ''MsgBox "Tempo gasto: " & DateDiff("s", Tempo, Time)
            frm_preview.Show 1
        End If
    End If
    If chkImprimeEntrada.Value = 1 Then
        If chkImprimeResumo.Value = 0 Then
            lTodosCombustiveis = False
            If (MsgBox("Imprime de todos os combustíveis?", vbQuestion + vbYesNo + vbDefaultButton2, "Combustíveis!")) = vbYes Then
                lTodosCombustiveis = True
            End If
        End If
        
        
        RelatorioEntrada
        
        
    End If
    If lPaginaOld > 0 Then
        txtPaginaInicial.Text = Format(lPaginaOld, "000")
    End If
End Sub
Private Function ObtenhaChaveNotaDestinada() As Boolean
    
    ObtenhaChaveNotaDestinada = False
    
    Dim xChavesAdicionadas As String
   
    If lChaveAcessoNFeDestinadas.Count > 0 Then
        lChaveAcessoNFeDestinadas.RemoveAll
    End If
    
    If lChaveAcessoNFeDestinadasNaoCadastradas.Count > 0 Then
        lChaveAcessoNFeDestinadasNaoCadastradas.RemoveAll
    End If
    
    
    
    lSQL = ""
    lSQL = lSQL & "SELECT [Chave de Acesso], [Ordem], [Data de Emissao], [Tipo de Combustivel] FROM NFeEntradaDestinadaItem"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data de Emissao] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND [Data de Emissao] <= " & preparaData(CDate(msk_data_f.Text))
    If lTodosCombustiveis = False Then
        lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(Mid(cbo_combustivel.Text, 1, 2))
    End If
    
    lSQL = lSQL & " ORDER BY [Data de Emissao], [Chave de Acesso], [Ordem], [Tipo de Combustivel]"
    
    Set rstNFeDestinada = New adodb.Recordset
    Set rstNFeDestinada = Conectar.RsConexao(lSQL)
    
    If rstNFeDestinada.RecordCount > 0 Then
        Do Until rstNFeDestinada.EOF
            If Len(xChavesAdicionadas) > 0 Then
                xChavesAdicionadas = xChavesAdicionadas & ","
            End If
            Dim xTipoCombustivel As String
            
            
            If (Len(Trim(rstNFeDestinada("Tipo de Combustivel").Value)) = 0) Then
                xTipoCombustivel = "Produto"
            Else
                xTipoCombustivel = rstNFeDestinada("Tipo de Combustivel").Value
            End If
            
            Call lChaveAcessoNFeDestinadas.Add(rstNFeDestinada("Chave de Acesso").Value & "|@|" & rstNFeDestinada("Ordem").Value & "|@|" & rstNFeDestinada("Tipo de Combustivel").Value & "|@|", rstNFeDestinada("Data de Emissao").Value) 'rstNFeDestinada("Ordem").Value)
            xChavesAdicionadas = xChavesAdicionadas + preparaTexto(rstNFeDestinada("Chave de Acesso").Value)
            rstNFeDestinada.MoveNext
        Loop
    End If
    
    
    lSQL = ""
    lSQL = lSQL & "SELECT [Chave de Acesso], 1 AS [Ordem], [Data de Emissao] FROM NFeEntradaDestinada"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [NFe Resumida] = " & preparaBooleano(True)
    lSQL = lSQL & "   AND [Data de Emissao] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND [Data de Emissao] <= " & preparaData(CDate(msk_data_f.Text))
    
    If Len(xChavesAdicionadas) > 0 Then
        lSQL = lSQL & "   AND [Chave de Acesso] not in ( " & xChavesAdicionadas & " )"
    End If

    lSQL = lSQL & " ORDER BY [Data de Emissao], [Chave de Acesso]"
    
    Set rstNFeDestinada = New adodb.Recordset
    Set rstNFeDestinada = Conectar.RsConexao(lSQL)

    If rstNFeDestinada.RecordCount > 0 Then
        Do Until rstNFeDestinada.EOF
            Call lChaveAcessoNFeDestinadas.Add(rstNFeDestinada("Chave de Acesso").Value & "|@|" & rstNFeDestinada("Ordem").Value & "|@|" & "Resumo" & "|@|", rstNFeDestinada("Data de Emissao").Value) ' rstNFeDestinada("Ordem").Value)
            rstNFeDestinada.MoveNext
        Loop
    End If
   
    If lChaveAcessoNFeDestinadas.Count > 0 Then
        ObtenhaChaveNotaDestinada = True
    End If
   
    If rstNFeDestinada.State = 1 Then
        rstNFeDestinada.Close
    End If
    

End Function


Private Sub RelatorioEntrada()
    ZeraVariaveisEntrada
    'Prepara SQL
    lSQL = ""
    'tarefa redmine 567
    lSQL = lSQL & "SELECT [Data], [Numero da Nota], [Valor do Litro], Quantidade, [Valor da Entrada],"
    lSQL = lSQL & "       [Codigo do Fornecedor], [Tipo de Combustivel]"
    lSQL = lSQL & "  FROM " & EntradaCombustivel.NomeTabela
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    'tarefa redmine 567
    lSQL = lSQL & "   AND [Data] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND [Data] <= " & preparaData(CDate(msk_data_f.Text))
'    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
'    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
    If lTodosCombustiveis = False Then
        lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(Mid(cbo_combustivel.Text, 1, 2))
    End If
    lSQL = lSQL & " ORDER BY [Data], [Numero da Nota], [Tipo de Combustivel]"
    'lSQL = lSQL & " ORDER BY Data, [Numero da Nota], [Tipo de Combustivel]"
    'Abre RecordSet
    Set rstEntradaCombustivel = New adodb.Recordset
    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rstEntradaCombustivel.RecordCount > 0 Then
        ImpCabEntrada
        ImpDadosEntrada
    End If
    If rstEntradaCombustivel.State = 1 Then
        rstEntradaCombustivel.Close
    End If
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex <> -1 Then
        l_tipo_combustivel = Mid(cbo_combustivel.Text, 1, 2)
        If Not Combustivel.LocalizarCodigo(g_empresa, l_tipo_combustivel) Then
            MsgBox "Combustível inexistente!", vbInformation, "Erro de Verificação!"
            cbo_combustivel.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_combustivel.SetFocus
    Else
        msk_data = RetiraGString(1)
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
    g_string = ""
    cbo_combustivel.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_combustivel.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            DoEvents
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cbo_combustivel.Text)
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione o combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    ElseIf chkImprimeEntrada.Value = 0 And chkImprimeResumo.Value = 0 Then
        MsgBox "Selecione um tipo de impressão.", vbInformation, "Atenção!"
        chkImprimeResumo.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    l_total_recebido = 0
    l_total_vendas = 0
    l_total_valor_vendas = 0
    l_total_afericao = 0
    l_perdas_sobras = 0
    lTotalCusto = 0
    lParcialEntrada = 0
    lParcialVenda = 0
    lParcialAfericao = 0
    lParcialPerdasSobras = 0
    lParcialMes = 0
    For iObs = 0 To 30
        lObservacao(iObs) = ""
    Next
    iObs = 0
End Sub
Private Sub ZeraVariaveisEntrada()
    lParcialEntrada = 0
    lParcialVenda = 0
    If chkImprimeResumo.Value = 1 Then
        lPaginaOld = Val(txtPaginaInicial.Text)
         txtPaginaInicial.Text = Format(lPagina + 1, "000")
    End If
    If chkImprimeResumo.Value = 0 Then
        lLinha = 0
    End If
    lPagina = 0
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        Call SelecionaImpressoraPadrao("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        AtivaBotoes (False)
        'If SelecionaImpressoraEpson(Me) Then
            DoEvents
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cbo_combustivel.Text)
            Relatorio
        'End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
        msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
        msk_data_i.SetFocus
        chkImprimeResumo.Value = 1
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
    
    MovimentoBombaNormal.NomeTabela = "Movimento_Bomba"
    If g_nome_usuario = "L.M.C." Then
        Me.Caption = Me.Caption & " - LMC"
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
    
    PreencheCboCombustivel
End Sub
Private Sub PreencheCboCombustivel()
    Dim rstCombustivel As New adodb.Recordset
    
    cbo_combustivel.Clear
    Set rstCombustivel = Conectar.RsConexao("SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstCombustivel
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_combustivel.AddItem !Codigo & " - " & !Nome
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstCombustivel = Nothing
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
        cbo_combustivel.SetFocus
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
Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub txtPaginaInicial_GotFocus()
    txtPaginaInicial.SelStart = 0
    txtPaginaInicial.SelLength = Len(txtPaginaInicial.Text)
End Sub
Private Sub txtPaginaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkImprimeEntrada.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtPaginaInicial_LostFocus()
    txtPaginaInicial.Text = Format(Val(txtPaginaInicial.Text), "000")
End Sub
