VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_venda_conveniencia 
   Caption         =   "Emissão das Vendas de Conveniencia"
   ClientHeight    =   3870
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_venda_conveniencia.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_venda_conveniencia.frx":030A
   ScaleHeight     =   3870
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_venda_conveniencia.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_venda_conveniencia.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_venda_conveniencia.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2940
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkCupomFiscal 
         Caption         =   "Cupom Fiscal"
         Height          =   255
         Left            =   3900
         TabIndex        =   22
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox chkNotaFiscal 
         Caption         =   "Nota Fiscal"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ComboBox cboCaixaFinal 
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1500
         Width           =   495
      End
      Begin VB.ComboBox cboCaixaInicial 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chk_acumulado 
         Caption         =   "Acumular Vendas"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1920
         Width           =   2055
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
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_venda_conveniencia.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_venda_conveniencia.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_venda_conveniencia.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
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
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Tip&o de venda"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Cai&xa final"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Cai&xa inicial"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Tip&o de de relatório"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   18
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_venda_conveniencia"
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
Dim lTotalQuantidade As Currency
Dim lTotalValor As Currency
Dim lValorCancelado As Currency
Dim lQtdCancelado As Currency
Dim lTotalComposicao As Currency
Dim lValorComposicao(0 To 30) As Currency
Dim lNomeComposicao(0 To 30) As String
Dim lSQl As String
Dim lCodigoProduto As Long
Dim lQtdCaixa As Integer

Private Configuracao As New cConfiguracao
Private Estoque As New cEstoque
Private Funcionario As New cFuncionario
Private MovJustificativa As New cMovimentoJustificativa
Dim rsMovComposicaoCaixa As New adodb.Recordset
Dim rsMovConveniencia As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Configuracao = Nothing
    Set rsMovComposicaoCaixa = Nothing
    Set Estoque = Nothing
    Set Funcionario = Nothing
    Set MovJustificativa = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotalQuantidade = 0
    lTotalValor = 0
    lValorCancelado = 0
    lQtdCancelado = 0
    lCodigoProduto = 0
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
    cbo_periodo_i.AddItem 5
    cbo_periodo_f.AddItem 5
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 5
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 5
End Sub
Private Sub PreencheCboCaixa()
    Dim i As Integer
    
    cboCaixaInicial.Clear
    cboCaixaFinal.Clear
    For i = 1 To lQtdCaixa
        cboCaixaInicial.AddItem i
        cboCaixaFinal.AddItem i
        cboCaixaInicial.ItemData(cboCaixaInicial.NewIndex) = i
        cboCaixaFinal.ItemData(cboCaixaFinal.NewIndex) = i
    Next
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    lSQl = ""
    lSQl = lSQl & " SELECT Movimento_Venda_Conveniencia.[Codigo do Produto], Movimento_Venda_Conveniencia.[Valor Unitario], Movimento_Venda_Conveniencia.Quantidade, (Movimento_Venda_Conveniencia.[Valor Total] - Movimento_Venda_Conveniencia.[Valor do Desconto]) AS [Valor Total], Movimento_Venda_Conveniencia.Data, Movimento_Venda_Conveniencia.Periodo, Movimento_Venda_Conveniencia.Operador, Movimento_Venda_Conveniencia.[Item Cancelado], Produto.Nome, Produto.Unidade, Movimento_Venda_Conveniencia.[Numero da Justificativa], Movimento_Venda_Conveniencia.[Valor do Desconto], "
    lSQl = lSQl & "        Movimento_Venda_Conveniencia.Hora"
    lSQl = lSQl & "   FROM Movimento_Venda_Conveniencia, Produto"
    lSQl = lSQl & "  WHERE Movimento_Venda_Conveniencia.Empresa = " & g_empresa
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Data >= " & preparaData(msk_data_i.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Data <= " & preparaData(msk_data_f.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Ilha >= " & Val(cboCaixaInicial.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.Ilha <= " & Val(cboCaixaFinal.Text)
    lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.[Codigo do Produto] = Produto.Codigo"
    If chkNotaFiscal.Value = 0 Or chkCupomFiscal.Value = 0 Then
        If chkNotaFiscal.Value = 1 Then
            lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.[Origem da Venda] LIKE " & preparaTexto("%CON%")
        ElseIf chkCupomFiscal.Value = 1 Then
            lSQl = lSQl & "    AND Movimento_Venda_Conveniencia.[Origem da Venda] LIKE " & preparaTexto("%ECF%")
        End If
    End If
    lSQl = lSQl & "  ORDER BY Movimento_Venda_Conveniencia.Data, Movimento_Venda_Conveniencia.Hora, Produto.Nome, Movimento_Venda_Conveniencia.Periodo, Movimento_Venda_Conveniencia.Operador"
    
    Set rsMovConveniencia = Conectar.RsConexao(lSQl)
    If rsMovConveniencia.RecordCount > 0 Then
        ImpDados
    End If
    rsMovConveniencia.Close
    Set rsMovConveniencia = Nothing
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    If chk_acumulado.Value = 0 Then
        LoopMovConveniencia
    Else
        LoopMovConvenienciaAcumulado
    End If
    LoopMovConvenienciaCancelados
    If lPagina > 0 Then
        ImpTotal
        ImpComposicaoCaixa
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório da Venda da Conveniencia|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopMovConveniencia()
    Dim xLinha As String
    
    With rsMovConveniencia
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 60 Then
                xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
                Mid(xLinha, 14, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            If ![Item Cancelado] = False Then
                Call ImpDet(!Data, !Hora, ![Codigo do Produto], !Nome, !Unidade, ![Valor Unitario], !Quantidade, ![Valor Total], ![Item Cancelado])
                lTotalQuantidade = lTotalQuantidade + !Quantidade
                lTotalValor = lTotalValor + ![Valor Total]
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopMovConvenienciaAcumulado()
    Dim xLinha As String
    Dim xQuantidade As Currency
    Dim xNome As String
    Dim xUnidade As String
    Dim xValorUnitario As Currency
    Dim xValorTotal As Currency
    Dim xValorVenda As Currency
    Dim xData As Date
    Dim xPeriodo As String
    Dim xCodigoFuncionario As Integer
    
    With rsMovConveniencia
        Do Until .EOF
            If lCodigoProduto <> ![Codigo do Produto] Then
                If xValorTotal > 0 Then
                    If lPagina = 0 Then
                        ImpCab
                    End If
                    If Estoque.LocalizarCodigo(g_empresa, lCodigoProduto) Then
                        If Estoque.PrecoVenda > 0 Then
                            xValorVenda = Estoque.PrecoVenda
                        End If
                    End If
                    xValorUnitario = xValorTotal / xQuantidade
                    Call ImpDet(CDate("00:00:00"), CDate("00:00:00"), lCodigoProduto, xNome, xUnidade, xValorUnitario, xQuantidade, xValorTotal, False)
                    xQuantidade = 0
                    xValorTotal = 0
                    If lLinha >= 60 Then
                        xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
                        Mid(xLinha, 14, 22) = " Cerrado Informática. "
                        BioImprime "@Printer.Print " & xLinha
                        BioImprime "@@Printer.NewPage"
                        ImpCab
                    End If
                End If
                lCodigoProduto = ![Codigo do Produto]
            End If
            If ![Item Cancelado] = False Then
                xValorTotal = xValorTotal + ![Valor Total]
            End If
            If ![Item Cancelado] = False Then
                lTotalQuantidade = lTotalQuantidade + !Quantidade
                lTotalValor = lTotalValor + ![Valor Total]
            End If
            xNome = !Nome
            xUnidade = !Unidade
            xQuantidade = xQuantidade + !Quantidade
            xData = !Data
            xPeriodo = !Periodo
            xCodigoFuncionario = !operador
            .MoveNext
        Loop
    End With
    If xValorTotal > 0 Then
        If lPagina = 0 Then
            ImpCab
        End If
        If Estoque.LocalizarCodigo(g_empresa, lCodigoProduto) Then
            If Estoque.PrecoVenda > 0 Then
                xValorVenda = Estoque.PrecoVenda
            End If
        End If
        xValorUnitario = xValorTotal / xQuantidade
        Call ImpDet(CDate("00:00:00"), CDate("00:00:00"), lCodigoProduto, xNome, xUnidade, xValorUnitario, xQuantidade, xValorTotal, False)
        xQuantidade = 0
        xValorTotal = 0
    End If

End Sub
Private Sub LoopMovConvenienciaCancelados()
    Dim xLinha As String
    Dim xQtd As Integer
    
    xQtd = 0
    With rsMovConveniencia
        .MoveFirst
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 60 Then
                xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
                Mid(xLinha, 14, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            If ![Item Cancelado] Then
'                If xQtd = 0 Then
'                    xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
'                    BioImprime "@Printer.Print " & xLinha
'                End If
                xQtd = xQtd + 1
                Call ImpDet(!Data, !Hora, ![Codigo do Produto], !Nome, !Unidade, ![Valor Unitario], !Quantidade, ![Valor Total], ![Item Cancelado])
                lQtdCancelado = lQtdCancelado + !Quantidade
                lValorCancelado = lValorCancelado + ![Valor Total]
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpDet(ByVal pData As Date, ByVal pHora As Date, ByVal pCodigoProduto As Long, ByVal pNome As String, ByVal pUN As String, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pCancelado As Boolean)
    Dim xLinha As String
    Dim i As Integer
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+----------+--------+-------+----------------------------------------+---+-------------+--------------------+--------------------+------+"
        Mid(xLinha, 15, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|          |        |       |                                        |   |             |                    |                    |      |"
    If pData <> "00:00:00" Then
        Mid(xLinha, 2, 10) = Format(pData, "dd/MM/yyyy")
    End If
    If pHora <> "00:00:00" Then
        Mid(xLinha, 13, 8) = Format(pHora, "HH:mm:ss")
    End If
    i = Len(Format(pCodigoProduto, "#,000"))
    Mid(xLinha, 24 + 5 - i, i) = Format(pCodigoProduto, "#,000")
    Mid(xLinha, 30, 40) = pNome
    Mid(xLinha, 71, 3) = pUN
    i = Len(Format(pQuantidade, "####,##0.00"))
    Mid(xLinha, 76 + 11 - i, i) = Format(pQuantidade, "####,##0.00")
    i = Len(Format(pValorUnitario, "###,###,##0.00"))
    Mid(xLinha, 92 + 14 - i, i) = Format(pValorUnitario, "###,###,##0.00")
    i = Len(Format(pValorTotal, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(pValorTotal, "###,###,##0.00")
    If pCancelado Then
        BioImprime "@Printer.Print " & "+----------+--------+-------+----------------------------------------+---+-------------+--------------------+--------------------+------+"
        Mid(xLinha, 132, 4) = "CANC"
        If rsMovConveniencia![Numero da Justificativa] > 0 Then
            BioImprime "@Printer.Print " & xLinha
            xLinha = "|                                                                                                                                       |"
            If MovJustificativa.LocalizarCodigo(rsMovConveniencia![Numero da Justificativa]) Then
                i = Len(Format(MovJustificativa.numero, "###,##0"))
                Mid(xLinha, 3 + 7 - i, i) = Format(MovJustificativa.numero, "###,##0")
                Mid(xLinha, 11, 50) = MovJustificativa.Justificativa
                Mid(xLinha, 62, 10) = Format(MovJustificativa.Data, "dd/MM/yyyy")
                Mid(xLinha, 73, 8) = Format(MovJustificativa.Hora, "HH:mm:ss")
                Mid(xLinha, 82, 25) = MovJustificativa.Computador
                If Funcionario.LocalizarCodigo(g_empresa, MovJustificativa.CodigoFuncionario) Then
                    Mid(xLinha, 108, 28) = Funcionario.Nome
                Else
                    Mid(xLinha, 108, 28) = "** NÃO CADASTRADO **        "
                End If
            Else
                Mid(xLinha, 17, 50) = "** SEM JUSTIFICATIVA **"
            End If
            lLinha = lLinha + 2
        End If
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpComposicaoCaixa()
    Dim xLinha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xQtdComposicao As Integer
    Dim xFaltaCaixa As Currency
    
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "   SELECT Composicao_Caixa.Ordem,"
    lSQl = lSQl & "          Movimento_Composicao_Caixa.[Codigo da Composicao],"
    lSQl = lSQl & "          SUM(Valor) AS Total,"
    lSQl = lSQl & "          Composicao_Caixa.Nome AS NomeComposicao"
    lSQl = lSQl & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    lSQl = lSQl & "    WHERE Movimento_Composicao_Caixa.Empresa = " & g_empresa
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] = " & 3
    lSQl = lSQl & "      AND Composicao_Caixa.Codigo = Movimento_Composicao_Caixa.[Codigo da Composicao]"
    lSQl = lSQl & " GROUP BY Ordem, [Codigo da Composicao], Composicao_Caixa.Nome"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQl)
    For i = 0 To 30
        lValorComposicao(i) = 0
        lNomeComposicao(i) = ""
    Next
    lTotalComposicao = 0
    i = -1
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        xQtdComposicao = rsMovComposicaoCaixa.RecordCount
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            lValorComposicao(i) = rsMovComposicaoCaixa("Total").Value
            lNomeComposicao(i) = rsMovComposicaoCaixa("NomeComposicao").Value
            lTotalComposicao = lTotalComposicao + rsMovComposicaoCaixa("Total").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    
    
    BioImprime "@Printer.Print " & "+---------------------------------------+--------------------------------------+"
    BioImprime "@Printer.Print " & "| COMPOSIÇÃO DO CAIXA                   |                                      |"
    BioImprime "@Printer.Print " & "+---------------------------------------+--------------------------------------+"
    For i = 0 To xQtdComposicao - 1
        xLinha = "| .....................:                |                                      |"
        Mid(xLinha, 3, 20) = lNomeComposicao(i)
        If lValorComposicao(i) > 0 Then
            i2 = Len(Format(lValorComposicao(i), "#,###,##0.00"))
            Mid(xLinha, 28 + 12 - i2, i2) = Format(lValorComposicao(i), "#,###,##0.00")
        End If
        BioImprime "@Printer.Print " & xLinha
    Next
    xLinha = "| Total................:                |                                      |"
    xFaltaCaixa = lTotalComposicao - lTotalValor
    If xFaltaCaixa = 0 Then
        Mid(xLinha, 43, 20) = "       CAIXA OK!"
    ElseIf xFaltaCaixa < 0 Then
        Mid(xLinha, 43, 20) = "Falta de Caixa.:"
        i = Len(Format(xFaltaCaixa, "#,###,##0.00"))
        Mid(xLinha, 67 + 12 - i, i) = Format(xFaltaCaixa, "#,###,##0.00")
    Else
        Mid(xLinha, 43, 20) = "Passou no Caixa:"
        i = Len(Format(xFaltaCaixa, "#,###,##0.00"))
        Mid(xLinha, 67 + 12 - i, i) = Format(xFaltaCaixa, "#,###,##0.00")
    End If
    If lTotalComposicao > 0 Then
        i = Len(Format(lTotalComposicao, "#,###,##0.00"))
        Mid(xLinha, 28 + 12 - i, i) = Format(lTotalComposicao, "#,###,##0.00")
    End If
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & "+---------------------------------------+--------------------------------------+"
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    Dim xValor As Currency
    
    xLinha = "+----------+--------+-------+----------------------------------------+---+-------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                      *** TOTAL DAS VENDAS              |             |                    |                    |      |"
    i = Len(Format(lTotalQuantidade, "####,##0.00"))
    Mid(xLinha, 76 + 11 - i, i) = Format(lTotalQuantidade, "####,##0.00")
    i = Len(Format(lTotalValor, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(lTotalValor, "###,###,##0.00")
    
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 7"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontBold = False"
    
    xLinha = "|                                      *** TOTAL DE CANCELAMENTOS        |             |                    |                    |      |"
    i = Len(Format(lQtdCancelado, "####,##0.00"))
    Mid(xLinha, 76 + 11 - i, i) = Format(lQtdCancelado, "####,##0.00")
    i = Len(Format(lValorCancelado, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(lValorCancelado, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "|                                      *** TOTAL GERAL                   |             |                    |                    |      |"
    xValor = lQtdCancelado + lTotalQuantidade
    i = Len(Format(xValor, "####,##0.00"))
    Mid(xLinha, 76 + 11 - i, i) = Format(xValor, "####,##0.00")
    xValor = lValorCancelado + lTotalValor
    i = Len(Format(xValor, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(xValor, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "+------------------------------------------------------------------------+-------------+--------------------+--------------------+------+"
    Mid(xLinha, 3, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontName = Courier New"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 10"
    BioImprime "@@Printer.CurrentY = 0"
    '                   1         2         3         4         5         6         7         8
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| VENDAS DE CONVENIENCIA (GERAL)                                  , __/__/____ |"
    
    
    If chkNotaFiscal.Value = 0 Or chkCupomFiscal.Value = 0 Then
        If chkNotaFiscal.Value = 1 Then
            Mid(xLinha, 26, 7) = "(NF)   "
        ElseIf chkCupomFiscal.Value = 1 Then
            Mid(xLinha, 26, 7) = "(ECF)  "
        End If
    End If
    
    
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____.                           |"
    Mid(xLinha, 29, 10) = msk_data_i.Text
    Mid(xLinha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| PERIODO INICIAL.........: X    PERIODO FINAL: X                              |"
    Mid(xLinha, 29, 1) = cbo_periodo_i.Text
    Mid(xLinha, 49, 1) = cbo_periodo_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(xLinha, 29, 1) = cboCaixaInicial.Text
    Mid(xLinha, 49, 1) = cboCaixaFinal.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 7"
    xLinha = "+----------+--------+-------+----------------------------------------+---+-------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|   DATA   |  HORA  | CODIGO|DISCRIMINAÇÃO DOS PRODUTOS              |UN.| QTD. VENDAS |  PRECO  DE  VENDA  |  TOTAL  DA  VENDA  | SIT. |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+--------+-------+----------------------------------------+---+-------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboCaixaInicial.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cboCaixaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cboCaixaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboCaixaFinal.ListIndex = cboCaixaInicial.ListIndex
        cboCaixaFinal.SetFocus
    End If
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
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Dados incompleto!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Dados incompleto!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Dados incompleto!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Dados incompleto!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Dados incompleto!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Dados incompleto!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Dados incompleto!"
        cbo_periodo_f.SetFocus
    ElseIf cboCaixaInicial.ListIndex = -1 Then
        MsgBox "Selecione o caixa inicial.", vbInformation, "Dados incompleto!"
        cboCaixaInicial.SetFocus
    ElseIf cboCaixaFinal.ListIndex = -1 Then
        MsgBox "Selecione o caixa final.", vbInformation, "Dados incompleto!"
        cboCaixaFinal.SetFocus
    ElseIf cboCaixaFinal.Text < cboCaixaInicial.Text Then
        MsgBox "O caixa final deve ser maior.", vbInformation, "Dados incompleto!"
        cboCaixaFinal.SetFocus
    ElseIf chkNotaFiscal.Value = 0 And chkCupomFiscal.Value = 0 Then
        MsgBox "Marque pelo menos um tipo de venda.", vbInformation, "Dados incompleto!"
        chkNotaFiscal.SetFocus
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
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 3
        cboCaixaInicial.ListIndex = 0
        cboCaixaFinal.ListIndex = 0
        chkNotaFiscal.Value = 1
        cmd_imprimir.SetFocus
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
    lQtdCaixa = 1
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdCaixa = Configuracao.QuantidadeIlha
    End If
    PreencheCboCaixa
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
        msk_data_f.Text = msk_data_i.Text
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
