VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_caixa_pista 
   Caption         =   "Emissão do Caixa de Pista"
   ClientHeight    =   5580
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   7155
   Icon            =   "emissao_caixa_pista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   7155
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1260
      Picture         =   "emissao_caixa_pista.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Visualiza Caixa (Simplificado)."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3180
      Picture         =   "emissao_caixa_pista.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Imprime Caixa (Simplificado)."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5100
      Picture         =   "emissao_caixa_pista.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4620
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.CheckBox chkImprimeConciliacao 
         Caption         =   "Imprime Conciliação de Cartão"
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   4140
         Width           =   2655
      End
      Begin VB.ComboBox cbo_funcionario 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1020
         Width           =   4875
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   8
         Top             =   660
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
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkImprimeDespesaCaixa 
         Caption         =   "Imprime Despesas de Caixa"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   3840
         Width           =   3495
      End
      Begin VB.CheckBox chkImprimeNotaAbastecimento 
         Caption         =   "Imprime Notas de Abastecimento"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   3540
         Width           =   3495
      End
      Begin VB.ComboBox cboTipoCaixa 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2220
         Width           =   2175
      End
      Begin VB.CheckBox chkImprimeMedidaTanque 
         Caption         =   "Imprime Medida de Combustíveis"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CheckBox chkImprimeLucro 
         Caption         =   "Imprime Lucro das Vendas"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   2940
         Width           =   3495
      End
      Begin VB.ComboBox cboIlhaF 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1860
         Width           =   615
      End
      Begin VB.ComboBox cboIlhaI 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1860
         Width           =   615
      End
      Begin VB.CheckBox chkImprimeLubrificante 
         Caption         =   "Imprime Venda de Lubrificante Detalhada"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "emissao_caixa_pista.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "emissao_caixa_pista.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6300
         Picture         =   "emissao_caixa_pista.frx":6C74
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
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "&Tipo de Caixa"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Ilha Final"
         Height          =   300
         Left            =   3840
         TabIndex        =   18
         Top             =   1860
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Ilha Inicial"
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1500
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
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_caixa_pista"
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
Dim lBombaAfericaoLitros(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaAfericaoValor(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaLitros(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaValorTotal(1 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lBombaTipoPreco(1 To gQUANTIDADE_MAXIMA_BICO) As String
Dim lBombaCombustivel(1 To gQUANTIDADE_MAXIMA_BICO) As String
Dim lBombaMensagem(1 To gQUANTIDADE_MAXIMA_BICO) As String
Dim lBombaLitrosAfericao As Currency
Dim lBombaValorAfericao As Currency
Dim lBombaLitrosA As Currency
Dim lBombaLitrosAA As Currency
Dim lBombaLitrosD As Currency
Dim lBombaLitrosDA As Currency
Dim lBombaLitrosG As Currency
Dim lBombaLitrosGA As Currency
Dim lBombaTotalLitros As Currency
Dim lBombaValorA As Currency
Dim lBombaValorAA As Currency
Dim lBombaValorD As Currency
Dim lBombaValorDA As Currency
Dim lBombaValorG As Currency
Dim lBombaValorGA As Currency
Dim lBombaCustoA As Currency
Dim lBombaCustoAA As Currency
Dim lBombaCustoD As Currency
Dim lBombaCustoDA As Currency
Dim lBombaCustoG As Currency
Dim lBombaCustoGA As Currency
Dim lTotalLubrificante As Currency
Dim lBombaTotalValor As Currency
Dim lBombaTotalAcrescimo As Currency
Dim lBombaTotalDesconto As Currency

Dim lTotalEntrada As Currency
Dim lTotalSaida As Currency
Dim lTotalVista As Currency
Dim lTotalPrazo As Currency
Dim lTotalNota As Currency
Dim lVendaParcialCombustivel As Boolean
Dim lTotalDuplicataRecebida As Currency
Dim lTotalBaixaCheque As Currency
Dim lTotalBaixaChequeDevolvido As Currency
Dim lExecutaActivate As Boolean
Dim lExistePendenciaCartao As Boolean

Dim lQtdComposicao As Integer
Dim lTotalComposicao As Currency
Dim lUltimoBico As Integer
Dim lTipoMovimento As Integer
Dim lInverteEncerrante As Boolean
Dim lDataPeriodo As String
Dim lUsarEncerranteMecanico As Boolean

Dim lSQL As String
Dim rs As New ADODB.Recordset
Dim rsDataPeriodo As New ADODB.Recordset
Private rsMovLubrificante As New ADODB.Recordset
Private rsMovimentoCaixaPista As New ADODB.Recordset
Private rsMovimentoBomba As New ADODB.Recordset
Private rsNotaAbastecimento As New ADODB.Recordset
Private rsBaixaDuplicataReceber As New ADODB.Recordset
Private rsBaixaCheque As New ADODB.Recordset
Private rsBaixaChequeDevolvido As New ADODB.Recordset

Private AberturaCaixa As New cAberturaCaixa
Private Bomba As New cBomba
Private CartaoCredito As New cCartaoCredito
Private Combustivel As New cCombustivel
Private ConciliacaoCartao As New cConciliacaoCartao
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private EncerranteAtual As New cEncerranteAtual
Private EntradaCombustivel As New cEntradaCombustivel
Private Funcionario As New cFuncionario
Private LancamentoPadrao As New cLancamentoPadrao
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoBombaMec As New cMovimentoBomba
Private PeriodoTrocaOleo As New cPeriodoTrocaOleo
Private rsTotalizador As New ADODB.Recordset
Private Sub AjustaCaixaPista()
    Dim xString As String
    
    xString = g_string
    g_string = ""

    txtDataInicial.Text = Format(CDate(RetiraString(3, xString)), "dd/mm/yyyy")
    txtDataFinal.Text = Format(CDate(RetiraString(3, xString)), "dd/mm/yyyy")
    cbo_periodo_i.ListIndex = Val(RetiraString(4, xString)) - 1
    cbo_periodo_f.ListIndex = Val(RetiraString(4, xString)) - 1
    lTipoMovimento = Val(RetiraString(5, xString))
    cboIlhaI.ListIndex = Val(RetiraString(6, xString)) - 1
    cboIlhaF.ListIndex = Val(RetiraString(6, xString)) - 1
    cboTipoCaixa.ListIndex = lTipoMovimento
    chkImprimeLubrificante.Value = 1
    If UCase(g_nome_empresa) Like "*JOSE OSVALDO*" Then
        chkImprimeLubrificante.Value = 0
    End If
    If fEcfInstalada = False Then
        If Not MovimentoBomba.ExisteMovimentoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text)) Then
            Me.Caption = Me.Caption & " - ECF"
            MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
        End If
    End If
    If RetiraString(2, xString) = "Visualizar" Then
        cmd_visualizar_Click
    Else
        cmd_imprimir_Click
    End If
    cmd_sair_Click
End Sub
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        lExecutaActivate = False
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        lExecutaActivate = True
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    
    Set rsMovLubrificante = Nothing
    Set rsMovimentoCaixaPista = Nothing
    Set rsMovimentoBomba = Nothing
    Set rsBaixaCheque = Nothing
    
    Set AberturaCaixa = Nothing
    Set Bomba = Nothing
    Set CartaoCredito = Nothing
    Set Combustivel = Nothing
    Set ConciliacaoCartao = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set EncerranteAtual = Nothing
    Set EntradaCombustivel = Nothing
    Set Funcionario = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoBombaMec = Nothing
End Sub
Public Sub FinalizaProcessoCaixa()
    Dim xArquivoDiscoTMP As TextStream
    
    On Error GoTo FileError
    
    Set xArquivoDiscoTMP = gArqTxt.CreateTextFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP")
    xArquivoDiscoTMP.WriteLine ("[Outras]")
    xArquivoDiscoTMP.WriteLine ("gString=" & g_string)
    xArquivoDiscoTMP.Close
    
    Call gArqTxt.MoveFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP", "C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.INI")
    Exit Sub
FileError:
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    lQtdComposicao = 0
    lTotalComposicao = 0
    lTotalEntrada = 0
    lTotalVista = 0
    lTotalPrazo = 0
    lTotalSaida = 0

    For i = 1 To gQUANTIDADE_MAXIMA_BICO
        lBombaAbertura(i) = 0
        lBombaEncerrante(i) = 0
        lBombaAfericaoLitros(i) = 0
        lBombaAfericaoValor(i) = 0
        lBombaLitros(i) = 0
        lBombaValorTotal(i) = 0
        lBombaTipoPreco(i) = ""
        lBombaCombustivel(i) = ""
        lBombaMensagem(i) = ""
    Next
    lBombaLitrosAfericao = 0
    lBombaValorAfericao = 0
    lBombaLitrosA = 0
    lBombaLitrosAA = 0
    lBombaLitrosD = 0
    lBombaLitrosDA = 0
    lBombaLitrosG = 0
    lBombaLitrosGA = 0
    lBombaTotalLitros = 0
    lBombaValorA = 0
    lBombaValorAA = 0
    lBombaValorD = 0
    lBombaValorDA = 0
    lBombaValorG = 0
    lBombaValorGA = 0
    lBombaCustoA = 0
    lBombaCustoAA = 0
    lBombaCustoD = 0
    lBombaCustoDA = 0
    lBombaCustoG = 0
    lBombaCustoGA = 0
    lTotalLubrificante = 0
    lBombaTotalValor = 0
    lTotalNota = 0
    lTotalDuplicataRecebida = 0
    lTotalBaixaCheque = 0
    lTotalBaixaChequeDevolvido = 0
    lDataPeriodo = ""
    lExistePendenciaCartao = False
    lBombaTotalAcrescimo = 0
    lBombaTotalDesconto = 0
End Sub
Private Sub PreparaDataPeriodo()
    Dim i As Integer
    
    lDataPeriodo = ""
    lSQL = ""
    lSQL = lSQL & "SELECT Convert(VarChar, [Data da Abertura], 103) + Convert(VarChar, Periodo) AS DataPeriodo"
    lSQL = lSQL & "     FROM AberturaCaixa"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Data da Abertura] >= " & preparaData(txtDataInicial.Text)
    lSQL = lSQL & "      AND [Data da Abertura] <= " & preparaData(txtDataFinal.Text)
    lSQL = lSQL & "      AND [Codigo do Funcionario] = " & Val(cbo_funcionario.ItemData(cbo_funcionario.ListIndex))
    lSQL = lSQL & " ORDER BY DataPeriodo"
    'Abre RecordSet
    Set rsDataPeriodo = New ADODB.Recordset
    Set rsDataPeriodo = Conectar.RsConexao(lSQL)
    
    If rsDataPeriodo.RecordCount > 0 Then
        lDataPeriodo = "("
        i = 0
        rsDataPeriodo.MoveFirst
        Do Until rsDataPeriodo.EOF
            i = i + 1
            If i > 1 Then
                lDataPeriodo = lDataPeriodo & ", " & preparaTexto(rsDataPeriodo("DataPeriodo").Value)
            Else
                lDataPeriodo = lDataPeriodo & preparaTexto(rsDataPeriodo("DataPeriodo").Value)
            End If
            rsDataPeriodo.MoveNext
        Loop
        lDataPeriodo = lDataPeriodo & ")"
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica se exiete composição de caixa
    
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        PreparaDataPeriodo
    End If
    
    lTipoMovimento = cboTipoCaixa.ItemData(cboTipoCaixa.ListIndex)
    If lTipoMovimento > 0 Then
        If Not LocalizarAberturaCaixa(Val(cbo_funcionario.ItemData(cbo_funcionario.ListIndex)), CDate(txtDataInicial.Text), Val(cbo_periodo_i.Text), Val(cboIlhaI.Text), lTipoMovimento) Then
        'If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(txtDataInicial.Text), "NF", Val(cbo_periodo_i.Text), Val(cboIlhaI.Text), lTipoMovimento) Then
            MsgBox "Não existe caixa aberto nesta data!", vbOKOnly + vbInformation, "Caixa Inexistente!"
            txtDataInicial.SetFocus
            Exit Sub
        End If
    End If
    TotalizaCaixaPista
    TotalizaLubrificante
    If g_nome_empresa Like "*VENTANIA*" Then
    Else
        TotalizaBaixaDuplicataAReceber
        TotalizaBaixaCheque
        TotalizaBaixaChequeDevolvido
    End If
    
    
    If CBool(chkImprimeConciliacao.Value) = True Then
        Call LoopConciliacaoCartao(0)
    End If
    
    'If MovimentoBomba.ExisteMovimentoPeriodo(g_empresa, CDate(txtdatainicial.Text), CDate(txtdatafinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_i.Text)) Then
        ImpDados
    'End If
    'cmd_sair.SetFocus
    If g_nivel_acesso = 4 Or g_usuario = 8 Then
        txtDataInicial.Text = Format(CDate(txtDataInicial.Text) + 1, "dd/mm/yyyy")
        txtDataFinal.Text = Format(CDate(txtDataFinal.Text) + 1, "dd/mm/yyyy")
        cmd_imprimir.Enabled = True
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub ImpDados()
    Dim i As Integer
    Dim xAbertura As Currency
    Dim xEncerrante As Currency
    Dim xLitros As Currency
    Dim xValorTotal As Currency
    
    CalculaAfericaoPorBico 'Aqui calcula as afericoes por bico
    If gUfEmpresa = "PA" Then
        LoopMovimentoBomba
    Else
        LoopMovimentoBombaRS
    End If
    CalculaAfericao 'Aqui calcula as afericoes por tipo de combustivel
    'If lBombaTotalLitros > 0 Then
        ImpCab
        ImpCabCombustivel
        If lTipoMovimento = 2 Then
            For i = 1 To lUltimoBico
                Call ImpDetBomba(i, lBombaAbertura(i), lBombaEncerrante(i), lBombaAfericaoLitros(i), lBombaLitros(i), lBombaValorTotal(i), lBombaTipoPreco(i), lBombaCombustivel(i))
                If lUsarEncerranteMecanico Then
                    If MovimentoBombaMec.LocalizarPrimeiroPeriodoBico(g_empresa, CDate(txtDataInicial.Text), i, 0) Then
                        xLitros = MovimentoBombaMec.TotalLitrosBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), i, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
                        xValorTotal = MovimentoBombaMec.TotalValorBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), i, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
                        If xLitros <> lBombaLitros(i) Then
                            xAbertura = MovimentoBombaMec.AberturaBicoDataPeriodo(g_empresa, CDate(txtDataInicial.Text), i, Val(cbo_periodo_i.Text))
                            xEncerrante = MovimentoBombaMec.EncerranteBicoDataPeriodo(g_empresa, CDate(txtDataInicial.Text), i, Val(cbo_periodo_f.Text))
                            Call ImpDetBomba(i, xAbertura, xEncerrante, 0, xLitros, xValorTotal, "DIFERENCA MEC.", "DIFERENCA MEC.")
                        End If
                    Else
                        Call ImpDetBomba(99, 0, 0, 0, 0, 0, "SEM MOV.MEC.", "SEM MOV.MEC.")
                    End If
                End If
            Next
        End If
        If Not UCase(g_cidade_empresa) Like "REDEN*" Then
            LoopMovimentoLubrificante
        End If
        LoopMovimentoCaixaPista
        
        ImpResumoCombustiveis
        If chkImprimeLucro.Value = 1 Then
            ImpResumoLucroCombustiveis
        End If
        If chkImprimeMedidaTanque.Value = 1 Then
            ImpResumoMedicaoCombustiveis
        End If
        If UCase(g_cidade_empresa) Like "REDEN*" Then
            LoopMovimentoLubrificante
        End If
        If chkImprimeNotaAbastecimento.Value = 1 Then
            LoopMovimentoNotaAbastecimento
            LoopBaixaNotaAbastecimento
            If g_nome_empresa Like "*VENTANIA*" Then
            Else
                LoopBaixaCheque
                LoopBaixaChequeDevolvido
            End If
        End If
        If chkImprimeDespesaCaixa.Value = 1 Then
            LoopMovimentoDespesaCaixa
        End If
        If g_nome_empresa Like "*VENTANIA*" Then
        Else
            LoopBaixaDuplicataAReceber
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Caixa de Pista|@|"
        frm_preview.Show 1
    'End If
End Sub
Private Function LocalizarAberturaCaixa(ByVal pCodigoFuncionario As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pIlha As Integer, ByVal pTipoCaixa As Integer) As Boolean
    LocalizarAberturaCaixa = False
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        If AberturaCaixa.LocalizarCxDataFunc(g_empresa, pData, "NF", pPeriodo, pIlha, pTipoCaixa, pCodigoFuncionario) Then
            LocalizarAberturaCaixa = True
        End If
    Else
        If AberturaCaixa.LocalizarCxData(g_empresa, pData, "NF", pPeriodo, pIlha, pTipoCaixa) Then
            LocalizarAberturaCaixa = True
        End If
    End If
End Function
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
Private Sub LoopBaixaDuplicataAReceber()
    Dim xLinha As String
    Dim i As Integer
    Dim xTotalDinheiro As Currency
    Dim xTotalChVista As Currency
    Dim xTotalChPrazo As Currency
    Dim xTotalBanco As Currency
    Dim xTotalCartao As Currency
    Dim xTotal As Currency
    Dim xContadorFormaPagamento As Integer
    
    Const FORMA_PAGAMENTO_DINHEIRO As String = "DINHEIRO"
    Const FORMA_PAGAMENTO_CH_VISTA As String = "CH VISTA"
    Const FORMA_PAGAMENTO_CH_PRAZO As String = "CH PRAZO"
    Const FORMA_PAGAMENTO_BANCO As String = "   BANCO"
    Const FORMA_PAGAMENTO_CARTAO As String = "  CARTÃO"
    
    Const TAM_CAMPO_FORMA_PAGAMENTO As Integer = 12
    Dim xTamanhoFormaPagamento As Integer
    
    xContadorFormaPagamento = 0
    'Verifica movimento
    If rsBaixaDuplicataReceber.RecordCount > 0 Then
        If lLinha >= 60 Then
            xLinha = "+-------+------------------------------------------------+--------+------------+"
            Mid(xLinha, 11, 21) = " Cerrado Informatica "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpCabBaixaDuplicataAReceber
        xTotalDinheiro = 0
        xTotalChVista = 0
        xTotalChPrazo = 0
        xTotalCartao = 0
        xTotal = 0
        
        'loop baixa de duplicata a receber
        rsBaixaDuplicataReceber.MoveFirst
        Do Until rsBaixaDuplicataReceber.EOF
            If lLinha >= 62 Then
                xLinha = "+-------+------------------------------------------------+--------+------------+"
                Mid(xLinha, 11, 21) = " Cerrado Informatica "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
                ImpCabBaixaDuplicataAReceber
            End If
            
            xContadorFormaPagamento = 0

            
            '         1         2         3         4         5         6         7         8         9        10        11        12        13     13
            '12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
            '+-------+------------------------------------------------+--------+------------+
            '|CÓDIGO | RAZÃO SOCIAL                                   |NUM.DOC.|    VALOR   |
            '|CLIENTE| DT.INICIAL  DT.  FINAL  DT.VENCIM.    VLR.VENC |RECEB.EM|  RECEBIDO  |
            '+-------+------------------------------------------------+--------+------------+
            '|   331 |COMERCIAL DK LTDA.                              |208.235 |      231,26|
            '|       | 01/02/2010  28/02/2010  10/03/2010      231,26 |13/03 P1|            |
            xLinha = "|       |                                                |        |            |"
            i = Len(Format(rsBaixaDuplicataReceber("Codigo do Cliente").Value, "##,##0"))
            Mid(xLinha, 2 + 6 - i, i) = Format(rsBaixaDuplicataReceber("Codigo do Cliente").Value, "##,##0")
            Mid(xLinha, 10, 40) = rsBaixaDuplicataReceber("Razao Social").Value
            i = Len(Format(rsBaixaDuplicataReceber("Numero do Documento").Value, "##,##0"))
            Mid(xLinha, 59 + 7 - i, i) = Format(rsBaixaDuplicataReceber("Numero do Documento").Value, "##,##0")
            i = Len(Format(rsBaixaDuplicataReceber("TotalRecebido").Value, "#,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("TotalRecebido").Value, "#,###,##0.00")
            
            BioImprime "@Printer.Print " & xLinha
            
            xLinha = "|       |                                                |        |            |"
            
            Mid(xLinha, 11, 10) = Format(rsBaixaDuplicataReceber("Data do Periodo Inicial").Value, "dd/MM/yyyy")
            Mid(xLinha, 23, 10) = Format(rsBaixaDuplicataReceber("Data do Periodo Final").Value, "dd/MM/yyyy")
            Mid(xLinha, 35, 10) = Format(rsBaixaDuplicataReceber("Data do Vencimento").Value, "dd/MM/yyyy")
            i = Len(Format(rsBaixaDuplicataReceber("Valor do Vencimento").Value, "#,###,##0.00"))
            Mid(xLinha, 46 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor do Vencimento").Value, "#,###,##0.00")
            Mid(xLinha, 59, 5) = Format(rsBaixaDuplicataReceber("Data do Pagamento").Value, "dd/MM")
            Mid(xLinha, 65, 2) = "P" & rsBaixaDuplicataReceber("Periodo").Value

            If rsBaixaDuplicataReceber("Valor Pago Dinheiro").Value > 0 Then
                i = Len(Format(rsBaixaDuplicataReceber("Valor Pago Dinheiro").Value, "#,###,##0.00"))
                Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor Pago Dinheiro").Value, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                
                xLinha = "|       |                                                |        |            |"
                
                xTamanhoFormaPagamento = Len(FORMA_PAGAMENTO_DINHEIRO)
                Mid(xLinha, 68 + TAM_CAMPO_FORMA_PAGAMENTO - xTamanhoFormaPagamento, xTamanhoFormaPagamento) = FORMA_PAGAMENTO_DINHEIRO
                BioImprime "@Printer.Print " & xLinha
                xContadorFormaPagamento = xContadorFormaPagamento + 1
                xLinha = "|       |                                                |        |            |"
            End If
            
            
            If rsBaixaDuplicataReceber("Valor Pago Cheque Vista").Value > 0 Then
                i = Len(Format(rsBaixaDuplicataReceber("Valor Pago Cheque Vista").Value, "#,###,##0.00"))
                Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor Pago Cheque Vista").Value, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                
                xLinha = "|       |                                                |        |            |"
                
                xTamanhoFormaPagamento = Len(FORMA_PAGAMENTO_CH_VISTA)
                Mid(xLinha, 68 + TAM_CAMPO_FORMA_PAGAMENTO - xTamanhoFormaPagamento, xTamanhoFormaPagamento) = FORMA_PAGAMENTO_CH_VISTA
                BioImprime "@Printer.Print " & xLinha
                xContadorFormaPagamento = xContadorFormaPagamento + 1
                xLinha = "|       |                                                |        |            |"
            End If
          
            If rsBaixaDuplicataReceber("Valor Pago Cheque Prazo").Value > 0 Then
                i = Len(Format(rsBaixaDuplicataReceber("Valor Pago Cheque Prazo").Value, "#,###,##0.00"))
                Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor Pago Cheque Prazo").Value, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                
                xLinha = "|       |                                                |        |            |"
                
                xTamanhoFormaPagamento = Len(FORMA_PAGAMENTO_CH_PRAZO)
                Mid(xLinha, 68 + TAM_CAMPO_FORMA_PAGAMENTO - xTamanhoFormaPagamento, xTamanhoFormaPagamento) = FORMA_PAGAMENTO_CH_PRAZO
                BioImprime "@Printer.Print " & xLinha
                xContadorFormaPagamento = xContadorFormaPagamento + 1
                xLinha = "|       |                                                |        |            |"
            End If
            
            
            If rsBaixaDuplicataReceber("Valor Pago Banco").Value > 0 Then
                i = Len(Format(rsBaixaDuplicataReceber("Valor Pago Banco").Value, "#,###,##0.00"))
                Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor Pago Banco").Value, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                
                xLinha = "|       |                                                |        |            |"
                
                xTamanhoFormaPagamento = Len(FORMA_PAGAMENTO_BANCO)
                Mid(xLinha, 68 + TAM_CAMPO_FORMA_PAGAMENTO - xTamanhoFormaPagamento, xTamanhoFormaPagamento) = FORMA_PAGAMENTO_BANCO
                BioImprime "@Printer.Print " & xLinha
                xContadorFormaPagamento = xContadorFormaPagamento + 1
                xLinha = "|       |                                                |        |            |"
            End If
            
            If rsBaixaDuplicataReceber("Valor Pago Cartao").Value > 0 Then
                i = Len(Format(rsBaixaDuplicataReceber("Valor Pago Cartao").Value, "#,###,##0.00"))
                Mid(xLinha, 68 + 12 - i, i) = Format(rsBaixaDuplicataReceber("Valor Pago Cartao").Value, "#,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                
                xLinha = "|       |                                                |        |            |"
                
                xTamanhoFormaPagamento = Len(FORMA_PAGAMENTO_CARTAO)
                Mid(xLinha, 68 + TAM_CAMPO_FORMA_PAGAMENTO - xTamanhoFormaPagamento, xTamanhoFormaPagamento) = FORMA_PAGAMENTO_CARTAO
                BioImprime "@Printer.Print " & xLinha
                xContadorFormaPagamento = xContadorFormaPagamento + 1
                xLinha = "|       |                                                |        |            |"
            End If
            
            If xContadorFormaPagamento = 0 Then
                BioImprime "@Printer.Print " & xLinha
            End If
            
            xLinha = "+-------+------------------------------------------------+--------+------------+"
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 3 + IIf(xContadorFormaPagamento > 1, xContadorFormaPagamento, 0)
            xTotalDinheiro = xTotalDinheiro + rsBaixaDuplicataReceber("Valor Pago Dinheiro").Value
            xTotalChVista = xTotalChVista + rsBaixaDuplicataReceber("Valor Pago Cheque Vista").Value
            xTotalChPrazo = xTotalChPrazo + rsBaixaDuplicataReceber("Valor Pago Cheque Prazo").Value
            xTotalBanco = xTotalBanco + rsBaixaDuplicataReceber("Valor Pago Banco").Value
            xTotalCartao = xTotalCartao + rsBaixaDuplicataReceber("Valor Pago Cartao").Value
            xTotal = xTotal + rsBaixaDuplicataReceber("TotalRecebido").Value
            rsBaixaDuplicataReceber.MoveNext
        Loop
        
        If xTotalDinheiro > 0 Then
            xLinha = "|                          *** DUPLICATAS RECEBIDAS EM DINHEIRO   |            |"
            i = Len(Format(xTotalDinheiro, "##,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(xTotalDinheiro, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
        End If
        If xTotalChVista > 0 Then
            xLinha = "|                          *** DUPLICATAS RECEBIDAS CH. A VISTA   |            |"
            i = Len(Format(xTotalChVista, "##,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(xTotalChVista, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
        End If
        If xTotalChPrazo > 0 Then
            xLinha = "|                          *** DUPLICATAS RECEBIDAS CH.PRE-DATADO |            |"
            i = Len(Format(xTotalChPrazo, "##,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(xTotalChPrazo, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
        End If
        If xTotalBanco > 0 Then
            xLinha = "|                          *** DUPLICATAS RECEBIDAS PELO BANCO    |            |"
            i = Len(Format(xTotalBanco, "##,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(xTotalBanco, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
        End If
        If xTotalCartao > 0 Then
            xLinha = "|                          *** DUPLICATAS RECEBIDAS EM CARTÃO     |            |"
            i = Len(Format(xTotalCartao, "##,###,##0.00"))
            Mid(xLinha, 68 + 12 - i, i) = Format(xTotalCartao, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
        End If
        
        xLinha = "|                          *** TOTAL DE DUPLICATAS RECEBIDAS      |            |"
        i = Len(Format(xTotal, "##,###,##0.00"))
        Mid(xLinha, 68 + 12 - i, i) = Format(xTotal, "##,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+-----------------------------------------------------------------+------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 3
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    If rsBaixaDuplicataReceber.State = 1 Then
        rsBaixaDuplicataReceber.Close
    End If
    Set rsBaixaDuplicataReceber = Nothing
End Sub
Private Sub LoopBaixaCheque()
    Dim xLinha As String
    Dim i As Integer
    
    'Verifica movimento
    If rsBaixaCheque.RecordCount > 0 Then
        rsBaixaCheque.MoveFirst
        If lLinha >= 60 Then
            xLinha = "+-------+------------------------------------------------+--------+------------+"
            Mid(xLinha, 11, 21) = " Cerrado Informatica "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpCabBaixaCheque
        'loop movimento de Cheques Baixados
        Do Until rsBaixaCheque.EOF
            If lLinha >= 65 Then
                xLinha = "+-------+-- Cerrado Informatica ------------------+------------+---------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
                ImpCabBaixaCheque
            End If
            
            '                  1         2         3         4         5         6         7         8         9        10        11        12   12
            '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
            xLinha = "|       |                                         |            |               |"
            Mid(xLinha, 3, 6) = rsBaixaCheque("Numero do Cheque").Value
            Mid(xLinha, 11, 40) = rsBaixaCheque("Emitente").Value
            Mid(xLinha, 53, 10) = Format(rsBaixaCheque("Data do Vencimento").Value, "dd/MM/yyyy")
            i = Len(Format(rsBaixaCheque("Valor").Value, "##,###,##0.00"))
            Mid(xLinha, 66 + 13 - i, i) = Format(rsBaixaCheque("Valor").Value, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
            rsBaixaCheque.MoveNext
        Loop
        xLinha = "+-------+-----------------------------------------+------------+---------------+"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|                              ** TOTAL DE CHEQUES BAIXADOS    |               |"
        i = Len(Format(lTotalBaixaCheque, "##,###,##0.00"))
        Mid(xLinha, 66 + 13 - i, i) = Format(lTotalBaixaCheque, "##,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+--------------------------------------------------------------+---------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 3
    End If
    If rsBaixaCheque.State = 1 Then
        rsBaixaCheque.Close
    End If
    Set rsBaixaCheque = Nothing
End Sub
Private Sub LoopBaixaChequeDevolvido()
    Dim xLinha As String
    Dim i As Integer
    
    If rsBaixaChequeDevolvido.RecordCount > 0 Then
        rsBaixaChequeDevolvido.MoveFirst
        If lLinha >= 60 Then
            xLinha = "+-------+------------------------------------------------+--------+------------+"
            Mid(xLinha, 11, 21) = " Cerrado Informatica "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpCabBaixaCheque
        'loop movimento de Cheques Devolvido Baixados
        Do Until rsBaixaChequeDevolvido.EOF
            If lLinha >= 65 Then
                xLinha = "+-------+-- Cerrado Informatica ------------------+------------+---------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
                ImpCabBaixaCheque
            End If
            
            '                  1         2         3         4         5         6         7         8         9        10        11        12   12
            '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
            xLinha = "|       |                                         |            |               |"
            Mid(xLinha, 3, 6) = rsBaixaChequeDevolvido("Numero do Cheque").Value
            Mid(xLinha, 11, 40) = rsBaixaChequeDevolvido("Emitente").Value
            Mid(xLinha, 53, 10) = Format(rsBaixaChequeDevolvido("Data do Vencimento").Value, "dd/MM/yyyy")
            i = Len(Format(rsBaixaChequeDevolvido("Valor").Value, "##,###,##0.00"))
            Mid(xLinha, 66 + 13 - i, i) = Format(rsBaixaChequeDevolvido("Valor").Value, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
            rsBaixaChequeDevolvido.MoveNext
        Loop
        xLinha = "+-------+-----------------------------------------+------------+---------------+"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|                   ** TOTAL DE CHEQUES DEVOLVIDOS BAIXADOS    |               |"
        i = Len(Format(lTotalBaixaChequeDevolvido, "##,###,##0.00"))
        Mid(xLinha, 66 + 13 - i, i) = Format(lTotalBaixaChequeDevolvido, "##,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+--------------------------------------------------------------+---------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 3
    End If
    If rsBaixaChequeDevolvido.State = 1 Then
        rsBaixaChequeDevolvido.Close
    End If
    Set rsBaixaChequeDevolvido = Nothing
End Sub
Private Sub LoopBaixaNotaAbastecimento()
    Dim xLinha As String
    Dim i As Integer
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Baixa_Nota_Abastecimento.[Codigo do Cliente], Baixa_Nota_Abastecimento.[Data do Abastecimento], Baixa_Nota_Abastecimento.[Numero da Nota],"
    lSQL = lSQL & "       Sum(Baixa_Nota_Abastecimento.[Valor Total]) AS Total, Cliente.[Razao Social] as NomeCliente, Sum(Round([Valor Desconto Unitario] * Quantidade,2)) AS Desconto"
    lSQL = lSQL & "  FROM Baixa_Nota_Abastecimento, Cliente"
    lSQL = lSQL & " WHERE Baixa_Nota_Abastecimento.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Baixa_Nota_Abastecimento.[Data do Abastecimento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Baixa_Nota_Abastecimento.[Data do Abastecimento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Baixa_Nota_Abastecimento.Periodo >= " & preparaTexto(Val(cbo_periodo_i.Text))
    lSQL = lSQL & "   AND Baixa_Nota_Abastecimento.Periodo <= " & preparaTexto(Val(cbo_periodo_f.Text))
    If Val(cboTipoCaixa.Text) > 0 Then
        lSQL = lSQL & "   AND Baixa_Nota_Abastecimento.[Tipo do Movimento] = " & preparaTexto(Val(cboTipoCaixa.Text))
    End If
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Baixa_Nota_Abastecimento.[Data do Abastecimento], 103) + CONVERT(VARCHAR, Baixa_Nota_Abastecimento.Periodo) IN " & lDataPeriodo
    End If
    lSQL = lSQL & "   AND Cliente.Codigo = Baixa_Nota_Abastecimento.[Codigo do Cliente]"
    lSQL = lSQL & " GROUP BY Cliente.[Razao Social], [Data do Abastecimento], [Numero da Nota], [Codigo do Cliente]"
    lSQL = lSQL & " ORDER BY NomeCliente, [Data do Abastecimento], [Numero da Nota]"
    'Abre RecordSet
    Set rsNotaAbastecimento = New ADODB.Recordset
    Set rsNotaAbastecimento = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsNotaAbastecimento.RecordCount > 0 Then
        'loop movimento de notas de abastecimento
        If rsNotaAbastecimento.RecordCount > 0 Then
            Do Until rsNotaAbastecimento.EOF
                
                'If lLinha >= 55 Then
                If lLinha >= 65 Then
                    xLinha = "+-------+-- Cerrado Informatica ------------------+------------+---------------+"
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                    ImpCabNotaAbastecimento
                End If
                
                '                  1         2         3         4         5         6         7         8         9        10        11        12   12
                '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
                xLinha = "|       |                                         |            |               |"


                i = Len(Format(rsNotaAbastecimento("Codigo do Cliente").Value, "##,##0"))
                Mid(xLinha, 2 + 6 - i, i) = Format(rsNotaAbastecimento("Codigo do Cliente").Value, "##,##0")
                Mid(xLinha, 10, 40) = "BX " & Mid(rsNotaAbastecimento("NomeCliente").Value, 1, 37)
                i = Len(Format(rsNotaAbastecimento("Numero da Nota").Value, "######,##0"))
                Mid(xLinha, 52 + 10 - i, i) = Format(rsNotaAbastecimento("Numero da Nota").Value, "######,##0")
                i = Len(Format(rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value, "##,###,##0.00"))
                Mid(xLinha, 66 + 13 - i, i) = Format(rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value, "##,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
                lTotalNota = lTotalNota + rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value
                rsNotaAbastecimento.MoveNext
            Loop
        End If
    End If
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|       |                                         |   ** TOTAL |               |"
    i = Len(Format(lTotalNota, "##,###,##0.00"))
    Mid(xLinha, 66 + 13 - i, i) = Format(lTotalNota, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 3
    If rsNotaAbastecimento.State = 1 Then
        rsNotaAbastecimento.Close
    End If
    Set rsNotaAbastecimento = Nothing
End Sub
Private Sub CalculaAfericaoPorBico()
    Dim xBico As Integer
    Dim rsMovAfericao As New ADODB.Recordset
    
    lSQL = ""
    lSQL = lSQL & "   SELECT [Codigo da Bomba], Quantidade, [Valor Total], [Tipo de Combustivel]"
    lSQL = lSQL & "     FROM Movimento_Afericao"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "      AND Data <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "      AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] >= " & Val(cboIlhaI.Text)
    lSQL = lSQL & "      AND [Numero da Ilha] <= " & Val(cboIlhaF.Text)
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & lDataPeriodo
    End If
    lSQL = lSQL & " ORDER BY [Codigo da Bomba], Data, Periodo"
    
    'Abre RecordSet
    Set rsMovAfericao = New ADODB.Recordset
    Set rsMovAfericao = Conectar.RsConexao(lSQL)
    If rsMovAfericao.RecordCount > 0 Then
        rsMovAfericao.MoveFirst
        Do Until rsMovAfericao.EOF
            xBico = rsMovAfericao("Codigo da Bomba").Value
            lBombaAfericaoLitros(xBico) = lBombaAfericaoLitros(xBico) + rsMovAfericao("Quantidade").Value
            lBombaAfericaoValor(xBico) = lBombaAfericaoValor(xBico) + rsMovAfericao("Valor Total").Value
            lBombaLitros(xBico) = lBombaLitros(xBico) - rsMovAfericao("Quantidade").Value
            lBombaValorTotal(xBico) = lBombaValorTotal(xBico) - rsMovAfericao("Valor Total").Value
            
            'Totaliza Por Combustível
            lBombaTotalLitros = lBombaTotalLitros - rsMovAfericao("Quantidade").Value
            lBombaTotalValor = lBombaTotalValor - rsMovAfericao("Valor Total").Value
            Select Case Trim(rsMovAfericao![Tipo de Combustivel])
                Case "A"
                    lBombaLitrosA = lBombaLitrosA - rsMovAfericao("Quantidade").Value
                    lBombaValorA = lBombaValorA - rsMovAfericao("Valor Total").Value
                Case "AA"
                    lBombaLitrosAA = lBombaLitrosAA - rsMovAfericao("Quantidade").Value
                    lBombaValorAA = lBombaValorAA - rsMovAfericao("Valor Total").Value
                Case "D"
                    lBombaLitrosD = lBombaLitrosD - rsMovAfericao("Quantidade").Value
                    lBombaValorD = lBombaValorD - rsMovAfericao("Valor Total").Value
                Case "DA"
                    lBombaLitrosDA = lBombaLitrosDA - rsMovAfericao("Quantidade").Value
                    lBombaValorDA = lBombaValorDA - rsMovAfericao("Valor Total").Value
                Case "G"
                    lBombaLitrosG = lBombaLitrosG - rsMovAfericao("Quantidade").Value
                    lBombaValorG = lBombaValorG - rsMovAfericao("Valor Total").Value
                Case "GA"
                    lBombaLitrosGA = lBombaLitrosGA - rsMovAfericao("Quantidade").Value
                    lBombaValorGA = lBombaValorGA - rsMovAfericao("Valor Total").Value
            End Select
            rsMovAfericao.MoveNext
        Loop
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
End Sub
'Private Sub CalculaAfericaoPorBico_PA()
'    Dim xBico As Integer
'    Dim xValorUnitario As Currency
'    Dim i As Integer
'
'    lVendaParcialCombustivel = False
'    'loop movimento de aferição
'    lUltimoBico = MovimentoBomba.UltimoBicoComMovimento(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
'    For xBico = 1 To lUltimoBico
'        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
'            lBombaLitros(xBico) = MovimentoBomba.TotalLitrosBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
'            lBombaValorTotal(xBico) = MovimentoBomba.TotalValorBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
'        End If
'
'            lBombaAfericaoLitros(xBico) = lBombaAfericaoLitros(xBico) + rsMovAfericao("Quantidade").Value
'            lBombaAfericaoValor(xBico) = lBombaAfericaoValor(xBico) + rsMovAfericao("Valor Total").Value
'            lBombaLitros(xBico) = lBombaLitros(xBico) - rsMovAfericao("Quantidade").Value
'            lBombaValorTotal(xBico) = lBombaValorTotal(xBico) - rsMovAfericao("Valor Total").Value
'
'
'
'    Next
'End Sub
Private Sub LoopConciliacaoCartao(ByVal pCodigoCartao As Integer)
    Dim rsConciliacaoCartaoPendente As New ADODB.Recordset
    Dim xLinha As String
    Dim i As Integer
    
    lSQL = ""
    lSQL = lSQL & "SELECT Valor, NSU, [Texto da Pendencia]"
    lSQL = lSQL & "     FROM ConciliacaoCartaoPendente"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Tipo de Conciliacao] = " & preparaTexto("V")
    lSQL = lSQL & "      AND [Data de Emissao] >= " & preparaData(txtDataInicial.Text)
    lSQL = lSQL & "      AND [Data de Emissao] <= " & preparaData(txtDataFinal.Text)
    'lSQL = lSQL & "      AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    'lSQL = lSQL & "      AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    If pCodigoCartao > 0 Then
        lSQL = lSQL & "  AND [Codigo do Cartao] = " & pCodigoCartao
    End If
    lSQL = lSQL & " ORDER BY [Data de Emissao], Periodo, [Numero do Lancamento]"
    'Abre RecordSet
    Set rsConciliacaoCartaoPendente = Conectar.RsConexao(lSQL)
    
    If rsConciliacaoCartaoPendente.RecordCount > 0 Then
        lExistePendenciaCartao = True
        If pCodigoCartao > 0 Then
            rsConciliacaoCartaoPendente.MoveFirst
            Do Until rsConciliacaoCartaoPendente.EOF
            
                '                  1         2         3         4         5         6         7         8         9        10        11        12   12
                '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
                xLinha = "|                                                                              |"
            
                Mid(xLinha, 3, 40) = Mid(rsConciliacaoCartaoPendente("Texto da Pendencia").Value, 1, 50)
                
                i = Len(Format(rsConciliacaoCartaoPendente("Valor").Value, "#,###,##0.00"))
                Mid(xLinha, 52 + 12 - i, i) = Format(rsConciliacaoCartaoPendente("Valor").Value, "#,###,##0.00")
    
                Mid(xLinha, 67, 12) = "NSU: " & rsConciliacaoCartaoPendente("NSU").Value
                
                BioImprime "@@Printer.FontBold = True"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.FontBold = False"
                lLinha = lLinha + 1
                
                rsConciliacaoCartaoPendente.MoveNext
            Loop
        End If
    End If
    Set rsConciliacaoCartaoPendente = Nothing
End Sub
Private Sub LoopMovimentoBomba()
    Dim xBico As Integer
    Dim xValorUnitario As Currency
    Dim i As Integer
    
    lVendaParcialCombustivel = False
    lBombaLitrosA = 0
    lBombaValorA = 0
    lBombaLitrosAA = 0
    lBombaValorAA = 0
    lBombaLitrosD = 0
    lBombaValorD = 0
    lBombaLitrosDA = 0
    lBombaValorDA = 0
    lBombaLitrosG = 0
    lBombaValorG = 0
    lBombaLitrosGA = 0
    lBombaValorGA = 0
    


    'loop movimento das bombas
    lUltimoBico = MovimentoBomba.UltimoBicoComMovimento(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    For xBico = 1 To lUltimoBico
        'Le apenas para utilizar o tipo de combustivel do bico
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            rsDataPeriodo.MoveFirst
            If MovimentoBomba.LocalizarPrimeiroPeriodoBico(g_empresa, CDate(txtDataInicial.Text), xBico, Val(Mid(rsDataPeriodo("DataPeriodo").Value, 11, 1))) Then
                lBombaAbertura(xBico) = MovimentoBomba.Abertura
            End If
            rsDataPeriodo.MoveLast
            If MovimentoBomba.LocalizarUltimoPeriodoBico(g_empresa, CDate(txtDataFinal.Text), xBico, Val(Mid(rsDataPeriodo("DataPeriodo").Value, 11, 1))) Then
                lBombaEncerrante(xBico) = MovimentoBomba.Encerrante
            End If
            lBombaLitros(xBico) = lBombaLitros(xBico) + MovimentoBomba.TotalLitrosBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
            lBombaValorTotal(xBico) = lBombaValorTotal(xBico) + MovimentoBomba.TotalValorBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
            
            lBombaTotalDesconto = lBombaTotalDesconto + MovimentoBomba.TotalDescontoBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
            lBombaTotalAcrescimo = lBombaTotalAcrescimo + MovimentoBomba.TotalAcrescimoBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
            
        Else
            If MovimentoBomba.LocalizarPrimeiroPeriodoBico(g_empresa, CDate(txtDataInicial.Text), xBico, 0) Then
            End If
            lBombaAbertura(xBico) = MovimentoBomba.AberturaBicoDataPeriodo(g_empresa, CDate(txtDataInicial.Text), xBico, Val(cbo_periodo_i.Text))
            lBombaEncerrante(xBico) = MovimentoBomba.EncerranteBicoDataPeriodo(g_empresa, CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_f.Text))
            lBombaLitros(xBico) = lBombaLitros(xBico) + MovimentoBomba.TotalLitrosBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
            lBombaValorTotal(xBico) = lBombaValorTotal(xBico) + MovimentoBomba.TotalValorBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
            
            lBombaTotalDesconto = lBombaTotalDesconto + MovimentoBomba.TotalDescontoBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
            lBombaTotalAcrescimo = lBombaTotalAcrescimo + MovimentoBomba.TotalAcrescimoBicoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), xBico, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "")
        End If
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
    If g_automacao And lUltimoBico = 0 Then
        lUltimoBico = Configuracao.QuantidadeBico
        lVendaParcialCombustivel = True
        For xBico = 1 To lUltimoBico
            'Le apenas para utilizar o tipo de combustivel do bico
            If MovimentoBomba.LocalizarBicoAnteriorData(g_empresa, CDate(txtDataInicial.Text), Val(cbo_periodo_f.Text), xBico) Then
                If EncerranteAtual.LocalizarCodigo(g_empresa, xBico) Then
                    xValorUnitario = 1
                    If Bomba.LocalizarCodigo(g_empresa, xBico) Then
                        xValorUnitario = Bomba.PrecoVenda
                        lBombaTipoPreco(xBico) = Bomba.TipoPreco
                        lBombaCombustivel(xBico) = Bomba.TipoCombustivel
                    End If
                    lBombaAbertura(xBico) = MovimentoBomba.Encerrante
                    lBombaEncerrante(xBico) = EncerranteAtual.Encerrante
                    
                    For i = 1 To 10
                        If (lBombaAbertura(xBico) - lBombaEncerrante(xBico)) > 800000 Then
                            lBombaEncerrante(xBico) = lBombaEncerrante(xBico) + 1000000
                        End If
                    Next
                    lBombaLitros(xBico) = lBombaEncerrante(xBico) - lBombaAbertura(xBico)
                    lBombaValorTotal(xBico) = Format(lBombaLitros(xBico) * xValorUnitario, "0000000000.00")
                    
                    'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                    lBombaValorTotal(xBico) = lBombaValorTotal(xBico) + MovimentoBomba.TotalAcrescimo
                    lBombaValorTotal(xBico) = lBombaValorTotal(xBico) - MovimentoBomba.TotalDesconto
                    
                    lBombaTotalAcrescimo = lBombaTotalAcrescimo + MovimentoBomba.TotalAcrescimo
                    lBombaTotalDesconto = lBombaTotalDesconto + MovimentoBomba.TotalDesconto
                    
                    Select Case Trim(lBombaCombustivel(xBico))
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
                End If
            End If
        Next
    End If
    
    'Verifica Continuidade de Encerrante
'    For xBico = 1 To lUltimoBico
'        If MovimentoBomba.LocalizarBicoAnteriorData(g_empresa, (CDate(txtdatainicial.Text) - 1), Val(cbo_periodo_i.Text), xBico) Then
'            If lBombaAbertura(xBico) <> MovimentoBomba.Encerrante Then
'                lBombaMensagem(xBico) = "A abertura do bico " & (xBico) & " deveria ser " & Format(MovimentoBomba.Encerrante, "##,###,##0.00")
'            End If
'        Else
'            lBombaMensagem(xBico) = "Não foi possível localizar encerrante anterior do bico " & (xBico)
'        End If
'    Next
End Sub
Private Sub LoopMovimentoBombaRS()
    Dim xBico As Integer
    Dim xValorUnitario As Currency
    Dim i As Integer
    
    lVendaParcialCombustivel = False
    lUltimoBico = MovimentoBomba.UltimoBicoComMovimento(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    
    'loop movimento das bombas
    lSQL = ""
    lSQL = lSQL & "SELECT Data, Periodo, [Codigo da Bomba], Abertura, Encerrante, "
    lSQL = lSQL & "       [Quantidade da Saida], [Preco de Custo], [Preco de Venda], [Tipo de Combustivel], "
    lSQL = lSQL & "       [Numero do Tanque], [Total Venda Bruto],[Total Desconto],[Total Acrescimo]"
    lSQL = lSQL & "  FROM " & MovimentoBomba.NomeTabela
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "   AND [Numero da Ilha] >= " & Val(cboIlhaI.Text)
    lSQL = lSQL & "   AND [Numero da Ilha] <= " & Val(cboIlhaF.Text)
    lSQL = lSQL & " ORDER BY Data, Periodo, SubCaixa, [Codigo da Bomba]"
    'Abre RecordSet
    Set rsMovimentoBomba = New ADODB.Recordset
    Set rsMovimentoBomba = Conectar.RsConexao(lSQL)
    If rsMovimentoBomba.RecordCount > 0 Then
        'Loop por bico
        For xBico = 1 To lUltimoBico
            'Define tipo de preço, (À Vista ou À Prazo)
            If Bomba.LocalizarCodigo(g_empresa, xBico) Then
                lBombaTipoPreco(xBico) = Bomba.TipoPreco
            End If
            
            'Filtra por bico
            rsMovimentoBomba.Filter = "[Codigo da Bomba] = " & xBico
            
            If rsMovimentoBomba.RecordCount > 0 Then
                'Busca Abertura do Bico
                lBombaAbertura(xBico) = rsMovimentoBomba!Abertura
                                
                Do Until rsMovimentoBomba.EOF
                
                    'Busca Encerrante do Bico
                    lBombaEncerrante(xBico) = rsMovimentoBomba!Encerrante
                    
                    'Define Tipo de Combustível do Bico, Pelo Último Movimento do Dia
                    lBombaCombustivel(xBico) = rsMovimentoBomba![Tipo de Combustivel]

                    'Totaliza Litros Vendidos (Bico)
                    lBombaLitros(xBico) = lBombaLitros(xBico) + rsMovimentoBomba![Quantidade da Saida]
                    
                    'Totaliza Valor Vendidos  (Bico)
                    lBombaValorTotal(xBico) = lBombaValorTotal(xBico) + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                    
                    'Totaliza Por Combustível
                    Select Case Trim(rsMovimentoBomba![Tipo de Combustivel])
                        Case "A"
                            lBombaLitrosA = lBombaLitrosA + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorA = lBombaValorA + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorA = lBombaValorA + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorA = lBombaValorA - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosA = lBombaLitrosA + lBombaLitros(xBico)
                            'lBombaValorA = lBombaValorA + lBombaValorTotal(xBico)
                        Case "AA"
                            lBombaLitrosAA = lBombaLitrosAA + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorAA = lBombaValorAA + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorAA = lBombaValorAA + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorAA = lBombaValorAA - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosAA = lBombaLitrosAA + lBombaLitros(xBico)
                            'lBombaValorAA = lBombaValorAA + lBombaValorTotal(xBico)
                        Case "D"
                            lBombaLitrosD = lBombaLitrosD + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorD = lBombaValorD + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorD = lBombaValorD + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorD = lBombaValorD - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosD = lBombaLitrosD + lBombaLitros(xBico)
                            'lBombaValorD = lBombaValorD + lBombaValorTotal(xBico)
                        Case "DA"
                            lBombaLitrosDA = lBombaLitrosDA + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorDA = lBombaValorDA + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorDA = lBombaValorDA + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorDA = lBombaValorDA - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosDA = lBombaLitrosDA + lBombaLitros(xBico)
                            'lBombaValorDA = lBombaValorDA + lBombaValorTotal(xBico)
                        Case "G"
                            lBombaLitrosG = lBombaLitrosG + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorG = lBombaValorG + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorG = lBombaValorG + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorG = lBombaValorG - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosG = lBombaLitrosG + lBombaLitros(xBico)
                            'lBombaValorG = lBombaValorG + lBombaValorTotal(xBico)
                        Case "GA"
                            lBombaLitrosGA = lBombaLitrosGA + rsMovimentoBomba![Quantidade da Saida]
                            lBombaValorGA = lBombaValorGA + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                            
                            'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                            lBombaValorGA = lBombaValorGA + rsMovimentoBomba![Total Acrescimo]
                            lBombaValorGA = lBombaValorGA - rsMovimentoBomba![Total Desconto]
                            
                            'lBombaLitrosGA = lBombaLitrosGA + lBombaLitros(xBico)
                            'lBombaValorGA = lBombaValorGA + lBombaValorTotal(xBico)
                    End Select
                    
                    'Totaliza Litros Vendidos (Geral)
                    lBombaTotalLitros = lBombaTotalLitros + rsMovimentoBomba![Quantidade da Saida]
                    'Totaliza Valor Vendidos  (Geral)
                    lBombaTotalValor = lBombaTotalValor + Round(rsMovimentoBomba![Quantidade da Saida] * rsMovimentoBomba![Preco de Venda], 2)
                    
                    'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                    lBombaTotalValor = lBombaTotalValor + rsMovimentoBomba![Total Acrescimo]
                    lBombaTotalValor = lBombaTotalValor - rsMovimentoBomba![Total Desconto]
                    
                    lBombaTotalAcrescimo = lBombaTotalAcrescimo + rsMovimentoBomba![Total Acrescimo]
                    lBombaTotalDesconto = lBombaTotalDesconto + rsMovimentoBomba![Total Desconto]
                            
                    
                    rsMovimentoBomba.MoveNext
                Loop
            End If
        Next
    End If
    
    'Busca Venda Parcial da Automação
    If g_automacao And lUltimoBico = 0 Then
        lUltimoBico = Configuracao.QuantidadeBico
        lVendaParcialCombustivel = True
        For xBico = 1 To lUltimoBico
            'Le apenas para utilizar o tipo de combustivel do bico
            If MovimentoBomba.LocalizarBicoAnteriorData(g_empresa, CDate(txtDataInicial.Text), Val(cbo_periodo_f.Text), xBico) Then
                If EncerranteAtual.LocalizarCodigo(g_empresa, xBico) Then
                    xValorUnitario = 1
                    If Bomba.LocalizarCodigo(g_empresa, xBico) Then
                        xValorUnitario = Bomba.PrecoVenda
                        lBombaTipoPreco(xBico) = Bomba.TipoPreco
                        lBombaCombustivel(xBico) = Bomba.TipoCombustivel
                    End If
                    lBombaAbertura(xBico) = MovimentoBomba.Encerrante
                    lBombaEncerrante(xBico) = EncerranteAtual.Encerrante
                    
                    For i = 1 To 10
                        If (lBombaAbertura(xBico) - lBombaEncerrante(xBico)) > 800000 Then
                            lBombaEncerrante(xBico) = lBombaEncerrante(xBico) + 1000000
                        End If
                    Next
                    lBombaLitros(xBico) = lBombaEncerrante(xBico) - lBombaAbertura(xBico)
                    lBombaValorTotal(xBico) = Format(lBombaLitros(xBico) * xValorUnitario, "0000000000.00")
                    
                    'ALTERÇÃO FEITA PARA UTILIZACAO DE DOIS PRECOS NA AUTOMAÇÃO
                    lBombaValorTotal(xBico) = lBombaValorTotal(xBico) + MovimentoBomba.TotalAcrescimo
                    lBombaValorTotal(xBico) = lBombaValorTotal(xBico) - MovimentoBomba.TotalDesconto
                    
                    lBombaTotalAcrescimo = lBombaTotalAcrescimo + MovimentoBomba.TotalAcrescimo
                    lBombaTotalDesconto = lBombaTotalDesconto + MovimentoBomba.TotalDesconto
                    
                    Select Case Trim(lBombaCombustivel(xBico))
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
                End If
            End If
        Next
    End If
End Sub
Private Sub LoopMovimentoCaixaPista()
    Dim xLinha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xString As String
    Dim xConciliacaoCartaoDivergente As Boolean
    
    ImpCabCaixaPista
    'loop movimento de caixa de pista
    If rsMovimentoCaixaPista.RecordCount > 0 Then
        Do Until rsMovimentoCaixaPista.EOF
            xConciliacaoCartaoDivergente = False
            
            'If lLinha >= 55 Then
            If lLinha >= 65 Then
                'xLinha = "+----+-- Cerrado Informatica -----------------+------------+------+------------+"
                xLinha = "+-----+-- Cerrado Informatica -------------------+--------------+--------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
                ImpCabCaixaPista
            End If
            
            '                  1         2         3         4         5         6         7         8         9        10        11        12   12
            '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
            xLinha = "|     |                                          |              |              |"
            i = Len(Format(rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value, "##0"))
            Mid(xLinha, 3 + 3 - i, i) = Format(rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value, "##0")
            Mid(xLinha, 9, 40) = rsMovimentoCaixaPista("NomePlano").Value
            If LancamentoPadrao.LocalizarCodigo(g_empresa, rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value) Then
                If rsMovimentoCaixaPista("ValorEntrada").Value > 0 Then
                    If LancamentoPadrao.NomePlanoCredito = False Then
                        Mid(xLinha, 9, 40) = Space(40)
                        Mid(xLinha, 9, 40) = LancamentoPadrao.Nome
                    End If
                End If
                If rsMovimentoCaixaPista("ValorSaida").Value > 0 Then
                    If LancamentoPadrao.NomePlanoDebito = False Then
                        Mid(xLinha, 9, 40) = Space(40)
                        Mid(xLinha, 9, 40) = LancamentoPadrao.Nome
                    End If
                End If
                ' Cartão de Crédito
                If rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value = 2 Then
                    Dim xNomeCartao As String
                    xNomeCartao = rsMovimentoCaixaPista("NomePlano").Value
                    xNomeCartao = Mid(rsMovimentoCaixaPista("NomePlano").Value, 8, Len(xNomeCartao))
                    xNomeCartao = Replace(xNomeCartao, "CRÉDITO", "CREDITO")
                    xNomeCartao = Replace(xNomeCartao, "DÉBITO", "DEBITO")
                    If CartaoCredito.LocalizarNome(xNomeCartao) = True Then
                        If ConciliacaoCartao.LocalizarCodigo(g_empresa, "V", CDate(txtDataInicial.Text), CartaoCredito.Codigo) = True Then
                            If ConciliacaoCartao.TotalBruto = rsMovimentoCaixaPista("ValorEntrada").Value Then
                                Mid(xLinha, 67, 12) = "Conciliado  "
                            Else
                                Mid(xLinha, 67, 12) = "Valor Difere"
                                xConciliacaoCartaoDivergente = True
                            End If
                        Else
                            Mid(xLinha, 66, 14) = "NÃO Conciliado"
                        End If
                    Else
                        Mid(xLinha, 66, 14) = "Cartão Inexist"
                    End If
                End If
            End If
            If rsMovimentoCaixaPista("ValorEntrada").Value > 0 Then
                i = Len(Format(rsMovimentoCaixaPista("ValorEntrada").Value, "#,###,##0.00"))
                Mid(xLinha, 52 + 12 - i, i) = Format(rsMovimentoCaixaPista("ValorEntrada").Value, "#,###,##0.00")
            End If
            If rsMovimentoCaixaPista("ValorSaida").Value > 0 Then
                i = Len(Format(rsMovimentoCaixaPista("ValorSaida").Value, "#,###,##0.00"))
                Mid(xLinha, 67 + 12 - i, i) = Format(rsMovimentoCaixaPista("ValorSaida").Value, "#,###,##0.00")
            End If
            If rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value <> 4 Then
                If rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value = 7 Then 'COMBUSTIVEIS
                    Mid(xLinha, 66, 14) = "              "
                    Dim xValorCombustiveis As Currency
                    xValorCombustiveis = rsMovimentoCaixaPista("ValorSaida").Value - lBombaValorAfericao + lBombaTotalAcrescimo - lBombaTotalDesconto
                    'i = Len(Format(rsMovimentoCaixaPista("ValorSaida").Value - lBombaValorAfericao, "#,###,##0.00"))
                    i = Len(Format(xValorCombustiveis, "#,###,##0.00"))
                    'Mid(xLinha, 67 + 12 - i, i) = Format(rsMovimentoCaixaPista("ValorSaida").Value - lBombaValorAfericao, "#,###,##0.00")
                    Mid(xLinha, 67 + 12 - i, i) = Format(xValorCombustiveis, "#,###,##0.00")
                    lTotalSaida = lTotalSaida - lBombaValorAfericao + lBombaTotalAcrescimo - lBombaTotalDesconto
                End If
                
                If xConciliacaoCartaoDivergente = True Then
                    BioImprime "@@Printer.FontBold = True"
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.FontBold = False"
                Else
                    BioImprime "@Printer.Print " & xLinha
                End If
                lLinha = lLinha + 1
                lTotalEntrada = lTotalEntrada + rsMovimentoCaixaPista("ValorEntrada").Value
                lTotalSaida = lTotalSaida + rsMovimentoCaixaPista("ValorSaida").Value
                If rsMovimentoCaixaPista("NomePlano").Value = "DINHEIRO" Or rsMovimentoCaixaPista("NomePlano").Value = "CHEQUE À VISTA" Then
                    lTotalVista = lTotalVista + rsMovimentoCaixaPista("ValorEntrada").Value
                Else
                    '2 = Cartão
                    '3 = Nota Abastecimento
                    '5 = Cheque Pré-Datado
                    'Obs: Testar Apenas "Cheque*" pois na contabilidade não existe Cheque Pré-Datado
                    If rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value = 2 Or rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value = 3 Or rsMovimentoCaixaPista("Codigo do Lancamento Padrao").Value = 5 Then
                        If rsMovimentoCaixaPista("NomePlano").Value Like "CARTAO*" Or rsMovimentoCaixaPista("NomePlano").Value = "NOTAS DE ABASTECIMENTO" Or rsMovimentoCaixaPista("NomePlano").Value Like "CHEQUE*" Then
                            lTotalPrazo = lTotalPrazo + rsMovimentoCaixaPista("ValorEntrada").Value
                        End If
                    End If
                End If
            End If
            'lTotalVista = 0
            'lTotalPrazo = 0
            If xConciliacaoCartaoDivergente = True And CBool(chkImprimeConciliacao.Value) = True Then
                Call LoopConciliacaoCartao(CartaoCredito.Codigo)
            End If
            rsMovimentoCaixaPista.MoveNext
        Loop
    End If
    
    
    xLinha = "+-----+------------------------------------------+--------------+--------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    For i2 = 1 To 4
        If lLinha >= 65 Then
            xLinha = "+-----+-- Cerrado Informatica -------------------+--------------+--------------+"
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCabCaixaPista
        End If
        xLinha = "| *** |                                          |              |              |"
        If i2 = 1 Then
            Mid(xLinha, 9, 40) = "VENDA A VISTA"
            i = Len(Format(lTotalVista, "#,###,##0.00"))
            Mid(xLinha, 52 + 12 - i, i) = Format(lTotalVista, "#,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
        ElseIf i2 = 2 Then
            Mid(xLinha, 9, 40) = "VENDA A PRAZO"
            i = Len(Format(lTotalPrazo, "#,###,##0.00"))
            Mid(xLinha, 52 + 12 - i, i) = Format(lTotalPrazo, "#,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
        ElseIf i2 = 3 Then
            Mid(xLinha, 9, 40) = "RECEBIMENTO DE DUPLICATAS (NOTAS)"
            i = Len(Format(lTotalDuplicataRecebida, "#,###,##0.00"))
            Mid(xLinha, 52 + 12 - i, i) = Format(lTotalDuplicataRecebida, "#,###,##0.00")
            If lTotalDuplicataRecebida > 0 Then
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
            End If
        ElseIf i2 = 4 Then
            Mid(xLinha, 9, 40) = "CHEQUES BAIXADOS"
            i = Len(Format(lTotalBaixaCheque, "#,###,##0.00"))
            Mid(xLinha, 52 + 12 - i, i) = Format(lTotalBaixaCheque, "#,###,##0.00")
            If lTotalBaixaCheque > 0 Then
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
            End If
        End If
    Next
    
    
    
    If lVendaParcialCombustivel Then
        xLinha = "| *** | VENDA PARCIAL DE COMBUSTIVEIS            |              |              |"
        i = Len(Format(lBombaTotalValor, "#,###,##0.00"))
        Mid(xLinha, 67 + 12 - i, i) = Format(lBombaTotalValor, "#,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        lTotalSaida = lTotalSaida + lBombaTotalValor
    End If
    
    'Caso seja 58 ou mais nao tera espaço para imprimir o resumo de combustivel na mesma folha
    If lLinha >= 58 Then
        xLinha = "+-----+-- Cerrado Informatica -------------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
        'ImpCabCaixaPista
    End If
    
    If lTipoMovimento = 1 Then
        xLinha = "+-----+------------------------------------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        xLinha = "|                                                |              |              |"
    Else
        xLinha = "+-----+-----+-------------+-------------+--------+--------------+--------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        xLinha = "|COMBUSTÍVEL|    LITROS   |    VALOR    |        |              |              |"
    End If
    i = Len(Format(lTotalEntrada, "#,###,##0.00"))
    Mid(xLinha, 52 + 12 - i, i) = Format(lTotalEntrada, "#,###,##0.00")
    i = Len(Format(lTotalSaida, "#,###,##0.00"))
    Mid(xLinha, 67 + 12 - i, i) = Format(lTotalSaida, "#,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    If lTipoMovimento = 1 Then
        xLinha = "|                                       +--------+--------------+--------------+"
    Else
        xLinha = "+-----------+-------------+-------------+--------+--------------+--------------+"
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    rsMovimentoCaixaPista.Close
    Set rsMovimentoCaixaPista = Nothing
End Sub
Private Sub LoopMovimentoLubrificante()
    'Prepara SQL
    lSQL = ""
    If chkImprimeLubrificante.Value = 1 Then
        lSQL = lSQL & "   SELECT Movimento_Lubrificante.[Codigo do Produto2], SUM(Movimento_Lubrificante.Quantidade) AS Quantidade, SUM(Movimento_Lubrificante.[Valor Total]) AS [Valor Total], Produto.Nome"
        lSQL = lSQL & "     FROM Movimento_Lubrificante, Produto"
        lSQL = lSQL & "    WHERE Movimento_Lubrificante.Empresa = " & g_empresa
        lSQL = lSQL & "      AND Movimento_Lubrificante.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "      AND Movimento_Lubrificante.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
        lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
        lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
        ''lSQL = lSQL & "      AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 2
        If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, AberturaCaixa.CodigoFuncionario) Then
            lSQL = lSQL & "      AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 3
        Else
            lSQL = lSQL & "      AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 2
        End If
        If lTipoMovimento > 0 Then
            lSQL = lSQL & "      AND Movimento_Lubrificante.[Tipo do Movimento] = " & lTipoMovimento
        End If
        lSQL = lSQL & "      AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            lSQL = lSQL & "    AND CONVERT(VARCHAR, Movimento_Lubrificante.Data, 103) + CONVERT(VARCHAR, Movimento_Lubrificante.Periodo) IN " & lDataPeriodo
        End If
        lSQL = lSQL & " GROUP BY Produto.Nome, Movimento_Lubrificante.[Codigo do Produto2]"
    Else
        lSQL = ""
        lSQL = lSQL & "SELECT Grupo.Codigo, Grupo.Nome, SUM(Movimento_Lubrificante.Quantidade) AS Quantidade, SUM(Movimento_Lubrificante.[Valor Total]) AS [Valor Total]"
        lSQL = lSQL & "  FROM Movimento_Lubrificante, Produto, Grupo"
        lSQL = lSQL & " WHERE Movimento_Lubrificante.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Movimento_Lubrificante.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "   AND Movimento_Lubrificante.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "   AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & "   AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
        lSQL = lSQL & "   AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
        lSQL = lSQL & "   AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
        lSQL = lSQL & "   AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 2
        If lTipoMovimento > 0 Then
            lSQL = lSQL & "      AND Movimento_Lubrificante.[Tipo do Movimento] = " & lTipoMovimento
        End If
        lSQL = lSQL & "   AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
        lSQL = lSQL & "   AND Grupo.Codigo = Produto.[Codigo do Grupo]"
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            lSQL = lSQL & "    AND CONVERT(VARCHAR, Movimento_Lubrificante.Data, 103) + CONVERT(VARCHAR, Movimento_Lubrificante.Periodo) IN " & lDataPeriodo
        End If
        lSQL = lSQL & " GROUP BY Grupo.Nome, Grupo.Codigo"
    End If
    
    
    'Abre RecordSet
    Set rsMovLubrificante = New ADODB.Recordset
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        ImpCabLubrificante
        rsMovLubrificante.MoveFirst
        Do Until rsMovLubrificante.EOF
            ImpDetLubrificante
            lTotalLubrificante = lTotalLubrificante
            rsMovLubrificante.MoveNext
        Loop
    End If
End Sub
Private Sub LoopMovimentoDespesaCaixa()
    Dim xLinha As String
    Dim i As Integer
    Dim xTotal As Currency
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Despesa_Caixa.[Codigo do Fornecedor], Movimento_Despesa_Caixa.Valor, Movimento_Despesa_Caixa.[Numero do Documento], "
    lSQL = lSQL & "       Movimento_Despesa_Caixa.[Data do Movimento], Movimento_Despesa_Caixa.Complemento, Fornecedor.Nome"
    lSQL = lSQL & "  FROM Movimento_Despesa_Caixa, Fornecedor"
    lSQL = lSQL & " WHERE Movimento_Despesa_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.[Data do Movimento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.[Data do Movimento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
    lSQL = lSQL & "   AND Movimento_Despesa_Caixa.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Movimento_Despesa_Caixa.[Data do Movimento], 103) + CONVERT(VARCHAR, Movimento_Despesa_Caixa.Periodo) IN " & lDataPeriodo
    End If
'    If Val(cboTipoCaixa.Text) > 0 Then
'        lSQL = lSQL & "   AND Movimento_Despesa_Caixa.[Tipo do Movimento] = " & preparaTexto(Val(cboTipoCaixa.Text))
'    End If
    lSQL = lSQL & "   AND Fornecedor.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Fornecedor.Codigo = Movimento_Despesa_Caixa.[Codigo do Fornecedor]"
    lSQL = lSQL & " ORDER BY [Data do Movimento], [Numero da Ilha], Periodo"
    'Abre RecordSet
    Set rsNotaAbastecimento = New ADODB.Recordset
    Set rsNotaAbastecimento = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsNotaAbastecimento.RecordCount > 0 Then
        'loop movimento de notas de abastecimento
        If rsNotaAbastecimento.RecordCount > 0 Then
            ImpCabDespesaCaixa
            xTotal = 0
            Do Until rsNotaAbastecimento.EOF
                
                'If lLinha >= 55 Then
                If lLinha >= 65 Then
                    xLinha = "+------------------------------------------+---------------+------------+------------+--------------------------------------------------+"
                    Mid(xLinha, 5, 21) = " Cerrado Informatica "
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                    ImpCabDespesaCaixa
                End If
                
                '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
                '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
                xLinha = "|                                          |               |            |            |                                                  |"
                Mid(xLinha, 3, 40) = rsNotaAbastecimento("Nome").Value
                i = Len(Format(rsNotaAbastecimento("Valor").Value, "##,###,##0.00"))
                Mid(xLinha, 46 + 13 - i, i) = Format(rsNotaAbastecimento("Valor").Value, "##,###,##0.00")
                Mid(xLinha, 62, 10) = rsNotaAbastecimento("Numero do Documento").Value
                Mid(xLinha, 75, 10) = Format(rsNotaAbastecimento("Data do Movimento").Value, "dd/mm/yyyy")
                Mid(xLinha, 88, 40) = rsNotaAbastecimento("Complemento").Value
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
                xTotal = xTotal + rsNotaAbastecimento("Valor").Value
                rsNotaAbastecimento.MoveNext
            Loop
            xLinha = "+------------------------------------------+---------------+------------+------------+--------------------------------------------------+"
            BioImprime "@Printer.Print " & xLinha
            xLinha = "|                            ***  TOTAL    |               |                                                                            |"
            i = Len(Format(xTotal, "##,###,##0.00"))
            Mid(xLinha, 46 + 13 - i, i) = Format(xTotal, "##,###,##0.00")
            BioImprime "@Printer.Print " & xLinha
            xLinha = "+------------------------------------------+---------------+------------+------------+--------------------------------------------------+"
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 3
        End If
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    If rsNotaAbastecimento.State = 1 Then
        rsNotaAbastecimento.Close
    End If
    Set rsNotaAbastecimento = Nothing
End Sub
Private Sub LoopMovimentoNotaAbastecimento()
    Dim xLinha As String
    Dim i As Integer
    Dim xDiferenca As Currency
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Nota_Abastecimento.[Codigo do Cliente], Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.[Numero da Nota],"
    lSQL = lSQL & "       Sum(Movimento_Nota_Abastecimento.[Valor Total]) AS Total, Cliente.[Razao Social] as NomeCliente, Sum(Round([Valor Desconto Unitario] * Quantidade,2)) AS Desconto,"
    lSQL = lSQL & "       Sum(Round(Movimento_Nota_Abastecimento.[Valor Unitario] * Quantidade,2)) AS TotalCalculado,"
    lSQL = lSQL & "       Sum(Quantidade) AS QuantidadeCalculada,"
    lSQL = lSQL & "       MAX([Valor Unitario]) AS ValorUnitarioCalculado"
    lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento, Cliente"
    lSQL = lSQL & " WHERE Movimento_Nota_Abastecimento.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Data do Abastecimento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Data do Abastecimento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.Periodo >= " & preparaTexto(Val(cbo_periodo_i.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.Periodo <= " & preparaTexto(Val(cbo_periodo_f.Text))
    If Val(cboTipoCaixa.Text) > 0 Then
        lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Tipo do Movimento] = " & preparaTexto(Val(cboTipoCaixa.Text))
    End If
    lSQL = lSQL & "   AND Cliente.Codigo = Movimento_Nota_Abastecimento.[Codigo do Cliente]"
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Movimento_Nota_Abastecimento.[Data do Abastecimento], 103) + CONVERT(VARCHAR, Movimento_Nota_Abastecimento.Periodo) IN " & lDataPeriodo
    End If
    lSQL = lSQL & " GROUP BY Cliente.[Razao Social], [Data do Abastecimento], [Numero da Nota], [Codigo do Cliente]"
    lSQL = lSQL & " ORDER BY NomeCliente, [Data do Abastecimento], [Numero da Nota]"
    'Abre RecordSet
    Set rsNotaAbastecimento = New ADODB.Recordset
    Set rsNotaAbastecimento = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsNotaAbastecimento.RecordCount > 0 Then
        'loop movimento de notas de abastecimento
        If rsNotaAbastecimento.RecordCount > 0 Then
            If lLinha >= 60 Then
                xLinha = "+-------+-- Cerrado Informatica ------------------+------------+---------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            ImpCabNotaAbastecimento
            Do Until rsNotaAbastecimento.EOF
                
                'If lLinha >= 55 Then
                If lLinha >= 62 Then
                    xLinha = "+-------+-- Cerrado Informatica ------------------+------------+---------------+"
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                    ImpCabNotaAbastecimento
                End If
                
                '                  1         2         3         4         5         6         7         8         9        10        11        12   12
                '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
                xLinha = "|       |                                         |            |               |"


                i = Len(Format(rsNotaAbastecimento("Codigo do Cliente").Value, "##,##0"))
                Mid(xLinha, 2 + 6 - i, i) = Format(rsNotaAbastecimento("Codigo do Cliente").Value, "##,##0")
                Mid(xLinha, 10, 40) = rsNotaAbastecimento("NomeCliente").Value
                i = Len(Format(rsNotaAbastecimento("Numero da Nota").Value, "######,##0"))
                Mid(xLinha, 52 + 10 - i, i) = Format(rsNotaAbastecimento("Numero da Nota").Value, "######,##0")
                i = Len(Format(rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value, "##,###,##0.00"))
                Mid(xLinha, 66 + 13 - i, i) = Format(rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value, "##,###,##0.00")
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
                
                
                lTotalNota = lTotalNota + rsNotaAbastecimento("Total").Value - rsNotaAbastecimento("Desconto").Value
                
                xDiferenca = rsNotaAbastecimento("ValorUnitarioCalculado").Value * rsNotaAbastecimento("QuantidadeCalculada").Value - rsNotaAbastecimento("Total").Value
                If (xDiferenca * -1) > 0.03 Then
                    xLinha = " ** PROBLEMA NA GERAÇÃO DE DESCONTO NA NOTA ACIMA. FAVOR ALTERA-LA. - " & Format(xDiferenca, "##,###,##0.00")
                    BioImprime "@Printer.Print " & xLinha
                    lLinha = lLinha + 1
                End If
                
                
                rsNotaAbastecimento.MoveNext
            Loop
'            xLinha = "+-------+-----------------------------------------+------------+---------------+"
'            BioImprime "@Printer.Print " & xLinha
'            xLinha = "|       |                                         |   ** TOTAL |               |"
'            i = Len(Format(xTotal, "##,###,##0.00"))
'            Mid(xLinha, 66 + 13 - i, i) = Format(xTotal, "##,###,##0.00")
'            BioImprime "@Printer.Print " & xLinha
'            xLinha = "+-------+-----------------------------------------+------------+---------------+"
'            BioImprime "@Printer.Print " & xLinha
'            lLinha = lLinha + 3
        End If
    End If
    If rsNotaAbastecimento.State = 1 Then
        rsNotaAbastecimento.Close
    End If
    Set rsNotaAbastecimento = Nothing
End Sub
Private Sub CalculaAfericao()
    Dim xLitrosAfericao As Currency
    Dim xValorAfericao As Currency
    
    'Alcool
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosA = lBombaLitrosA - xLitrosAfericao
    'lBombaValorA = lBombaValorA - xValorAfericao
    'Alcool Aditivado
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosAA = lBombaLitrosAA - xLitrosAfericao
    'lBombaValorAA = lBombaValorAA - xValorAfericao
    'Diesel
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosD = lBombaLitrosD - xLitrosAfericao
    'lBombaValorD = lBombaValorD - xValorAfericao
    'Diesel Aditivado
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosDA = lBombaLitrosDA - xLitrosAfericao
    'lBombaValorDA = lBombaValorDA - xValorAfericao
    'Gasolina
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosG = lBombaLitrosG - xLitrosAfericao
    'lBombaValorG = lBombaValorG - xValorAfericao
    'Gasolina Aditivada
    xLitrosAfericao = MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", lDataPeriodo)
    xValorAfericao = MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", lDataPeriodo)
    lBombaLitrosAfericao = lBombaLitrosAfericao + xLitrosAfericao
    lBombaValorAfericao = lBombaValorAfericao + xValorAfericao
    'lBombaLitrosGA = lBombaLitrosGA - xLitrosAfericao
    'lBombaValorGA = lBombaValorGA - xValorAfericao
    
    'lBombaTotalLitros = lBombaTotalLitros - lBombaLitrosAfericao
    'lBombaTotalValor = lBombaTotalValor - lBombaValorAfericao
End Sub
Private Sub TotalizaBaixaCheque()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Emitente, [Numero do Cheque], Valor, [Data do Vencimento], [Data do Pagamento], [Periodo do Pagamento]"
    lSQL = lSQL & "  FROM Baixa_Cheque"
    lSQL = lSQL & " WHERE Baixa_Cheque.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Baixa_Cheque.[Data do Pagamento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Baixa_Cheque.[Data do Pagamento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Baixa_Cheque.[Periodo do Pagamento] >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Baixa_Cheque.[Periodo do Pagamento] <= " & Val(cbo_periodo_f.Text)
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Baixa_Cheque.[Data do Pagamento], 103) + CONVERT(VARCHAR, Baixa_Cheque.[Periodo do Pagamento]) IN " & lDataPeriodo
    End If
    lSQL = lSQL & " ORDER BY Emitente, [Numero do Cheque], [Data do Vencimento], [Data do Pagamento], [Periodo do Pagamento]"
    'Abre RecordSet
    Set rsBaixaCheque = New ADODB.Recordset
    Set rsBaixaCheque = Conectar.RsConexao(lSQL)
    
    If rsBaixaCheque.RecordCount > 0 Then
        Do Until rsBaixaCheque.EOF
            lTotalBaixaCheque = lTotalBaixaCheque + rsBaixaCheque("Valor").Value
            rsBaixaCheque.MoveNext
        Loop
    End If
End Sub
Private Sub TotalizaBaixaChequeDevolvido()
    'Prepara SQL
    lSQL = ""
    'lSQL = lSQL & "SELECT Emitente, [Numero do Cheque], Valor, [Data do Vencimento], [Data do Pagamento], [Periodo do Pagamento]"
    lSQL = lSQL & "SELECT Emitente, [Numero do Cheque], [Data do Vencimento], [Data do Pagamento], Periodo, "
    lSQL = lSQL & "([Valor Pago Dinheiro] + [Valor Pago Cheque a Vista] + [Valor Pago Cheque a Prazo]) AS Valor"
    lSQL = lSQL & "  FROM Baixa_Cheque_Devolvido"
    lSQL = lSQL & " WHERE Baixa_Cheque_Devolvido.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Baixa_Cheque_Devolvido.[Data do Pagamento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Baixa_Cheque_Devolvido.[Data do Pagamento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Baixa_Cheque_Devolvido.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Baixa_Cheque_Devolvido.Periodo <= " & Val(cbo_periodo_f.Text)
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Baixa_Cheque_Devolvido.[Data do Pagamento], 103) + CONVERT(VARCHAR, Baixa_Cheque_Devolvido.[Periodo]) IN " & lDataPeriodo
    End If
    lSQL = lSQL & " ORDER BY Emitente, [Numero do Cheque], [Data do Vencimento], [Data do Pagamento], [Periodo]"
    'Abre RecordSet
    Set rsBaixaChequeDevolvido = New ADODB.Recordset
    Set rsBaixaChequeDevolvido = Conectar.RsConexao(lSQL)
    
    If rsBaixaChequeDevolvido.RecordCount > 0 Then
        Do Until rsBaixaChequeDevolvido.EOF
            lTotalBaixaChequeDevolvido = lTotalBaixaChequeDevolvido + rsBaixaChequeDevolvido("Valor").Value
            rsBaixaChequeDevolvido.MoveNext
        Loop
    End If
End Sub
Private Sub TotalizaBaixaDuplicataAReceber()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT [Numero do Documento], [Codigo do Cliente], Cliente.[Razao Social], "
    lSQL = lSQL & "       [Data do Periodo Inicial], [Data do Periodo Final], [Data do Vencimento], "
    lSQL = lSQL & "       [Valor do Vencimento], [Data do Pagamento], Periodo, "
    lSQL = lSQL & "       [Valor Pago] AS [Valor Pago Dinheiro], [Valor Pago Cheque Vista], "
    lSQL = lSQL & "       [Valor Pago Cheque PRAZO], [Valor Pago Banco], [Valor Pago Cartao], "
    lSQL = lSQL & "      ([Valor Pago] + [Valor Pago Cheque Vista] + [Valor Pago Cheque PRAZO] + [Valor Pago Banco] + [Valor Pago Cartao]) AS TotalRecebido"
    lSQL = lSQL & "  FROM [baixa_duplicata_receber], Cliente"
    lSQL = lSQL & " WHERE Baixa_Duplicata_Receber.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Baixa_Duplicata_Receber.[Data do Pagamento] >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "   AND Baixa_Duplicata_Receber.[Data do Pagamento] <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "   AND Baixa_Duplicata_Receber.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Baixa_Duplicata_Receber.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQL = lSQL & "   AND Baixa_Duplicata_Receber.[Codigo do Cliente] = Cliente.Codigo"
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Baixa_Duplicata_Receber.[Data do Pagamento], 103) + CONVERT(VARCHAR, Baixa_Duplicata_Receber.Periodo) IN " & lDataPeriodo
    End If
    lSQL = lSQL & " ORDER BY Cliente.[Razao Social], [Data do Pagamento], [Data do Vencimento]"
    'Abre RecordSet
    Set rsBaixaDuplicataReceber = New ADODB.Recordset
    Set rsBaixaDuplicataReceber = Conectar.RsConexao(lSQL)
    
    If rsBaixaDuplicataReceber.RecordCount > 0 Then
        Do Until rsBaixaDuplicataReceber.EOF
            lTotalDuplicataRecebida = lTotalDuplicataRecebida + rsBaixaDuplicataReceber("TotalRecebido").Value
            rsBaixaDuplicataReceber.MoveNext
        Loop
    End If
End Sub
Private Sub TotalizaCaixaPista()
    Dim i As Integer

    'If txtdatainicial.Text = txtdatafinal.Text And Val(cbo_periodo_i.Text) = Val(cbo_periodo_f.Text) And Val(cboIlhaI.Text) = Val(cboIlhaF.Text) Then
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT LancamentoPadrao.Nome,"
        lSQL = lSQL & "       MovimentoCaixaPista.[Codigo do Lancamento Padrao],"
        lSQL = lSQL & "       0 AS ValorSaida, SUM(Valor) AS ValorEntrada,"
        lSQL = lSQL & "       Plano_Conta.Nome AS NomePlano"
        lSQL = lSQL & "  FROM MovimentoCaixaPista, LancamentoPadrao, Plano_Conta"
        lSQL = lSQL & " WHERE MovimentoCaixaPista.Empresa = " & g_empresa
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo >= " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo <= " & Val(cbo_periodo_f.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
        If lTipoMovimento > 0 Then
            lSQL = lSQL & "   AND MovimentoCaixaPista.[Tipo do Movimento] = " & lTipoMovimento
        End If
        If bdAccess Then
            lSQL = lSQL & "   AND MID(MovimentoCaixaPista.[Numero da Conta Credito],1,5) = " & preparaTexto("11101")
        ElseIf bdSqlServer Then
            lSQL = lSQL & "   AND LEFT(MovimentoCaixaPista.[Numero da Conta Credito],5) = " & preparaTexto("11101")
        End If
        lSQL = lSQL & "   AND LancamentoPadrao.Empresa = " & g_empresa
        lSQL = lSQL & "   AND LancamentoPadrao.Codigo = MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        lSQL = lSQL & "   AND Plano_Conta.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Plano_Conta.Codigo = MovimentoCaixaPista.[Numero da Conta Debito]"
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            lSQL = lSQL & "    AND CONVERT(VARCHAR, MovimentoCaixaPista.Data, 103) + CONVERT(VARCHAR, MovimentoCaixaPista.Periodo) IN " & lDataPeriodo
        End If
        lSQL = lSQL & " GROUP BY Plano_Conta.Nome, LancamentoPadrao.Nome, MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        If bdAccess Then
            lSQL = lSQL & " ORDER BY Plano_Conta.Nome, LancamentoPadrao.Nome, MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        End If
        lSQL = lSQL & " "
        lSQL = lSQL & " UNION"
        lSQL = lSQL & " "
        lSQL = lSQL & "SELECT LancamentoPadrao.Nome,"
        lSQL = lSQL & "       MovimentoCaixaPista.[Codigo do Lancamento Padrao],"
        lSQL = lSQL & "       SUM(Valor) AS ValorSaida, 0 AS ValorEntrada,"
        lSQL = lSQL & "       Plano_Conta.Nome AS NomePlano"
        lSQL = lSQL & "  FROM MovimentoCaixaPista, LancamentoPadrao, Plano_Conta"
        lSQL = lSQL & " WHERE MovimentoCaixaPista.Empresa = " & g_empresa
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo >= " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo <= " & Val(cbo_periodo_f.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
        If lTipoMovimento > 0 Then
            lSQL = lSQL & "   AND MovimentoCaixaPista.[Tipo do Movimento] = " & lTipoMovimento
        End If
        If bdAccess Then
            lSQL = lSQL & "   AND MID(MovimentoCaixaPista.[Numero da Conta Debito],1,5) = " & preparaTexto("11101")
        ElseIf bdSqlServer Then
            lSQL = lSQL & "   AND LEFT(MovimentoCaixaPista.[Numero da Conta Debito],5) = " & preparaTexto("11101")
        End If
        lSQL = lSQL & "   AND LancamentoPadrao.Empresa = " & g_empresa
        lSQL = lSQL & "   AND LancamentoPadrao.Codigo = MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        lSQL = lSQL & "   AND Plano_Conta.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Plano_Conta.Codigo = MovimentoCaixaPista.[Numero da Conta Credito]"
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            lSQL = lSQL & "    AND CONVERT(VARCHAR, MovimentoCaixaPista.Data, 103) + CONVERT(VARCHAR, MovimentoCaixaPista.Periodo) IN " & lDataPeriodo
        End If
        lSQL = lSQL & " GROUP BY Plano_Conta.Nome, LancamentoPadrao.Nome, MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        lSQL = lSQL & " "
        lSQL = lSQL & " UNION"
        lSQL = lSQL & " "
        lSQL = lSQL & "SELECT LancamentoPadrao.Nome,"
        lSQL = lSQL & "       MovimentoCaixaPista.[Codigo do Lancamento Padrao],"
        lSQL = lSQL & "       0 AS ValorSaida, SUM(Valor) AS ValorEntrada,"
        lSQL = lSQL & "       LancamentoPadrao.Nome AS NomePlano"
        lSQL = lSQL & "  FROM MovimentoCaixaPista, LancamentoPadrao"
        lSQL = lSQL & " WHERE MovimentoCaixaPista.Empresa = " & g_empresa
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data >= " & preparaData(CDate(txtDataInicial.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Data <= " & preparaData(CDate(txtDataFinal.Text))
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo >= " & Val(cbo_periodo_i.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.Periodo <= " & Val(cbo_periodo_f.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
        lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
        If lTipoMovimento > 0 Then
            lSQL = lSQL & "   AND MovimentoCaixaPista.[Tipo do Movimento] = " & lTipoMovimento
        End If
        If bdAccess Then
            lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Conta Debito] = " & preparaTexto("")
        ElseIf bdSqlServer Then
            lSQL = lSQL & "   AND MovimentoCaixaPista.[Numero da Conta Debito] = " & preparaTexto("")
        End If
        lSQL = lSQL & "   AND LancamentoPadrao.Empresa = " & g_empresa
        lSQL = lSQL & "   AND LancamentoPadrao.Codigo = MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
            lSQL = lSQL & "    AND CONVERT(VARCHAR, MovimentoCaixaPista.Data, 103) + CONVERT(VARCHAR, MovimentoCaixaPista.Periodo) IN " & lDataPeriodo
        End If
        lSQL = lSQL & " GROUP BY LancamentoPadrao.Nome, MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        
        lSQL = lSQL & " ORDER BY Plano_Conta.Nome, LancamentoPadrao.Nome, MovimentoCaixaPista.[Codigo do Lancamento Padrao]"
        'Abre RecordSet
        Set rsMovimentoCaixaPista = New ADODB.Recordset
        Set rsMovimentoCaixaPista = Conectar.RsConexao(lSQL)
    'End If
    
End Sub
Private Sub TotalizaLubrificante()
    'loop movimento de lubrificante
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT SUM(Movimento_Lubrificante.[Valor Total]) AS Total"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data >= " & preparaData(txtDataInicial.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data <= " & preparaData(txtDataFinal.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 2
    If lTipoMovimento > 0 Then
        lSQL = lSQL & "      AND Movimento_Lubrificante.[Tipo do Movimento] = " & lTipoMovimento
    End If
    If cbo_funcionario.ItemData(cbo_funcionario.ListIndex) > 0 And (txtDataInicial.Text <> txtDataFinal.Text) Then
        lSQL = lSQL & "    AND CONVERT(VARCHAR, Movimento_Lubrificante.Data, 103) + CONVERT(VARCHAR, Movimento_Lubrificante.Periodo) IN " & lDataPeriodo
    End If
    'Abre RecordSet
    Set rsMovLubrificante = New ADODB.Recordset
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    rsMovLubrificante.MoveFirst
    If Not IsNull(rsMovLubrificante("Total").Value) Then
        lTotalLubrificante = rsMovLubrificante("Total").Value
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
End Sub
Private Sub ImpResumoCombustiveis()
    Dim xLinha As String
    Dim i As Integer
    Dim xString(0 To 5) As String
    Dim xValor As Currency
    Dim xNLinha As Integer
    Dim xQtdLinha As Integer
    
    xNLinha = -1
    xQtdLinha = 0
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
    
    
    
    
    
    'Prepara xString
    For i = 0 To 5
        xString(i) = Space(36)
    Next
    xValor = lTotalEntrada - lTotalSaida
    If xValor < 0 Then
        Mid(xString(0), 1, 23) = "ESTÁ FALTANDO NO CAIXA "
        i = Len(Format(xValor, "#,###,##0.00"))
        Mid(xString(0), 25 + 12 - i, i) = Format(xValor, "#,###,##0.00")
    ElseIf xValor > 0 Then
        Mid(xString(0), 3, 40) = "ESTÁ PASSANDO NO CAIXA  "
        i = Len(Format(xValor, "#,###,##0.00"))
        Mid(xString(0), 25 + 12 - i, i) = Format(xValor, "#,###,##0.00")
    ElseIf xValor = 0 Then
        Mid(xString(0), 3, 40) = "O CAIXA ESTÁ CORRETO    "
        i = Len(Format(xValor, "#,###,##0.00"))
        Mid(xString(0), 25 + 12 - i, i) = Format(xValor, "#,###,##0.00")
    End If
    xString(1) = "------------------------------------"
    xString(2) = "                                    "
    xString(3) = "____________________________________"
    xString(4) = "                                    "
    If Funcionario.LocalizarCodigo(g_empresa, AberturaCaixa.CodigoFuncionario) Then
        i = Len(Mid(Trim(Funcionario.Nome), 1, 36))
        Mid(xString(4), 1 + (36 - i) / 2, i) = Mid(Trim(Funcionario.Nome), 1, 36)
    End If
    xString(5) = "       RESPONSAVEL PELO CAIXA       "
    
    
    
    If lBombaLitrosA > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("ÁLCOOL    ", lBombaLitrosA, lBombaValorA, xString(xNLinha))
    End If
    If lBombaLitrosAA > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("ÁLCOOL +  ", lBombaLitrosAA, lBombaValorAA, xString(xNLinha))
    End If
    If lBombaLitrosD > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("DIESEL    ", lBombaLitrosD, lBombaValorD, xString(xNLinha))
    End If
    If lBombaLitrosDA > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("DIESEL +  ", lBombaLitrosDA, lBombaValorDA, xString(xNLinha))
    End If
    If lBombaLitrosG > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("GASOLINA  ", lBombaLitrosG, lBombaValorG, xString(xNLinha))
    End If
    If lBombaLitrosGA > 0 Then
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("GASOLINA +", lBombaLitrosGA, lBombaValorGA, xString(xNLinha))
    End If
    
    Do Until xNLinha >= 3
        xNLinha = xNLinha + 1
        Call ImpDetCombustivel("          ", 0, 0, xString(xNLinha))
    Loop
    
    
    If xNLinha = 5 Then
        xNLinha = 0
        xString(0) = Space(36)
    Else
        xNLinha = xNLinha + 1
    End If
    Call ImpDetCombustivel("----------", 0, 0, xString(xNLinha))
    
    
    If xNLinha = 5 Then
        xNLinha = 0
        xString(0) = Space(36)
    Else
        xNLinha = xNLinha + 1
    End If
    Call ImpDetCombustivel("TOTAL     ", lBombaTotalLitros, lBombaTotalValor, xString(xNLinha))
    
    If chkImprimeLucro.Value = 0 Then
        If lTipoMovimento = 1 Then
            BioImprime "@Printer.Print " & "+---------------------------------------+--------------------------------------+"
            lLinha = lLinha + 1
        Else
            BioImprime "@Printer.Print " & "+-----------+-------------+-------------+--------------------------------------+"
            lLinha = lLinha + 1
        End If
    End If
    
End Sub
Private Sub ImpResumoLucroCombustiveis()
    Dim xLinha As String
    Dim i As Integer
    Dim xValor As Currency
    Dim xTotalCusto As Currency
    Dim xCustoLubrificante As Currency
    Dim xVendaLubrificante As Currency
    Dim xQuantidadeLubrificante As Currency
    
    xCustoLubrificante = 0
    xVendaLubrificante = 0
    xQuantidadeLubrificante = 0
    
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "|COMBUSTÍVEL|    LITROS   |    VALOR    |LUCRO  MEDIO|LUCRO  VENDA| % DO LUCRO |"
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    lLinha = lLinha + 3
    
    If lBombaLitrosA > 0 Then
        lBombaCustoA = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "A ", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoA = lBombaCustoA - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "A ", lDataPeriodo)
        Call ImpDetLucroCombustivel("ÁLCOOL    ", lBombaLitrosA, lBombaValorA, lBombaCustoA)
    End If
    If lBombaLitrosAA > 0 Then
        lBombaCustoAA = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "AA", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoAA = lBombaCustoAA - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "AA", lDataPeriodo)
        Call ImpDetLucroCombustivel("ÁLCOOL +  ", lBombaLitrosAA, lBombaValorAA, lBombaCustoAA)
    End If
    If lBombaLitrosD > 0 Then
        lBombaCustoD = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "D ", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoD = lBombaCustoD - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "D ", lDataPeriodo)
        Call ImpDetLucroCombustivel("DIESEL    ", lBombaLitrosD, lBombaValorD, lBombaCustoD)
    End If
    If lBombaLitrosDA > 0 Then
        lBombaCustoDA = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "DA", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoDA = lBombaCustoDA - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "DA", lDataPeriodo)
        Call ImpDetLucroCombustivel("DIESEL +  ", lBombaLitrosDA, lBombaValorDA, lBombaCustoDA)
    End If
    If lBombaLitrosG > 0 Then
        lBombaCustoG = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "G ", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoG = lBombaCustoG - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "G ", lDataPeriodo)
        Call ImpDetLucroCombustivel("GASOLINA  ", lBombaLitrosG, lBombaValorG, lBombaCustoG)
    End If
    If lBombaLitrosGA > 0 Then
        lBombaCustoGA = MovimentoBomba.ValorCustoVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "GA", Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), lDataPeriodo)
        lBombaCustoGA = lBombaCustoGA - MovimentoAfericao.ValorTotalCustoPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), "GA", lDataPeriodo)
        Call ImpDetLucroCombustivel("GASOLINA +", lBombaLitrosGA, lBombaValorGA, lBombaCustoGA)
    End If
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT SUM(Movimento_Lubrificante.[Valor Total]) AS TotalVenda, SUM(Movimento_Lubrificante.[Valor Custo] * Movimento_Lubrificante.Quantidade) AS TotalCusto, SUM(Movimento_Lubrificante.Quantidade) AS TotalQuantidade"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "      AND Movimento_Lubrificante.Data <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] >= " & Val(cboIlhaI.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Numero da Ilha] <= " & Val(cboIlhaF.Text)
    lSQL = lSQL & "      AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & 2
    If lTipoMovimento > 0 Then
        lSQL = lSQL & "      AND Movimento_Lubrificante.[Tipo do Movimento] = " & lTipoMovimento
    End If
    'Abre RecordSet
    Set rsMovLubrificante = New ADODB.Recordset
    Set rsMovLubrificante = Conectar.RsConexao(lSQL)
    If rsMovLubrificante.RecordCount > 0 Then
        If Not IsNull(rsMovLubrificante("TotalCusto").Value) Then
            xCustoLubrificante = rsMovLubrificante("TotalCusto").Value
            xVendaLubrificante = rsMovLubrificante("TotalVenda").Value
            xQuantidadeLubrificante = rsMovLubrificante("TotalQuantidade").Value
        End If
    End If
    rsMovLubrificante.Close
    Set rsMovLubrificante = Nothing
    If xCustoLubrificante > 0 Then
        Call ImpDetLucroCombustivel("PRODUTOS  ", xQuantidadeLubrificante, xVendaLubrificante, xCustoLubrificante)
    End If
    
    xTotalCusto = lBombaCustoA + lBombaCustoAA + lBombaCustoD + lBombaCustoDA + lBombaCustoG + lBombaCustoGA + xCustoLubrificante
    lBombaTotalLitros = lBombaTotalLitros + xQuantidadeLubrificante
    lBombaTotalValor = lBombaTotalValor + xVendaLubrificante
    BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
    lLinha = lLinha + 1
    Call ImpDetLucroCombustivel("TOTAL      ", lBombaTotalLitros, lBombaTotalValor, xTotalCusto)
    If chkImprimeMedidaTanque.Value = 0 Then
        BioImprime "@Printer.Print " & "+-----------+-------------+-------------+------------+------------+------------+"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpResumoMedicaoCombustiveis()
    Dim i As Integer
    
    
    'If lLinha >= 55 Then
    If lLinha >= 65 Then
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    
    If chkImprimeLucro.Value = 1 Then
        BioImprime "@Printer.Print " & "+-----------+-----------+-+-------+-----+-----+------+-----+------+--+---------+"
        lLinha = lLinha + 1
    Else
        BioImprime "@Printer.Print " & "+-----------+-----------+---------+-----------+------------+---------+---------+"
        lLinha = lLinha + 1
    End If
    BioImprime "@Printer.Print " & "|COMBUSTÍVEL|EST.INICIAL| ENTRADAS|QTD. SAIDAS| EST.ESCRIT.|EST.FINAL|PERD/SOBR|"
    lLinha = lLinha + 1
    BioImprime "@Printer.Print " & "+-----------+-----------+---------+-----------+------------+---------+---------+"
    lLinha = lLinha + 1
    Call ImpDetMedicaoCombustivel("A ")
    Call ImpDetMedicaoCombustivel("AA")
    Call ImpDetMedicaoCombustivel("D ")
    Call ImpDetMedicaoCombustivel("DA")
    Call ImpDetMedicaoCombustivel("G ")
    Call ImpDetMedicaoCombustivel("GA")
    BioImprime "@Printer.Print " & "+-----------+-----------+---------+-----------+------------+---------+---------+"
    lLinha = lLinha + 1
    BioImprime "@@Printer.FontName = Draft 10cpi"
End Sub
Private Sub ImpDetBomba(ByVal pBico As Integer, ByVal pAbertura As Currency, ByVal pEncerrante As Currency, ByVal pAfericaoLitros As Currency, ByVal pLitros As Currency, ByVal pValorTotal As Currency, ByVal pTipoPreco As String, ByVal pCombustivel As String)
    Dim xLinha As String
    Dim xValorUnitario As String
    Dim i As Integer
    Dim xStringAbertura As String
    Dim xStringEncerrante As String
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|  |             |             |         |         |         |           |  |  |"

    If pTipoPreco = "SEM MOV.MEC." And pCombustivel = "SEM MOV.MEC." Then
        Mid(xLinha, 2, 2) = Format(pBico, "00")
        Mid(xLinha, 5, 12) = " **  SEM    "
        Mid(xLinha, 19, 12) = " MOVIMENTO  "
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        Exit Sub
    End If

    Mid(xLinha, 2, 2) = Format(pBico, "00")
    xStringAbertura = Space(13)
    xStringEncerrante = Space(13)
    
    i = Len(Format(pAbertura, "#,###,##0.00"))
    Mid(xStringAbertura, 2 + 12 - i, i) = Format(pAbertura, "#,###,##0.00")
    i = Len(Format(pEncerrante, "#,###,##0.00"))
    Mid(xStringEncerrante, 2 + 12 - i, i) = Format(pEncerrante, "#,###,##0.00")
    If lInverteEncerrante Then
        Mid(xLinha, 5, 13) = xStringEncerrante
        Mid(xLinha, 19, 13) = xStringAbertura
    Else
        Mid(xLinha, 5, 13) = xStringAbertura
        Mid(xLinha, 19, 13) = xStringEncerrante
    End If
    i = Len(Format(pAfericaoLitros, "##,##0.00"))
    Mid(xLinha, 33 + 9 - i, i) = Format(pAfericaoLitros, "##,##0.00")
    i = Len(Format(pLitros, "##,##0.00"))
    Mid(xLinha, 43 + 9 - i, i) = Format(pLitros, "##,##0.00")
    If pLitros = 0 Then
        xValorUnitario = 0
        If MovimentoBomba.LocalizarCodigo(g_empresa, CDate(txtDataFinal.Text), Val(cbo_periodo_f.Text), pBico, 999) Then
            xValorUnitario = MovimentoBomba.PrecoVenda
        End If
    Else
        xValorUnitario = Format(pValorTotal / pLitros, "00000.0000")
    End If
    i = Len(Format(xValorUnitario, "###0.0000"))
    Mid(xLinha, 53 + 9 - i, i) = Format(xValorUnitario, "###0.0000")
    i = Len(Format(pValorTotal, "##,###,##0.00"))
    Mid(xLinha, 61 + 13 - i, i) = Format(pValorTotal, "##,###,##0.00")
    
    If pTipoPreco = "V" Then
        Mid(xLinha, 75, 2) = "VI"
    ElseIf pTipoPreco = "P" Then
        Mid(xLinha, 75, 2) = "PR"
    End If
    Mid(xLinha, 78, 2) = pCombustivel
    If pTipoPreco = "DIFERENCA MEC." And pCombustivel = "DIFERENCA MEC." Then
        Mid(xLinha, 75, 5) = "E.MEC"
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    If lBombaMensagem(pBico) <> "" Then
        xLinha = "**                                                                            **"
        Mid(xLinha, 4, 70) = lBombaMensagem(pBico)
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    If pTipoPreco = "DIFERENCA MEC." And pCombustivel = "DIFERENCA MEC." Then
        xLinha = "**  DIFERENÇA NA VENDA DO BICO ACIMA. ENCERRANTE ELETRONICO X MECANICO        **"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetCombustivel(ByVal pCombustivel As String, ByVal pLitros As Currency, ByVal pTotal As Currency, ByVal pString As String)
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
    xLinha = "|           |             |             |                                      |"
    
    If pCombustivel = "----------" Then
        Mid(xLinha, 1, 41) = "+-----------+-------------+-------------+"
        Mid(xLinha, 43, 36) = pString
        If pString = "------------------------------------" Then
            Mid(xLinha, 41, 40) = "+--------------------------------------+"
        End If
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
        Mid(xLinha, 43, 36) = pString
        If pString = "------------------------------------" Then
            Mid(xLinha, 41, 40) = "+--------------------------------------+"
        End If
    End If
    
    If lTipoMovimento = 1 Then
        Mid(xLinha, 1, 40) = "|                                       "
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetLucroCombustivel(ByVal pCombustivel As String, ByVal pLitros As Currency, ByVal pTotalVenda As Currency, ByVal pTotalCusto As Currency)
    Dim xLinha As String
    Dim i As Integer
    Dim xLucroMedio As Currency
    Dim xLucroVenda As Currency
    Dim xPercentual As Currency
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|           |             |             |            |            |            |"

    xLucroMedio = (pTotalVenda - pTotalCusto) / pLitros
    xLucroVenda = pTotalVenda - pTotalCusto
    If pTotalVenda > 0 Then
        xPercentual = xLucroVenda * 100 / pTotalVenda
    Else
        xPercentual = 0
    End If
'    If pCombustivel Like "*TOTAL*" Then
'    Else
        Mid(xLinha, 3, 10) = pCombustivel
        i = Len(Format(pLitros, "#,###,##0.00"))
        Mid(xLinha, 15 + 12 - i, i) = Format(pLitros, "#,###,##0.00")
        i = Len(Format(pTotalVenda, "#,###,##0.00"))
        Mid(xLinha, 28 + 12 - i, i) = Format(pTotalVenda, "#,###,##0.00")
        i = Len(Format(xLucroMedio, "#,##0.0000"))
        Mid(xLinha, 43 + 10 - i, i) = Format(xLucroMedio, "#,##0.0000")
        i = Len(Format(xLucroVenda, "#,###,##0.00"))
        Mid(xLinha, 55 + 12 - i, i) = Format(xLucroVenda, "#,###,##0.00")
        i = Len(Format(xPercentual, "#,##0.0000"))
        Mid(xLinha, 69 + 10 - i, i) = Format(xPercentual, "#,##0.0000")
'    End If
    
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetMedicaoCombustivel(ByVal pTipoCombustivel As String)
    Dim xLinha As String
    Dim i As Integer
    Dim xEstoqueInicial As Currency
    Dim xEstoqueFinal As Currency
    Dim xQtdEntrada As Currency
    Dim xQtdSaida As Currency
    Dim xEstEscritural As Currency
    Dim xPerdaSobra As Currency
    
    xLinha = "|           |           |         |           |            |         |         |"

    
    xEstoqueInicial = 0
    xEstoqueFinal = 0
    
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Mid(xLinha, 2, 11) = Combustivel.Nome
    End If
    
    xEstoqueInicial = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(txtDataInicial.Text), pTipoCombustivel, 0)
    xEstoqueFinal = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(txtDataFinal.Text) + 1, pTipoCombustivel, 0)
    xQtdEntrada = EntradaCombustivel.TotalEntradaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), pTipoCombustivel, 0)
    
    xQtdSaida = MovimentoBomba.TotalVendaPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), pTipoCombustivel, Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text))
    xQtdSaida = xQtdSaida - MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), pTipoCombustivel, lDataPeriodo)
    
    i = Len(Format(xEstoqueInicial, "#,###,##0.0"))
    Mid(xLinha, 14 + 11 - i, i) = Format(xEstoqueInicial, "#,###,##0.0")
    i = Len(Format(xQtdEntrada, "#,###,##0"))
    Mid(xLinha, 26 + 9 - i, i) = Format(xQtdEntrada, "#,###,##0")
    
    xEstEscritural = xEstoqueInicial + xQtdEntrada - xQtdSaida
    xPerdaSobra = xEstoqueFinal - xEstEscritural
    
    i = Len(Format(xQtdSaida, "####,##0.00"))
    Mid(xLinha, 36 + 11 - i, i) = Format(xQtdSaida, "####,##0.00")
    
    i = Len(Format(xEstEscritural, "#,###,##0.00"))
    Mid(xLinha, 48 + 12 - i, i) = Format(xEstEscritural, "#,###,##0.00")
    
    i = Len(Format(xEstoqueFinal, "#,###,##0"))
    Mid(xLinha, 61 + 9 - i, i) = Format(xEstoqueFinal, "#,###,##0")

    i = Len(Format(xPerdaSobra, "##,##0.00"))
    Mid(xLinha, 71 + 9 - i, i) = Format(xPerdaSobra, "##,##0.00")
    
    If xEstoqueInicial <> 0 Or xEstoqueFinal <> 0 Or xQtdSaida <> 0 Then
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetLubrificante()
    Dim xLinha As String
    Dim xValorUnitario As Currency
    Dim i As Integer
    
    'If lLinha >= 55 Then
    If lLinha >= 63 Then
        xLinha = "+----+-- Cerrado Informatica -----------------+------------+------+------------+"
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
        ImpCabLubrificante
    End If
    
    '                  1         2         3         4         5         6         7         8         9        10        11        12   12
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
    xLinha = "|    |                                        |            |      |            |"
    
    If chkImprimeLubrificante.Value = 1 Then
        i = Len(Format(rsMovLubrificante("Codigo do Produto2").Value, "###0"))
        Mid(xLinha, 2 + 4 - i, i) = Format(rsMovLubrificante("Codigo do Produto2").Value, "###0")
    Else
        i = Len(Format(rsMovLubrificante("Codigo").Value, "###0"))
        Mid(xLinha, 2 + 4 - i, i) = Format(rsMovLubrificante("Codigo").Value, "###0")
    End If
    Mid(xLinha, 7, 40) = rsMovLubrificante("Nome").Value
            
    xValorUnitario = Format(rsMovLubrificante("Valor Total").Value / rsMovLubrificante("Quantidade").Value, "0000000000.00")
    i = Len(Format(xValorUnitario, "#,###,###.00"))
    Mid(xLinha, 48 + 12 - i, i) = Format(xValorUnitario, "#,###,###.00")
    
    If (CCur(rsMovLubrificante("Quantidade").Value) - Val(rsMovLubrificante("Quantidade").Value)) = 0 Then
        i = Len(Format(rsMovLubrificante("Quantidade").Value, "##,###"))
        Mid(xLinha, 61 + 6 - i, i) = Format(rsMovLubrificante("Quantidade").Value, "##,###")
    Else
        i = Len(FormatNumber(rsMovLubrificante("Quantidade").Value, 2))
        Mid(xLinha, 61 + 6 - i, i) = FormatNumber(rsMovLubrificante("Quantidade").Value, 2)
    End If
            
    i = Len(Format(rsMovLubrificante("Valor Total").Value, "#,###,###.00"))
    Mid(xLinha, 68 + 12 - i, i) = Format(rsMovLubrificante("Valor Total").Value, "#,###,###.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDet(x_i As Integer, x_abertura As Currency, x_encerrante As Currency, x_litros As Currency, x_valor As Currency, x_historico As String, x_variavel As String)
    Dim xLinha As String
    Dim i As Integer
    xLinha = "|  |          |          |         |              |                            |"
    If x_i > 0 Then
        Mid(xLinha, 2, 2) = Format(x_i, "00")
    End If
    If CCur(x_abertura) Or CCur(x_encerrante) > 0 Then
        i = Len(Format(x_abertura, "####,##0.0"))
        Mid(xLinha, 5 + 10 - i, i) = Format(x_abertura, "####,##0.0")
        i = Len(Format(x_encerrante, "####,##0.0"))
        Mid(xLinha, 16 + 10 - i, i) = Format(x_encerrante, "####,##0.0")
        i = Len(Format(x_litros, "###,##0.0"))
        Mid(xLinha, 27 + 9 - i, i) = Format(x_litros, "###,##0.0")
        i = Len(Format(x_valor, "##,###,##0.00"))
        Mid(xLinha, 37 + 13 - i, i) = Format(x_valor, "##,###,##0.00")
    End If
    Mid(xLinha, 52, 13) = x_historico
    If Mid(x_variavel, 1, 3) = "@N@" Then
        If CCur(Mid(x_variavel, 4, Len(x_variavel) - 3)) > 0 Then
            x_variavel = Mid(x_variavel, 4, Len(x_variavel) - 3)
            i = Len(Format(x_variavel, "##,###,##0.00"))
            Mid(xLinha, 66 + 13 - i, i) = Format(x_variavel, "##,###,##0.00")
        End If
    Else
        Mid(xLinha, 65, 15) = Mid(x_variavel, 4, Len(x_variavel) - 3)
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
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
'    '                  1         2         3         4         5         6         7         8
'    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
'    '                                             123456789012345678901234567890
    xLinha = "| CAIXA DE PISTA                                            cidade, __/__/____ |"
    If lTipoMovimento = 1 Then
        Mid(xLinha, 3, 27) = "CAIXA DE CONVENIENCIA      "
    End If
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 24, 8) = "- L.M.C."
    End If
    If fEcfInstalada Then
        Mid(xLinha, 24, 8) = "- E.C.F."
    End If
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = txtDataEmissao.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(xLinha, 29, 10) = txtDataInicial.Text
    Mid(xLinha, 42, 10) = txtDataFinal.Text
    Mid(xLinha, 55, 20) = cbo_funcionario.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| PERIODO INICIAL.: X   PERIODO FINAL.: X    ILHA INICIAL.: X   ILHA FINAL.: X |"
    Mid(xLinha, 21, 1) = cbo_periodo_i.Text
    Mid(xLinha, 41, 1) = cbo_periodo_f.Text
    Mid(xLinha, 61, 1) = cboIlhaI.Text
    Mid(xLinha, 78, 1) = cboIlhaF.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| IMPRESSO EM.....: __/__/____ AS __:__:__ POR.:                               |"
    Mid(xLinha, 21, 10) = Format(Date, "dd/mm/yyyy")
    Mid(xLinha, 35, 8) = Format(Time, "HH:mm:ss")
    Mid(xLinha, 50, 29) = g_nome_usuario
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 6
    
    If lExistePendenciaCartao = True Then
        xLinha = "| *****  EXISTE PROBLEMAS NA CONCILIAÇÃO DE CARTÕES, VERIFIQUE O MOTIVO  ***** |"
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
    
    If CDate(txtDataInicial.Text) = CDate(txtDataFinal.Text) And Val(cbo_periodo_i.Text) = Val(cbo_periodo_f.Text) Then
        
        If AberturaCaixa.LocalizarCxData(g_empresa, CDate(txtDataInicial.Text), "NF", Val(cbo_periodo_i.Text), Val(cboIlhaI.Text), Val(cboTipoCaixa.Text)) Then
            If AberturaCaixa.DataFechamento = CDate("00:00:00") Then
                xLinha = "| *****  CAIXA NAO ESTÁ FECHADO - CUIDADO, VERIFIQUE O MOTIVO            ***** |"
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
            Else
                xLinha = "| CAIXA FECHADO EM.: __/__/____ AS __:__:__   FECHADO NO NIVEL:                |"
                '                  1         2         3         4         5         6         7         8
                '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
                Mid(xLinha, 22, 10) = fMascaraData(AberturaCaixa.DataFechamento)
                Mid(xLinha, 36, 8) = fMascaraHora(AberturaCaixa.HoraFechamento)
                Mid(xLinha, 65, 14) = NivelAcesso(AberturaCaixa.FechadoPeloNivel)
                BioImprime "@Printer.Print " & xLinha
                lLinha = lLinha + 1
            End If
        End If
    
    End If
    
    
End Sub
Private Sub ImpCabBaixaCheque()
    Dim xLinha As String
    
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|NUM.CH.| EMITENTE                                | DATA VENC. |     VALOR     |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 3
End Sub
Private Sub ImpCabBaixaDuplicataAReceber()
    Dim xLinha As String
    
    xLinha = "+-------+------------------------------------------------+--------+------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|CÓDIGO | RAZÃO SOCIAL                                   |NUM.DOC.|    VALOR   |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|CLIENTE| DT.INICIAL  DT.  FINAL  DT.VENCIM.    VLR.VENC |RECEB.EM|  RECEBIDO  |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------+------------------------------------------------+--------+------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 4
End Sub
Private Sub ImpCabCaixaPista()
    Dim xLinha As String
    
    If chkImprimeLubrificante.Value = 0 Then
        xLinha = "+--+--+----------+-------------+---------+-------+-+---------+--+--------+--+--+"
        lLinha = lLinha + 1
    Else
        If lTipoMovimento = 1 Then
            xLinha = "+----++---------------------------------------+--+---------+----+-+------------+"
            lLinha = lLinha + 1
        Else
            xLinha = "+--+--+----------+-------------+---------+-------+-+---------+--+--------+--+--+"
            lLinha = lLinha + 1
        End If
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| COD | DESCRICAO DA CONTA                       | VLR. ENTRADA | VLR.   SAIDA |"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    xLinha = "+-----+------------------------------------------+--------------+--------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCabCombustivel()
    Dim xLinha As String
    
    If lTipoMovimento = 2 Then
        xLinha = "+--+-------------+-------------+---------+---------+---------+-----------+--+--+"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|N.|   ABERTURA  |  ENCERRANTE |LT.AFERIC|LTS.SAIDA|VLR LITRO|VALOR SAIDA|PR|CB|"
        If lInverteEncerrante Then
            Mid(xLinha, 5, 13) = "  ENCERRANTE "
            Mid(xLinha, 19, 13) = "   ABERTURA  "
        End If
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+--+-------------+-------------+---------+---------+---------+-----------+--+--+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 3
    End If
End Sub
Private Sub ImpCabLubrificante()
    Dim xLinha As String
    If lTipoMovimento = 2 Then
        BioImprime "@Printer.Print " & "+--+-+-----------+-------------+---------+----+-----+------+------++--------+--+"
        lLinha = lLinha + 1
    Else
        BioImprime "@Printer.Print " & "+----+----------------------------------------+------------+------+------------+"
    lLinha = lLinha + 1
    End If
    BioImprime "@Printer.Print " & "|COD.|NOME DO PRODUTO                         |VLR.UNITARIO|QUANT.| VLR. TOTAL |"
    lLinha = lLinha + 1
    BioImprime "@Printer.Print " & "+----+----------------------------------------+------------+------+------------+"
    lLinha = lLinha + 1
End Sub
Private Sub ImpCabDespesaCaixa()
    Dim xLinha As String
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+------------------------------------------+---------------+------------+------------+--------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| NOME DO FORNECEDOR                       |   V A L O R   | NUMERO  DO |  DATA  DA  | COMPLEMENTO                                      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                          |               |  DOCUMENTO |   EMISSAO  |                                                  |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------------------------------------+---------------+------------+------------+--------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 4
End Sub
Private Sub ImpCabNotaAbastecimento()
    Dim xLinha As String
    
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|CÓDIGO | RAZÃO SOCIAL                            |   NUMERO   | VALOR DA NOTA |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|CLIENTE|                                         |  DA  NOTA  |    LÍQUIDO    |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------+-----------------------------------------+------------+---------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 4
End Sub
Private Sub cbo_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlhaI.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cboIlhaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoCaixa.SetFocus
    End If
End Sub
Private Sub cboIlhaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlhaF.ListIndex = cboIlhaI.ListIndex
        cboIlhaF.SetFocus
    End If
End Sub
Private Sub cboTipoCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
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
    cbo_funcionario.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = txtDataInicial.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cbo_funcionario.SetFocus
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
            AtivaBotoes (False)
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
            AtivaBotoes (True)
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
    ElseIf Val(cbo_periodo_f.Text) < Val(cbo_periodo_i.Text) Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cboIlhaI.ListIndex = -1 Then
        MsgBox "Selecione uma ilha inicial.", vbInformation, "Atenção!"
        cboIlhaI.SetFocus
    ElseIf cboIlhaF.ListIndex = -1 Then
        MsgBox "Selecione uma ilha final.", vbInformation, "Atenção!"
        cboIlhaF.SetFocus
    ElseIf Val(cboIlhaF.Text) < Val(cboIlhaI.Text) Then
        MsgBox "A ilha final deve ser maior.", vbInformation, "Atenção!"
        cboIlhaF.SetFocus
    ElseIf cboTipoCaixa.ListIndex = -1 Then
        MsgBox "Selecione um tipo de caixa.", vbInformation, "Atenção!"
        cboTipoCaixa.SetFocus
    ElseIf Not ValidaCaixaFechado Then
        txtDataInicial.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Function ValidaCaixaFechado() As Boolean
    ValidaCaixaFechado = True
    If UCase(g_cidade_empresa) Like "*REDEN*" Or UCase(g_cidade_empresa) Like "*CUMAR*" Or UCase(g_cidade_empresa) Like "*CONCEI*" Then
        Exit Function
    End If
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Prioriza Segurança") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            If LocalizarAberturaCaixa(Val(cbo_funcionario.ItemData(cbo_funcionario.ListIndex)), CDate(txtDataInicial.Text), Val(cbo_periodo_i.Text), Val(cboIlhaI.Text), Val(cboTipoCaixa.Text)) Then
                If AberturaCaixa.DataFechamento = CDate("00:00:00") Then
                    MsgBox "Este caixa ainda não está fechado!" & vbCrLf & "Por este motivo não será possível imprimi-lo ou visualiza-lo.", vbOKOnly + vbInformation, "Segurança Ativada!"
                    ValidaCaixaFechado = False
                    If g_nivel_acesso <= 2 Then
                        If (MsgBox("O usuário logado tem privilégio a imprimir o caixa mesmo aberto!" & vbCrLf & "Deseja realmente imprimi-lo?", vbQuestion + vbYesNo + vbDefaultButton1, "Segurança Ativada!")) = vbYes Then
                            ValidaCaixaFechado = True
                        End If
                    End If
                End If
            Else
                MsgBox "Não foi possível localizar a abertura do caixa inicial!", vbOKOnly + vbInformation, "Erro de Integridade!"
                ValidaCaixaFechado = False
            End If
'            If AberturaCaixa.LocalizarCxData(g_empresa, CDate(txtDataInicial.Text), "NF", Val(cbo_periodo_i.Text), Val(cboIlhaI.Text), Val(cboTipoCaixa.Text)) Then
'            Else
'            End If
        End If
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
            AtivaBotoes (False)
            Call GravaAuditoria(1, Me.name, 6, "")
            DoEvents
            Relatorio
            AtivaBotoes (True)
        'End If
        cmd_sair.SetFocus
    End If
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
    Set rs = New ADODB.Recordset
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
    cbo_funcionario.ListIndex = 0
End Sub
Private Sub PreencheCboIlha()
    Dim i As Integer
    
    cboIlhaI.Clear
    cboIlhaF.Clear
    lInverteEncerrante = False
    If Configuracao.LocalizarCodigo(g_empresa) Then
        For i = 1 To Configuracao.QuantidadeIlha
            cboIlhaI.AddItem i
            cboIlhaI.ItemData(cboIlhaI.NewIndex) = i
            cboIlhaF.AddItem i
            cboIlhaF.ItemData(cboIlhaF.NewIndex) = i
            lInverteEncerrante = Configuracao.InverteEncerrantenaPlanilha
            If Configuracao.RelacaoNotasnoCaixa Then
                chkImprimeNotaAbastecimento.Value = 1
            End If
        Next
    End If
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
Private Sub PreencheCboTipoCaixa()
    Dim rstTipoCaixa As ADODB.Recordset
    cboTipoCaixa.Clear
    cboTipoCaixa.AddItem "Todos os Tipos de Caixa"
    Set rstTipoCaixa = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoMovimentoCaixa ORDER BY Codigo")
    Do Until rstTipoCaixa.EOF
        cboTipoCaixa.AddItem rstTipoCaixa!Codigo & " " & rstTipoCaixa!Nome
        cboTipoCaixa.ItemData(cboTipoCaixa.NewIndex) = rstTipoCaixa!Codigo
        rstTipoCaixa.MoveNext
    Loop
    rstTipoCaixa.Close
    Set rstTipoCaixa = Nothing
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
    
    MovimentoBombaMec.NomeTabela = "Movimento_Bomba_Mec"
    If g_nome_usuario = "L.M.C." Then
        Me.Caption = Me.Caption & " - LMC"
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        Me.Caption = Me.Caption & " - ECF"
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
    PreencheCboIlha
    PreencheCboPeriodo
    PreencheCboTipoCaixa
    PreencheCboFuncionario
    lTipoMovimento = 2
    If Not IsDate(txtDataEmissao.Text) Then
        txtDataEmissao.Text = Format(Date, "dd/mm/yyyy")
        txtDataInicial.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        txtDataFinal.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        cboIlhaI.ListIndex = 0
        cboIlhaF.ListIndex = 0
    End If
    lUsarEncerranteMecanico = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Usar Encerrante Mecanico") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lUsarEncerranteMecanico = True
        End If
    End If
    lExecutaActivate = True
End Sub
Private Sub Form_Activate()
    If lExecutaActivate = True Then
        Call GravaAuditoria(1, Me.name, 1, "")
        Screen.MousePointer = 1
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        End If
        If cbo_periodo_i.Enabled = True And cbo_periodo_i.Visible = True Then
            cbo_periodo_i.SetFocus
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
    txtDataFinal.SelLength = 4
    txtDataFinal.MaxLength = 8
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_funcionario.SetFocus
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
    txtDataInicial.SelLength = 4
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
