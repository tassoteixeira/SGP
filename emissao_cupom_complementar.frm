VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_cupom_complementar 
   Caption         =   "Emissão do Cupom Complementar"
   ClientHeight    =   3270
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_cupom_complementar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_cupom_complementar.frx":030A
   ScaleHeight     =   3270
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdImprimiEncerranteAtual 
      Caption         =   "Imprimir Encerrante Atual"
      Height          =   615
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_cupom_complementar.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3240
      Picture         =   "emissao_cupom_complementar.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5160
      Picture         =   "emissao_cupom_complementar.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2280
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chbPorFuncionario 
         Caption         =   "Por Funcionário"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "emissao_cupom_complementar.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar.frx":6CBA
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
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
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_cupom_complementar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lSQL As String
Dim rstMovimentoBomba As New adodb.Recordset
Dim rstMovimentoAbastecimento As New adodb.Recordset
Dim rstEncerranteAtual As New adodb.Recordset
Dim rst As New adodb.Recordset
Dim rst2 As New adodb.Recordset
Dim rst3 As New adodb.Recordset

'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
Dim lNomeArquivoTXT As String
'Fim de variáveis padrão para relatório

Dim BemaRetorno As Integer
Dim lTipoCombustivel(1 To 6) As String
Dim lQtdBombaV(1 To 6) As Currency
Dim lTotalBombaV(1 To 6) As Currency
Dim lQtdAfericaoV(1 To 6) As Currency
Dim lTotalAfericaoV(1 To 6) As Currency
Dim lQtdCupomV(1 To 6) As Currency
Dim lTotalCupomV(1 To 6) As Currency

Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpDaruma As Boolean

Dim lQtdBico As Integer
Dim lPeriodo As Integer
Dim lNumeroCupom As Long
Dim lOrdemCupom As Integer
Dim lDataCupom As Date
Dim lHoraCupom As Date
Dim l_flag_cupom_fiscal As String
Dim lCodigoEcf As Integer
Dim lSerieECF As String
Dim lAutomacao As Boolean
Dim lImprimeCupomComplementar As Boolean
Dim lCaixaIndividual As Boolean
Dim lCodigoFuncionario As Integer
Dim lEcfTruncamento As Boolean
Dim lEcfQtdCasasDecimais As Integer

Private AcertoVendaECF As New cAcertoVendaECF
Private Aliquota As New cAliquota
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private ECF As New cEcf
Private Estoque As New cEstoque
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private MovimentoAbastecimento As New cMovimentoAbastecimento
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
Private MovimentoCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovNotaFiscalSaidaItem As New cMovNotaFiscalSaidaItem
Private PrevisaoVendaPrazo As New cPrevisaoVendaPrazo
Private Produto As New cProduto
Private ReducaoZ As New cReducaoZ
Private Sub AtualizaConstantes()
    Dim dados As String
    Dim i As Integer
    
    lQtdBico = 0
    lPeriodo = 5
    lAutomacao = False
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdBico = Configuracao.QuantidadeBico
        If Mid(Configuracao.OutrasConfiguracoes, 5, 1) = "S" Then
            lAutomacao = True
        End If
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
        lPeriodo = LiberacaoDigitacao.PeriodoInicial
    End If
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lImpQuick = False
    lImpDaruma = False
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    Me.Caption = Me.Caption & " - ECF: " & dados
    If dados = "BEMATECH" Then
        lImpBematech = True
        BemaRetorno = Bematech_FI_FlagsFiscais(i)
        If i <> 0 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro na Rotina: AtualizaConstantes - Bematech_FI_FlagsFiscais(i):" & i)
        End If
    ElseIf dados = "SCHALTER" Then
        lImpSchalter = True
    ElseIf dados = "MECAF" Then
        lImpMecaf = True
    ElseIf dados = "QUICK" Then
        lImpQuick = True
    ElseIf dados = "DARUMA" Then
        lImpDaruma = True
        BemaRetorno = Daruma_FI_VerificaImpressoraLigada
    End If
    lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    lCodigoEcf = 1
    If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
        lCodigoEcf = ECF.Codigo
    End If
    
    lImprimeCupomComplementar = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Imprimir Cupom Complementar") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lImprimeCupomComplementar = True
        End If
    End If
End Sub
Private Sub AtualizaTabelaCupomFiscal(ByVal pNumeroCupom As Long, ByVal pOrdem As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pCodigoProduto As Long, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pCodigoAliquota As Integer, ByVal pLinhaArquivo As String)
    On Error GoTo FileError
    
    MovimentoCupomFiscal.Empresa = g_empresa
    MovimentoCupomFiscal.NumeroCupom = pNumeroCupom
    MovimentoCupomFiscal.Ordem = pOrdem
    MovimentoCupomFiscal.Data = pData
    MovimentoCupomFiscal.Hora = pHora
    MovimentoCupomFiscal.DataCupom = pData
    MovimentoCupomFiscal.Periodo = lPeriodo
    MovimentoCupomFiscal.TipoMovimento = 2 'Pista
    MovimentoCupomFiscal.CodigoCliente = 0
    MovimentoCupomFiscal.CodigoConveniado = 0
    MovimentoCupomFiscal.CodigoProduto = pCodigoProduto
    MovimentoCupomFiscal.ValorUnitario = pValorUnitario
    MovimentoCupomFiscal.Quantidade = pQuantidade
    MovimentoCupomFiscal.ValorTotal = pValorTotal
    MovimentoCupomFiscal.FormaPagamento = 1
    MovimentoCupomFiscal.ValorRecebido = pValorTotal
    MovimentoCupomFiscal.NumeroCheque = ""
    MovimentoCupomFiscal.Telefone = ""
    MovimentoCupomFiscal.operador = 1
    MovimentoCupomFiscal.CupomCancelado = False
    MovimentoCupomFiscal.ItemCancelado = False
    MovimentoCupomFiscal.CodigoAliquota = pCodigoAliquota
    MovimentoCupomFiscal.ValorDesconto = 0
    MovimentoCupomFiscal.Nome = ""
    MovimentoCupomFiscal.CPFCNPJ = ""
    MovimentoCupomFiscal.ValorDescontoEmbutido = 0
    If Produto.LocalizarCodigo(pCodigoProduto) Then
        MovimentoCupomFiscal.TipoCombustivel = Produto.TipoCombustivel
    Else
        MovimentoCupomFiscal.TipoCombustivel = "cc"
    End If
    MovimentoCupomFiscal.CodigoECF = lCodigoEcf
    MovimentoCupomFiscal.CodigoGrupo = Produto.CodigoGrupo
    MovimentoCupomFiscal.TipoSubEstoque = 2 'Pista
    
    If Not MovimentoCupomFiscal.Incluir Then
        Call CriaLogCupom(Time & " - ERRO Emissão do Cupom Complementar: Erro na Rotina: AtualizaTabelaCupomFiscal - pLinhaArquivo:" & pLinhaArquivo)
        MsgBox "Não foi possível incluir registro", vbInformation, "Erro de Integridade!"
    End If
    
    Exit Sub
    
FileError:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro na Rotina: AtualizaTabelaCupomFiscal - pLinhaArquivo:" & pLinhaArquivo)
    MsgBox "Erro Gravando Cupom: " & Error
    Exit Sub
End Sub
Private Sub AtualizaTabelaCupomFiscalItem(ByVal pCodigoAliquota As Integer)
    On Error GoTo FileError
    
    MovimentoCupomFiscalItem.Empresa = g_empresa
    MovimentoCupomFiscalItem.NumeroCupom = MovimentoCupomFiscal.NumeroCupom
    MovimentoCupomFiscalItem.Ordem = MovimentoCupomFiscal.Ordem
    MovimentoCupomFiscalItem.Data = MovimentoCupomFiscal.Data
    MovimentoCupomFiscalItem.CodigoProduto = MovimentoCupomFiscal.CodigoProduto
    MovimentoCupomFiscalItem.ValorUnitario = MovimentoCupomFiscal.ValorUnitario
    MovimentoCupomFiscalItem.Quantidade = MovimentoCupomFiscal.Quantidade
    MovimentoCupomFiscalItem.ValorTotal = MovimentoCupomFiscal.ValorTotal
    MovimentoCupomFiscalItem.ItemCancelado = MovimentoCupomFiscal.ItemCancelado
    MovimentoCupomFiscalItem.ValorDesconto = MovimentoCupomFiscal.ValorDesconto
    MovimentoCupomFiscalItem.ValorAcrescimo = 0
    MovimentoCupomFiscalItem.DescontoEmbutido = False
    MovimentoCupomFiscalItem.Periodo = MovimentoCupomFiscal.Periodo
    MovimentoCupomFiscalItem.TipoCombustivel = MovimentoCupomFiscal.TipoCombustivel
    MovimentoCupomFiscalItem.CodigoECF = MovimentoCupomFiscal.CodigoECF
    MovimentoCupomFiscalItem.CodigoAliquota = pCodigoAliquota
    MovimentoCupomFiscalItem.CodigoGrupo = Produto.CodigoGrupo
    If Not MovimentoCupomFiscalItem.Incluir Then
        Call CriaLogCupom(Time & " - ERRO Emissão do Cupom Complementar: Erro na Rotina: AtualizaTabelaCupomFiscalItem - ECF:" & MovimentoCupomFiscalItem.NumeroCupom)
        MsgBox "Não foi possível incluir registro de item", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub
    
FileError:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro na Rotina: AtualizaTabelaCupomFiscalItem - ECF:" & MovimentoCupomFiscalItem.NumeroCupom)
    MsgBox "Erro Gravando Cupom: " & Error
    Exit Sub
End Sub
Private Sub BuscaNumeroCupom()
    Dim xString As String
    Dim NumeroArquivo As Integer
    Dim xData As String
    Dim xHora As String
    
    On Error GoTo FileError
    
    If lImpBematech Then
        If Not Testa_ImpressoraCF Then
            NumeroArquivo = 99999
        End If
        If l_flag_cupom_fiscal = "F" Then
            l_flag_cupom_fiscal = "A"
            'busca numero do cupom da impressora fiscal
            xString = Space(6)
            BemaRetorno = Bematech_FI_NumeroCupom(xString)
            If BemaRetorno <> 1 Then
                Call AnalizaRetornoBematech(BemaRetorno)
            End If
            lNumeroCupom = CLng(xString) + 1
            lOrdemCupom = 1
        End If
        'busca data/hora da impressora fiscal
        xData = Space(6)
        xHora = Space(6)
        BemaRetorno = Bematech_FI_DataHoraImpressora(xData, xHora)
        lDataCupom = CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
        lHoraCupom = Format(Mid(xHora, 1, 2), "00") & ":" & Format(Mid(xHora, 3, 2), "00") & ":" & Format(Mid(xHora, 5, 2), "00")
    ElseIf lImpQuick Then
        EcfQuickSetaArquivoLog
        EcfQuickObtemNomeLog
        lNumeroCupom = CLng(EcfQuickLeRegistrador("COO", "Long", 5)) + 1
        lOrdemCupom = 1
        lDataCupom = CDate(EcfQuickBuscaData)
        lHoraCupom = CDate(EcfQuickBuscaHora)
    ElseIf lImpDaruma Then
        xString = Space(6)
        BemaRetorno = Daruma_FI_NumeroCupom(xString)
        lNumeroCupom = CLng(xString) ' + 1
        lOrdemCupom = 1
        xData = Space(6)
        xHora = Space(6)
        BemaRetorno = Daruma_FI_DataHoraImpressora(xData, xHora)
        lDataCupom = CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
        lHoraCupom = Format(Mid(xHora, 1, 2), "00") & ":" & Format(Mid(xHora, 3, 2), "00") & ":" & Format(Mid(xHora, 5, 2), "00")
    End If
    Exit Sub
FileError:
    MsgBox "Não foi possível criar o novo cupom fiscal.", vbCritical, "BuscaNumeroCupom"
    Exit Sub
End Sub
Function BuscaUltimaHora() As Date
    If Configuracao.HoraFechamento8 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento8)
    ElseIf Configuracao.HoraFechamento7 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento7)
    ElseIf Configuracao.HoraFechamento6 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento6)
    ElseIf Configuracao.HoraFechamento5 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento5)
    ElseIf Configuracao.HoraFechamento4 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento4)
    ElseIf Configuracao.HoraFechamento3 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento3)
    ElseIf Configuracao.HoraFechamento2 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento2)
    ElseIf Configuracao.HoraFechamento1 <> "00:00:00" Then
        BuscaUltimaHora = fMascaraHora(Configuracao.HoraFechamento1)
    End If
End Function
Private Sub AtivaBotoes(xAtiva As Boolean)
    cmd_visualizar.Enabled = xAtiva
    cmd_imprimir.Enabled = xAtiva
    cmd_sair.Enabled = xAtiva
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rstMovimentoBomba = Nothing
    Set rst = Nothing
    Set rst2 = Nothing
    
    Set AcertoVendaECF = Nothing
    Set Aliquota = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set ECF = Nothing
    Set Estoque = Nothing
    Set LiberacaoDigitacao = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set MovimentoCupomFiscalItem = Nothing
    Set MovimentoAbastecimento = Nothing
    Set MovimentoAfericao = Nothing
    Set MovNotaFiscalSaidaItem = Nothing
    Set PrevisaoVendaPrazo = Nothing
    Set Produto = Nothing
    Set ReducaoZ = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    lTipoCombustivel(1) = "A "
    lTipoCombustivel(2) = "AA"
    lTipoCombustivel(3) = "D "
    lTipoCombustivel(4) = "DA"
    lTipoCombustivel(5) = "G "
    lTipoCombustivel(6) = "GA"
    For i = 1 To 6
        lQtdBombaV(i) = 0
        lTotalBombaV(i) = 0
        lQtdAfericaoV(i) = 0
        lTotalAfericaoV(i) = 0
        lQtdCupomV(i) = 0
        lTotalCupomV(i) = 0
    Next
    l_flag_cupom_fiscal = "F"
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    If lCaixaIndividual And chbPorFuncionario.Value = 1 Then
        'Verifica Movimento_Abastecimento e BaixaAbastecimento
        lSQL = ""
        lSQL = lSQL & "SELECT [Tipo de Combustivel], [Codigo do Produto], Sum(Quantidade) As Quantidade, Sum([Valor Total]) As ValorTotal"
        lSQL = lSQL & "  FROM Movimento_Abastecimento"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
        lSQL = lSQL & " GROUP BY [Tipo de Combustivel], [Codigo do Produto]"
        lSQL = lSQL & ""
        lSQL = lSQL & " UNION"
        lSQL = lSQL & ""
        lSQL = lSQL & " SELECT [Tipo de Combustivel], [Codigo do Produto], Sum(Quantidade) As Quantidade, Sum([Valor Total]) As ValorTotal"
        lSQL = lSQL & "   FROM BaixaAbastecimento"
        lSQL = lSQL & "  WHERE Empresa = " & g_empresa
        lSQL = lSQL & "    AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "    AND Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & "    AND [Codigo do Funcionario] = " & lCodigoFuncionario
        lSQL = lSQL & " GROUP BY [Tipo de Combustivel], [Codigo do Produto]"
        lSQL = lSQL & ""
        lSQL = lSQL & " ORDER BY [Tipo de Combustivel], [Codigo do Produto]"
        Set rstMovimentoAbastecimento = Conectar.RsConexao(lSQL)
        If Not rstMovimentoAbastecimento.EOF Then
            ImpDados
            If g_automacao = True And lLocal = 1 Then
                If Not MovimentoAbastecimento.DescarregarAbastecimentoFuncionario(g_empresa, CDate(msk_data_i.Text), lCodigoEcf, "CP", lCodigoFuncionario) Then
                    MsgBox "Não foi possível descarregar os abastecimentos deste funcionário!", vbInformation, "Erro de Descarregamento"
                End If
            End If
        Else
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Informado ao Usuário que Não Existia abastecimento deste funcionário no Período.")
            MsgBox "Não existe movimento de abastecimento deste funcionário no período.", vbInformation, "Erro de Verificação!"
        End If
        rstMovimentoAbastecimento.Close
    Else
        'Verifica movimento_bomba
        lSQL = ""
        lSQL = lSQL & "SELECT * "
        lSQL = lSQL & "  FROM Movimento_Bomba_Cupom"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
        Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
        If Not rstMovimentoBomba.EOF Then
            ImpDados
            If g_automacao = True And lLocal = 1 Then
                If Not MovimentoAbastecimento.DescarregarAbastecimento(g_empresa, CDate(msk_data_i.Text), lCodigoEcf, "CP") Then
                    MsgBox "Não foi possível descarregar os abastecimentos!", vbInformation, "Erro de Descarregamento"
                End If
            End If
        Else
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Informado ao Usuário que Não Existia Movimento de Bomba Digitado no Período.")
            MsgBox "Não existe movimento de bomba digitada no período.", vbInformation, "Erro de Verificação!"
        End If
        rstMovimentoBomba.Close
    End If
    Call AtivaBotoes(True)
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim i As Integer
    Dim xPrecoCusto As Currency
    Dim xPrecoVenda As Currency
    Dim xPrecoMedio As Currency
    Dim xPrecoUsado As Currency
    Dim xQtd As Currency
    Dim xQtdDiferenca As Currency
    Dim xValor As Currency
    Dim xStrConfiguracaoDiversa As String
    Dim xStrLogAuditoria As String
    Dim xQtdLimite As Currency
    Dim xQtdNaoImpressa As Currency
    
    If lCaixaIndividual And chbPorFuncionario.Value = 1 Then
        TotalizaMovimentoAbastecimento
    Else
        TotalizaMovimentoBombaCupom
    End If
    
    'TotalizaMovimentoCupomFiscal
    TotalizaMovimentoAfericao
    TotalizaAcertoVendaECF
    For i = 1 To 6
        'À Vista
        xPrecoCusto = 0
        xPrecoVenda = 0
        xPrecoMedio = 0
        xPrecoUsado = 0
        
        'aquiaqui
        Call TotalizaMovimentoCupomFiscal2(lTipoCombustivel(i))
        
        'Soma Notas Fiscais de Saidas
        lQtdCupomV(i) = lQtdCupomV(i) + MovNotaFiscalSaidaItem.QuantidadeCombustivelVendaData(g_empresa, lTipoCombustivel(i), CDate(msk_data_i.Text), CDate(msk_data_f.Text), False, False)
        lTotalCupomV(i) = lTotalCupomV(i) + MovNotaFiscalSaidaItem.ValorCombustivelVendaData(g_empresa, lTipoCombustivel(i), CDate(msk_data_i.Text), CDate(msk_data_f.Text), False, False)
        
        lSQL = ""
        lSQL = lSQL & "SELECT [Codigo do Produto], [Tipo de Combustivel]"
        lSQL = lSQL & "  FROM Bomba"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel(i))
        'lSQl = lSQl & "   AND [Tipo de Preco] = " & preparaTexto("V")
        Set rst = Conectar.RsConexao(lSQL)
        If Not rst.EOF Then
            lSQL = ""
            lSQL = lSQL & "SELECT Codigo, Nome, Unidade, [Codigo da Aliquota], [Preco de Custo]"
            lSQL = lSQL & "  FROM Produto"
            lSQL = lSQL & " WHERE Codigo = " & rst![Codigo do Produto]
            Set rst2 = Conectar.RsConexao(lSQL)
            If Not rst2.EOF Then
                If Estoque.LocalizarCodigo(g_empresa, rst2!Codigo) Then
                    xPrecoCusto = rst2![Preco de Custo]
                    xPrecoVenda = Estoque.PrecoVenda
                    xPrecoUsado = Estoque.PrecoVenda
                Else
                    MsgBox "Não foi possível localizar o Preço do Produto!", vbCritical, "Erro Fatal!"
                    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro ao localizar preço de venda do produto: " & rst2!Codigo)
                    Call GravaAuditoria(1, Me.name, 26, "Erro ao localizar preço do produto=" & rst2!Codigo)
                    End
                End If
                lSQL = ""
                lSQL = lSQL & "SELECT [Preco Medio] FROM Combustivel"
                lSQL = lSQL & " WHERE Empresa = " & g_empresa
                lSQL = lSQL & "   AND Codigo = " & preparaTexto(rst![Tipo de Combustivel])
                Set rst3 = Conectar.RsConexao(lSQL)
                If Not rst3.EOF Then
                    xPrecoMedio = rst3![Preco Medio]
                    If xPrecoMedio > xPrecoCusto And xPrecoMedio < xPrecoVenda Then
                        xPrecoUsado = xPrecoMedio
                    End If
                End If
                rst3.Close
                
                'xQtdDiferenca = Qtd Venda Bomba - Qtd Venda Cupom
                xQtdDiferenca = lQtdBombaV(i) - lQtdCupomV(i)
                lTotalBombaV(i) = lQtdBombaV(i) * xPrecoUsado
                
                xStrLogAuditoria = "Tipo de Combustivel: " & lTipoCombustivel(i) & " - Cod.Produto: " & Format(rst2!Codigo, "0000")
                Call CriaLogCupom(xStrLogAuditoria)
                Call GravaAuditoria(1, Me.name, 33, xStrLogAuditoria)
                
                xStrLogAuditoria = "Qtd Bomba: " & lQtdBombaV(i) & " - Qtd Cupom: " & lQtdCupomV(i) & " - Qtd Diferença: " & xQtdDiferenca
                Call CriaLogCupom(xStrLogAuditoria)
                Call GravaAuditoria(1, Me.name, 33, xStrLogAuditoria)
                
                ' Busca Quantidade de Limite para Impressão do Cupom Complementar
                ' Necessária por problemas de erros de encerrantes.
                xQtdLimite = 0
                xStrConfiguracaoDiversa = "CUPOM COMPLEMENTAR PROD:0000 Qtd Maxima"
                Mid(xStrConfiguracaoDiversa, 25, 4) = Format(rst2!Codigo, "0000")
                If ConfiguracaoDiversa.LocalizarCodigo(1, xStrConfiguracaoDiversa) Then
                    xQtdLimite = ConfiguracaoDiversa.Valor
                    If xQtdLimite < xQtdDiferenca Then
                        
                        xQtdNaoImpressa = xQtdDiferenca - xQtdLimite
                        
                        xStrLogAuditoria = "Qtd Diferença passou de: " & xQtdDiferenca
                        xQtdDiferenca = xQtdLimite
                        xStrLogAuditoria = xStrLogAuditoria & " para: " & xQtdDiferenca & " que é o limite configurado. "
                        xStrLogAuditoria = xStrLogAuditoria & " Quantidade impressa: " & xQtdLimite
                        xStrLogAuditoria = xStrLogAuditoria & " Quantidade que não imprimiu: " & xQtdNaoImpressa
                        Call CriaLogCupom(xStrLogAuditoria)
                        Call GravaAuditoria(1, Me.name, 33, xStrLogAuditoria)
                    Else
                        xStrLogAuditoria = "O limite: " & xQtdLimite & " está acima da Diferença: " & xQtdDiferenca
                        Call CriaLogCupom(xStrLogAuditoria)
                        Call GravaAuditoria(1, Me.name, 33, xStrLogAuditoria)
                    End If
                Else
                    xStrLogAuditoria = "*** ATENÇÃO *** Não existe limite configurado."
                    Call CriaLogCupom(xStrLogAuditoria)
                    Call GravaAuditoria(1, Me.name, 33, xStrLogAuditoria)
                End If
                
                
                
                ' Essa previsão de venda está dando ERRO
                ' Como não mais a usamos, vou comentar. 02/07/2016
                'xQtd = Quantidade de Venda a Prazo
'                If PrevisaoVendaPrazo.GravaVendaPrazoDia(g_empresa, lTipoCombustivel(i), CDate(msk_data_i.Text), xQtdDiferenca, lQtdBombaV(i)) Then
'                    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: GravaVendaPrazoDia OK: ")
'                Else
'                    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: ERRO GravaVendaPrazoDia : TipoCombustivel:" & lTipoCombustivel(i) & " - Data:" & msk_data_i.Text & " - Qtd.Diferença:" & xQtdDiferenca & " - Qtd.BombaV:" & lQtdBombaV(i))
'                End If
'                '"Qtd Previsao Venda a Prazo"
'                xQtd = PrevisaoVendaPrazo.TotalVendaPrazoDia(g_empresa, lTipoCombustivel(i), CDate(msk_data_i.Text))
'                If xQtd > 0 Then
'                    xValor = lTotalCupomV(i) / lQtdCupomV(i)
'                    lQtdCupomV(i) = lQtdCupomV(i) + xQtd
'                    lTotalCupomV(i) = lQtdCupomV(i) * xValor
'                Else
'                    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xQTD = 0  TipoCombustivel:" & lTipoCombustivel(i) & " - Data:" & msk_data_i.Text & " - Qtd.Diferença:" & xQtdDiferenca & " - Qtd.BombaV:" & lQtdBombaV(i))
'                End If
                
                
                If lLocal = 1 Then 'Imprimir
                        
                    'na condição anterior (comentada como 'old 10/08/2016 ) mesmo com o limite definido imprimia
                    'os cupons fiscais (caso fossem de
                    'valor acima de 10.000 ate a imprimir todo o volume a ser impresso no dia) com o valor limite
                    'exemplo: faltando imprimir um movimento de 20.000 lts com sistema limitando a 1500 iria imprimir
                    'tres cupons de 1500 lts.pois a cada cupom ele internamente diminui 9500 lts mas imprime o valor
                    'definido como limite
                                                            
                    'If (lQtdBombaV(i) - lQtdCupomV(i)) >= 10000 Then 'old 10/08/2016
                    If xQtdDiferenca >= 10000 Then 'new 10/08/2016
                            
                        Dim xDiferencaQtdTotal As Currency
                        Dim xDiferencaValorTotal As Currency
                        Dim xQtdCupomAImprimir As Currency
                        Dim xValorCupomAImprimir As Currency
                            
                        xDiferencaQtdTotal = lQtdBombaV(i) - lQtdCupomV(i)
                        xDiferencaValorTotal = lTotalBombaV(i) - lTotalCupomV(i)
                            
                        Do Until xDiferencaQtdTotal < 9500
                            xQtdCupomAImprimir = 9500
                            xValorCupomAImprimir = Round(xQtdCupomAImprimir * xPrecoUsado, 2)
                                
                            xDiferencaQtdTotal = xDiferencaQtdTotal - xQtdCupomAImprimir
                            xDiferencaValorTotal = xDiferencaValorTotal - xValorCupomAImprimir
                            Call ImpDet(rst![Codigo do Produto], rst2!Nome, rst2!Unidade, 0, 0, xQtdCupomAImprimir, xValorCupomAImprimir, rst2![Codigo da Aliquota], xPrecoUsado)
                        Loop
                        Call ImpDet(rst![Codigo do Produto], rst2!Nome, rst2!Unidade, 0, 0, xDiferencaQtdTotal, xDiferencaValorTotal, rst2![Codigo da Aliquota], xPrecoUsado)
                    Else
                        Call ImpDet(rst![Codigo do Produto], rst2!Nome, rst2!Unidade, lQtdCupomV(i), lTotalCupomV(i), lQtdBombaV(i), lTotalBombaV(i), rst2![Codigo da Aliquota], xPrecoUsado)
                    End If
                        
                    
                Else
                    Call ImpDet(rst![Codigo do Produto], rst2!Nome, rst2!Unidade, lQtdCupomV(i), lTotalCupomV(i), lQtdBombaV(i), lTotalBombaV(i), rst2![Codigo da Aliquota], xPrecoUsado)
                End If
                
                
            End If
            rst2.Close
        End If
        rst.Close
'        'À Prazo
'        lSQl = ""
'        lSQl = lSQl & "SELECT [Codigo do Produto]"
'        lSQl = lSQl & "  FROM Bomba"
'        lSQl = lSQl & " WHERE [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel(i))
'        lSQl = lSQl & "   AND [Tipo de Preco] = " & preparaTexto("P")
'        Set rst = Conectar.RsConexao(lSQl)
'        If Not rst.EOF Then
'            lSQl = "SELECT Codigo, Nome, Unidade, [Codigo da Aliquota], [Preco de Venda] FROM Produto WHERE Codigo = " & rst![Codigo do Produto]
'            Set rst2 = Conectar.RsConexao(lSQl)
'            If Not rst2.EOF Then
'                Call ImpDet(rst![Codigo do Produto], rst2!Nome, rst2!unidade, lQtdCupomP(i), lTotalCupomP(i), lQtdBombaP(i), lTotalBombaP(i), rst2![Codigo da Aliquota], rst2![Preco de Venda])
'            End If
'            rst2.Close
'        End If
'        rst.Close
    Next
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        If lLocal = 1 Then
            'If (MsgBox("Após a emissão do cupom complementar será impresso a REDUÇÃO Z." & Chr(13) & "E não será mais aceito a emissão de cupom fiscal nesta data." & Chr(13) & Chr(13) & "Deseja realmente imprimir o cupom complementar?", vbQuestion + vbYesNo + vbDefaultButton2, "Emissão do Cupom Complementar")) = 6 Then
            If lImprimeCupomComplementar Then
                ImpCupomComplementar
                Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: O Usuário foi Informado da Conclusão da Impressão")
                MsgBox "Impressão do cupom complementar concluída", vbInformation, "Impressão Concluída!"
            End If
            'End If
        Else
            g_string = lLocal & lNomeArquivo & "|@|Emissão do Cupom Complementar|@|"
            frm_preview.Show 1
        End If
    End If
End Sub
Private Sub ImpDet(xCodigo As Long, xNome As String, xUnidade As String, xQtdCupom As Currency, xValorCupom As Currency, xQtdBomba As Currency, xValorBomba As Currency, xCodigoAliquota As Integer, xPrecoVenda As Currency)
    Dim x_linha As String
    Dim i As Integer
    Dim xQtd As Currency
    Dim xValor As Currency
    Dim xNome_produto As String
    
    xNome_produto = Space(40)
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
        Mid(x_linha, 12, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    x_linha = "|      |                                           |   |          |               |          |               |          |               |"
    i = Len(Format(xCodigo, "##,000"))
    Mid(x_linha, 2 + 6 - i, i) = Format(xCodigo, "##,000")
    Mid(x_linha, 10, 40) = xNome
    Mid(x_linha, 53, 3) = xUnidade
    i = Len(Format(xQtdCupom, "###,##0.00"))
    Mid(x_linha, 57 + 10 - i, i) = Format(xQtdCupom, "###,##0.00")
    i = Len(Format(xValorCupom, "###,###,##0.00"))
    Mid(x_linha, 69 + 14 - i, i) = Format(xValorCupom, "###,###,##0.00")
    i = Len(Format(xQtdBomba, "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(xQtdBomba, "###,###,##0.00")
    i = Len(Format(xValorBomba, "###,###,##0.00"))
    Mid(x_linha, 96 + 14 - i, i) = Format(xValorBomba, "###,###,##0.00")
    
    xQtd = Format(xQtdBomba - xQtdCupom, "000,000,000.000")
    i = Len(Format(xQtd, "##,##0.000"))
    Mid(x_linha, 111 + 10 - i, i) = Format(xQtd, "##,##0.000")
    
    xValor = Format(xQtd * xPrecoVenda, "000,000,000.00")
    i = Len(Format(xValor, "###,###,##0.00"))
    Mid(x_linha, 123 + 14 - i, i) = Format(xValor, "###,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
    If lLocal = 1 Then
        If (xQtdBomba - xQtdCupom) > 0 Then
            i = Len(Trim(xNome))
            Mid(xNome_produto, 1, i) = Trim(xNome)
            x_linha = Format(xCodigo, "00000")
            x_linha = x_linha & xNome_produto
            x_linha = x_linha & Mid(xUnidade, 1, 2)
            x_linha = x_linha & Format(xCodigoAliquota, "00")
            x_linha = x_linha & Format(xQtd, "000000000.000")
            x_linha = x_linha & Format(xPrecoVenda, "0000000000.0000")
            x_linha = x_linha & Format(xValor, "0000000000.00")
            'Print #3, x_linha
            gArquivoTXT.WriteLine (x_linha)
        End If
    End If
End Sub
Private Sub ImprimeEncerrante(ByVal pOrigem As String)
    'Relatório Gerencial
    Dim i As Integer
    Dim i2 As Integer
    Dim xString As String
    Dim xLinha As String
    Dim xValor As Currency
    Dim xQuantidade As Currency
    Dim xValorTotal As Currency
    Dim xUltimoPeriodo As Integer
    
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Início da Impressão dos Encerrantes")
    xValor = 0
    xQuantidade = 0
    xUltimoPeriodo = 0
    
    'Busca último período do movimento_bomba_cupom da data inicial
    lSQL = "SELECT MAX(Periodo) AS Periodo"
    lSQL = lSQL & " FROM Movimento_Bomba_Cupom"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data = " & preparaData(CDate(msk_data_i.Text))
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    If Not rstMovimentoBomba.EOF Then
        rstMovimentoBomba.MoveFirst
        xUltimoPeriodo = rstMovimentoBomba("Periodo").Value
    End If
    Set rstMovimentoBomba = Nothing
    
    
    'lSQL = "SELECT * FROM Movimento_Bomba_Cupom"
    'lSQL = lSQL & " WHERE Empresa = " & g_empresa
    'lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    'lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    'lSQL = lSQL & " ORDER BY Data ASC, Periodo DESC, [Codigo da Bomba] ASC"
    lSQL = "SELECT [Codigo da Bomba], MIN(Abertura) AS Abertura, MAX(Encerrante) AS Encerrante, "
    lSQL = lSQL & " SUM([Quantidade da Saida]) AS [Quantidade da Saida],  SUM([Quantidade da Saida] * [Preco de Venda]) AS Total"
    lSQL = lSQL & " FROM Movimento_Bomba_Cupom"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Periodo = " & xUltimoPeriodo
    lSQL = lSQL & " GROUP BY [Codigo da Bomba]"
    lSQL = lSQL & " ORDER BY [Codigo da Bomba] ASC"
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: lSQL=" & lSQL)
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    If Not rstMovimentoBomba.EOF Then
        rstMovimentoBomba.MoveFirst
        For i = 1 To lQtdBico
            '                  1         2         3         4       4
            '         123456789012345678901234567890123456789012345678
            'xLinha = "1) 1123.123,12 1123.123,12  23,123.12 123.123.12"
            xLinha = "  )       0,00       0,00        0,00       0,00"
            Mid(xLinha, 1, 2) = Format(i, "00")
            xValorTotal = 0
            If Not rstMovimentoBomba.EOF Then
            
                i2 = Len(Format(rstMovimentoBomba("Abertura").Value, "######0.00"))
                Mid(xLinha, 5 + 10 - i2, i2) = Format(rstMovimentoBomba("Abertura").Value, "######0.00")
                
                i2 = Len(Format(rstMovimentoBomba("Encerrante").Value, "######0.00"))
                Mid(xLinha, 16 + 10 - i2, i2) = Format(rstMovimentoBomba("Encerrante").Value, "######0.00")
                i2 = Len(Format(rstMovimentoBomba("Quantidade da Saida").Value, "##,##0.00"))
                Mid(xLinha, 29 + 9 - i2, i2) = Format(rstMovimentoBomba("Quantidade da Saida").Value, "##,##0.00")
                'xValorTotal = Format(rstMovimentoBomba("Quantidade da Saida").Value * rstMovimentoBomba("Preco de Venda").Value, "0000000000.00")
                xValorTotal = rstMovimentoBomba("Total").Value
                i2 = Len(Format(xValorTotal, "###,##0.00"))
                Mid(xLinha, 39 + 10 - i2, i2) = Format(xValorTotal, "###,##0.00")
                xValor = xValor + xValorTotal
                xQuantidade = xQuantidade + rstMovimentoBomba("Quantidade da Saida").Value
            End If
            If xValorTotal > 0 Then
                xString = xString & xLinha
            End If
            rstMovimentoBomba.MoveNext
        Next
    End If
    rstMovimentoBomba.Close
    xLinha = "          TOTAL DO CAIXA:        0,00       0,00"
    Mid(xLinha, 1, 1) = pOrigem
    i2 = Len(Format(xQuantidade, "###,##0.00"))
    Mid(xLinha, 28 + 10 - i2, i2) = Format(xQuantidade, "###,##0.00")
    i2 = Len(Format(xValor, "###,##0.00"))
    Mid(xLinha, 39 + 10 - i2, i2) = Format(xValor, "###,##0.00")
    xString = xString & xLinha
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: TOTAL=" & xLinha)
    
    'Abertura do Relatório Gerencial
    If lImpBematech Then
        If Len(xString) <= 618 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
            'Abre Relatorio Gerencial
            BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
            'Fechamento de Relatório Gerencial
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
        Else
            BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1, 576))
            If Len(xString) <= 1152 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, Len(xString) - 576))
            ElseIf Len(xString) <= 1728 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1153, Len(xString) - 1152))
            ElseIf Len(xString) <= 2304 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1153, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1729, Len(xString) - 1728))
            End If
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
    ElseIf lImpQuick Then
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
        'Abre Relatorio Gerencial
        If EcfQuickDefineGerencial(0, "Gerencial") Then
            BemaRetorno = 1
            If EcfQuickAbreGerencial(0, "Gerencial") Then
                BemaRetorno = 1
            Else
                BemaRetorno = 0
            End If
        Else
            BemaRetorno = 0
        End If
        'Imprime detalhes do relatorio gerencial
        If Len(xString) <= 618 Then
            If EcfQuickImprimeTexto(xString) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        Else
            If EcfQuickImprimeTexto(Mid(xString, 1, 576)) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
            
            If EcfQuickImprimeTexto(Mid(xString, 577, Len(xString) - 576)) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        End If
        'Fecha Relatorio Gerencial
        If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
            BemaRetorno = 1
        Else
            BemaRetorno = -1
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: EcfQuickEncerraDocumento=" & BemaRetorno)
    ElseIf lImpDaruma Then
        If Len(xString) <= 618 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
            'Abre Relatorio Gerencial
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
            
            BemaRetorno = Daruma_FI_RelatorioGerencial(xString)
            
            'Fechamento de Relatório Gerencial
            BemaRetorno = Daruma_FI_FechaRelatorioGerencial()
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
        Else
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
            BemaRetorno = Daruma_FI_RelatorioGerencial(Mid(xString, 1, 576))
            BemaRetorno = Daruma_FI_RelatorioGerencial(Mid(xString, 577, Len(xString) - 576))
            BemaRetorno = Daruma_FI_FechaRelatorioGerencial()
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
    End If
    
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Concluído a Impressão dos Encerrantes")
End Sub
Private Sub ImprimeEncerranteAutomacaoAtual(ByVal pOrigem As String)
    'Relatório Gerencial
    Dim i As Integer
    Dim i2 As Integer
    Dim xString As String
    Dim xString2 As String
    Dim xLinha As String
    Dim xLinhaCabecalho As String
    Dim xValor As Currency
    Dim xQuantidade As Currency
    Dim xValorTotal As Currency
    Dim xUltimoPeriodo As Integer
    
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Início da Impressão dos Encerrantes")
    xValor = 0
    xQuantidade = 0
    xUltimoPeriodo = 0
    
    'Busca último período do movimento_bomba_cupom da data inicial
'    lSQL = "SELECT MAX(Periodo) AS Periodo"
'    lSQL = lSQL & " FROM Movimento_Bomba_Cupom"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & " AND Data = " & preparaData(CDate(msk_data_i.Text))
'    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
'    If Not rstMovimentoBomba.EOF Then
'        rstMovimentoBomba.MoveFirst
'        xUltimoPeriodo = rstMovimentoBomba("Periodo").Value
'    End If
'    Set rstMovimentoBomba = Nothing
'
    
    'lSQL = "SELECT * FROM Movimento_Bomba_Cupom"
    'lSQL = lSQL & " WHERE Empresa = " & g_empresa
    'lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    'lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    'lSQL = lSQL & " ORDER BY Data ASC, Periodo DESC, [Codigo da Bomba] ASC"
    
'    lSQL = "SELECT [Codigo da Bomba], MIN(Abertura) AS Abertura, MAX(Encerrante) AS Encerrante, "
'    lSQL = lSQL & " SUM([Quantidade da Saida]) AS [Quantidade da Saida],  SUM([Quantidade da Saida] * [Preco de Venda]) AS Total"
'    lSQL = lSQL & " FROM Movimento_Bomba_Cupom"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
'    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
'    lSQL = lSQL & " AND Periodo = " & xUltimoPeriodo
'    lSQL = lSQL & " GROUP BY [Codigo da Bomba]"
'    lSQL = lSQL & " ORDER BY [Codigo da Bomba] ASC"
'    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: lSQL=" & lSQL)
'    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
'
    lSQL = "SELECT [Codigo da Bomba], Encerrante "
    lSQL = lSQL & " FROM Encerrante_Atual"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " GROUP BY [Codigo da Bomba], Encerrante"
    lSQL = lSQL & " ORDER BY [Codigo da Bomba] ASC, Encerrante"
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: lSQL=" & lSQL)
    Set rstEncerranteAtual = Conectar.RsConexao(lSQL)
    If Not rstEncerranteAtual.EOF Then
        rstEncerranteAtual.MoveFirst
        For i = 1 To lQtdBico
            '                   1         2         3         4       4
            '          123456789012345678901234567890123456789012345678
            'xLinha = "1) 1123.123,12 1123.123,12  23,123.12 123.123.12"
            xLinha = "  )                                              "
            Mid(xLinha, 1, 2) = Format(i, "00")
            xValorTotal = 0
            '                            1         2         3         4
            '                   123456789012345678901234567890123456789012345678
            '                  "  )                                             "
            xLinhaCabecalho = "Bico: Encerrante:                                "
            xString = xLinhaCabecalho
            If Not rstEncerranteAtual.EOF Then
'                i2 = rstEncerranteAtual("Codigo da Bomba")
'                Mid(xLinha, 5 + 10 - i2, i2) = rstEncerranteAtual("Codigo da Bomba")
                i2 = Len(Format(rstEncerranteAtual("Encerrante").Value, "######0.00"))
                Mid(xLinha, 5 + 11 - i2, i2) = Format(rstEncerranteAtual("Encerrante").Value, "######0.00")
'                i2 = Len(Format(rstMovimentoBomba("Abertura").Value, "######0.00"))
'                Mid(xLinha, 5 + 10 - i2, i2) = Format(rstMovimentoBomba("Abertura").Value, "######0.00")
'                i2 = Len(Format(rstMovimentoBomba("Encerrante").Value, "######0.00"))
'                Mid(xLinha, 16 + 10 - i2, i2) = Format(rstMovimentoBomba("Encerrante").Value, "######0.00")
'                i2 = Len(Format(rstMovimentoBomba("Quantidade da Saida").Value, "##,##0.00"))
'                Mid(xLinha, 29 + 9 - i2, i2) = Format(rstMovimentoBomba("Quantidade da Saida").Value, "##,##0.00")
'                'xValorTotal = Format(rstMovimentoBomba("Quantidade da Saida").Value * rstMovimentoBomba("Preco de Venda").Value, "0000000000.00")
'                xValorTotal = rstMovimentoBomba("Total").Value
'                i2 = Len(Format(xValorTotal, "###,##0.00"))
'                Mid(xLinha, 39 + 10 - i2, i2) = Format(xValorTotal, "###,##0.00")
'                xValor = xValor + xValorTotal
'                xQuantidade = xQuantidade + rstMovimentoBomba("Quantidade da Saida").Value
            End If
            'If xValorTotal > 0 Then
                xString2 = xString2 & xLinha
            'End If
            rstEncerranteAtual.MoveNext
        Next
    End If
    rstEncerranteAtual.Close
    xString = xString + xString2
'    xLinha = "          TOTAL DO CAIXA:        0,00       0,00"
'    Mid(xLinha, 1, 1) = pOrigem
'    i2 = Len(Format(xQuantidade, "###,##0.00"))
'    Mid(xLinha, 28 + 10 - i2, i2) = Format(xQuantidade, "###,##0.00")
'    i2 = Len(Format(xValor, "###,##0.00"))
'    Mid(xLinha, 39 + 10 - i2, i2) = Format(xValor, "###,##0.00")
'    xString = xString & xLinha
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: TOTAL=" & xLinha)
    
    'Abertura do Relatório Gerencial
    If lImpBematech Then
        If Len(xString) <= 618 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
            'Abre Relatorio Gerencial
            BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
            'Fechamento de Relatório Gerencial
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
        Else
            BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1, 576))
            If Len(xString) <= 1152 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, Len(xString) - 576))
            ElseIf Len(xString) <= 1728 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1153, Len(xString) - 1152))
            ElseIf Len(xString) <= 2304 Then
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 577, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1153, 576))
                BemaRetorno = Bematech_FI_RelatorioGerencial(Mid(xString, 1729, Len(xString) - 1728))
            End If
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
    ElseIf lImpQuick Then
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
        'Abre Relatorio Gerencial
        If EcfQuickDefineGerencial(0, "Gerencial") Then
            BemaRetorno = 1
            If EcfQuickAbreGerencial(0, "Gerencial") Then
                BemaRetorno = 1
            Else
                BemaRetorno = 0
            End If
        Else
            BemaRetorno = 0
        End If
        'Imprime detalhes do relatorio gerencial
        If Len(xString) <= 618 Then
            If EcfQuickImprimeTexto(xString) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        Else
            If EcfQuickImprimeTexto(Mid(xString, 1, 576)) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
            
            If EcfQuickImprimeTexto(Mid(xString, 577, Len(xString) - 576)) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        End If
        'Fecha Relatorio Gerencial
        If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
            BemaRetorno = 1
        Else
            BemaRetorno = -1
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: EcfQuickEncerraDocumento=" & BemaRetorno)
    ElseIf lImpDaruma Then
        If Len(xString) <= 618 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: xString=" & xString)
            'Abre Relatorio Gerencial
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
            
            BemaRetorno = Daruma_FI_RelatorioGerencial(xString)
            
            'Fechamento de Relatório Gerencial
            BemaRetorno = Daruma_FI_FechaRelatorioGerencial()
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
        Else
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
            BemaRetorno = Daruma_FI_RelatorioGerencial(Mid(xString, 1, 576))
            BemaRetorno = Daruma_FI_RelatorioGerencial(Mid(xString, 577, Len(xString) - 576))
            BemaRetorno = Daruma_FI_FechaRelatorioGerencial()
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Bematech_FI_RelatorioGerencial=" & BemaRetorno)
    End If
    
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Concluído a Impressão dos Encerrantes")
End Sub
Private Sub ImprimeReducaoZ()
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    
    Call GravaAuditoria(1, Me.name, 26, "Imprime Reducao Z. lImpBematech=" & lImpBematech)
    If lImpBematech Then
        xData = Format(Date, "dd/mm/yyyy")
        xHora = Format(Time, "hh:mm:ss")
        Call GravaAuditoria(1, Me.name, 26, "xData=" & xData & " xHora=" & xHora)
        BemaRetorno = Bematech_FI_ReducaoZ(xData, xHora)
        Call GravaAuditoria(1, Me.name, 26, "Imprime Reducao Z. - BemaRetorno=" & BemaRetorno)
        If BemaRetorno = -1 Then
            xData = Format(Date, "ddmmyy")
            xHora = Format(Time, "hhmmss")
            Call GravaAuditoria(1, Me.name, 26, "xData=" & xData & " xHora=" & xHora)
            BemaRetorno = Bematech_FI_ReducaoZ(xData, xHora)
            Call GravaAuditoria(1, Me.name, 26, "Imprime Reducao Z. - BemaRetorno=" & BemaRetorno)
        End If
    ElseIf lImpSchalter Then
        Retorno = ecfReducaoZ("caixa")
    ElseIf lImpMecaf Then
        xRetorno = ReducaoZ(Asc("0"))
        Sleep 25000
    ElseIf lImpDaruma Then
        xData = Format(Date, "dd/mm/yyyy")
        xHora = Format(Time, "hh:mm:ss")
        Call GravaAuditoria(1, Me.name, 26, "xData=" & xData & " xHora=" & xHora)
        BemaRetorno = Daruma_FI_ReducaoZAjustaDataHora(xData, xHora)
        Call GravaAuditoria(1, Me.name, 26, "Imprime Reducao Z. - BemaRetorno=" & BemaRetorno)
        If BemaRetorno = -1 Then
            xData = Format(Date, "ddmmyy")
            xHora = Format(Time, "hhmmss")
            Call GravaAuditoria(1, Me.name, 26, "xData=" & xData & " xHora=" & xHora)
            BemaRetorno = Daruma_FI_ReducaoZAjustaDataHora(xData, xHora)
            Call GravaAuditoria(1, Me.name, 26, "Imprime Reducao Z. - BemaRetorno=" & BemaRetorno)
        End If
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    
    If lLocal = 1 Then
        'Print #3, "FIM"
        gArquivoTXT.WriteLine ("FIM")
    End If
    
    For i = 2 To 6
        lQtdBombaV(1) = lQtdBombaV(1) + lQtdBombaV(i)
        lTotalBombaV(1) = lTotalBombaV(1) + lTotalBombaV(i)
        lQtdCupomV(1) = lQtdCupomV(1) + lQtdCupomV(i)
        lTotalCupomV(1) = lTotalCupomV(1) + lTotalCupomV(i)
    Next
    
        lQtdBombaV(1) = lQtdBombaV(1)
        lTotalBombaV(1) = lTotalBombaV(1)
        lQtdCupomV(1) = lQtdCupomV(1)
        lTotalCupomV(1) = lTotalCupomV(1)
    
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                               *** TOTAL DO RELATORIO |          |               |          |               |          |               |"
    i = Len(Format(lQtdCupomV(1), "###,##0.00"))
    Mid(x_linha, 57 + 10 - i, i) = Format(lQtdCupomV(1), "###,##0.00")
    i = Len(Format(lTotalCupomV(1), "###,###,##0.00"))
    Mid(x_linha, 69 + 14 - i, i) = Format(lTotalCupomV(1), "###,###,##0.00")
    i = Len(Format(lQtdBombaV(1), "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(lQtdBombaV(1), "###,###,##0.00")
    i = Len(Format(lTotalBombaV(1), "###,###,##0.00"))
    Mid(x_linha, 96 + 14 - i, i) = Format(lTotalBombaV(1), "###,###,##0.00")
    
    
    i = Len(Format(lQtdBombaV(1) - lQtdCupomV(1), "##,##0.000"))
    Mid(x_linha, 111 + 10 - i, i) = Format(lQtdBombaV(1) - lQtdCupomV(1), "##,##0.000")
    
    i = Len(Format(lTotalBombaV(1) - lTotalCupomV(1), "###,###,##0.00"))
    Mid(x_linha, 123 + 14 - i, i) = Format(lTotalBombaV(1) - lTotalCupomV(1), "###,###,##0.00")
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print x_linha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+------------------------------------------------------+----------+---------------+----------+---------------+----------+---------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                                                                           Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| CUPOM COMPLEMENTAR DO PERIODO: __/__/____ A __/__/____.                                                           Goiânia, __/__/____ |"
    Mid(x_linha, 34, 10) = msk_data_i.Text
    Mid(x_linha, 47, 10) = msk_data_f.Text
    Mid(x_linha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    x_linha = "+------+-------------------------------------------+---+--------------------------+--------------------------+--------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|CODIGO|                                           |   | C U P O M    F I S C A L |        V E N D A S       |   CUPOM   COMPLEMENTAR   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|  DO  | DISCRIMINAÇÃO DOS PRODUTOS                |UN.+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PROD.|                                           |   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|      |                                           |   |          |               |          |               |          |               |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpCupomComplementar()
    Dim x_linha As String
    Dim xString As String
    Dim xString2 As String
    Dim xDescricao As String
    Dim i As Integer
    
    Dim CodigoProduto As String
    Dim NomeProduto As String
    Dim xAliquota As String
    Dim Quantidade As String
    Dim Valor As String
    Dim ValorDesconto As String
    Dim ValorAcrescimo As String
    Dim Departamento As String
    Dim Un As String
    
    Dim x_valor_acrescimo As Currency
    Dim x_valor_desconto As Currency
    Dim x_total As Currency
    Dim xValorUnitario As String * 9
    Dim xValorUnitario2 As Currency
    Dim xValorTotal As Currency
    Dim xQuantidade As String * 7
    Dim xQuantidade2 As Currency
    Dim xCodigoProduto As Long
    Dim xCodigoAliquota As Integer
    Dim xCodigoFiscal As String * 2
    Dim xStringEmail As String
    
    Dim xTruncaValor As Double
    Dim xTruncaQuantidade As Double
    Dim xTruncaTotalCalculado As Currency
    
    On Error GoTo ErroImpCupomComplementar
    
    'Close #3
    gArquivoTXT.Close
    'Open "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT" For Input As #3
    Set gArquivoTXT = gArqTxt.OpenTextFile(lNomeArquivoTXT, ForReading)
    'Do Until EOF(3)
    xStringEmail = "Iniciada a impressão em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & "   " & vbCrLf
    Do Until gArquivoTXT.AtEndOfStream
        'Line Input #3, x_linha
        x_linha = gArquivoTXT.ReadLine
        If Mid(x_linha, 1, 3) = "FIM" Then
            Exit Do
        End If
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Acionado a Emissão do C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
        Call BuscaNumeroCupom
        x_total = 0
        
        'Busca se ECF está truncando ou Não
        VerificaSeEcfTruncamento
        
        'Abre o cupom fiscal
        If lImpBematech Then
            BemaRetorno = Bematech_FI_AbreCupom("")
        ElseIf lImpQuick Then
            If EcfQuickAbreCupomFiscal("", "", "") Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        ElseIf lImpDaruma Then
            BemaRetorno = Daruma_FI_AbreCupom("")
        End If
        'Codigo da Aliquota
        'xCodigoAliquota = Mid(x_linha, 48, 2)
        'Venda de Item com entrada de departamento,
        'Verifica se há diferença do total
        xString = Format(Format(fValidaValor(Mid(x_linha, 63, 15)) * fValidaValor(Mid(x_linha, 50, 13)), "###,##0.0000"), "###,##0.0000")
        i = Len(xString)
        xString = Mid(xString, 1, i - 2)
        x_valor_acrescimo = 0
        x_valor_desconto = 0
        If fValidaValor(Mid(x_linha, 78, 13)) > fValidaValor(xString) Then
            x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - fValidaValor(xString)
        ElseIf fValidaValor(Mid(x_linha, 78, 13)) < fValidaValor(xString) Then
            x_valor_desconto = fValidaValor(xString) - fValidaValor(Mid(x_linha, 78, 13))
        Else
        End If
        
        
        'código do produto
        xCodigoProduto = Mid(x_linha, 1, 5)
        CodigoProduto = Format(Mid(x_linha, 1, 5), "####0")
        'nome do produto
        NomeProduto = Mid(x_linha, 6, 40)
        xStringEmail = xStringEmail & Mid(NomeProduto, 1, 30)
        'tipo de tributação
        xCodigoAliquota = Mid(x_linha, 48, 2)
        lSQL = "SELECT [Codigo Fiscal] FROM Aliquota WHERE Codigo = " & xCodigoAliquota
        Set rst = Conectar.RsConexao(lSQL)
        If Not rst.EOF Then
            xAliquota = rst![Codigo Fiscal]
        Else
            xAliquota = "II"
        End If
        If Produto.LocalizarCodigo(CodigoProduto) Then
            If Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                xCodigoAliquota = Aliquota.Codigo
                xAliquota = Aliquota.CodigoFiscal
            Else
                Call CriaLogCupom(Time & " - ERRO Cupom Complementar: Aliquota não encontrada=" & Produto.CodigoAliquota & " -SerieECF=" & lSerieECF)
            End If
        Else
            Call CriaLogCupom(Time & " - ERRO Cupom Complementar: Produto não encontrada=" & CodigoProduto)
        End If
        rst.Close
        'Valor Unitário
        xString = Format(Mid(x_linha, 63, 15), "000000.000")
        xStringEmail = xStringEmail & " Vlr:" & xString
        Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
        xValorUnitario2 = xString
        'Quantidade
        'xString = Format(Mid(x_linha, 50, 13), "0000.000")
        'Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
        Quantidade = Mid(x_linha, 55, 4) + Mid(x_linha, 60, 3)
        xStringEmail = xStringEmail & " Qtd:" & Quantidade
        xQuantidade2 = Format(Mid(x_linha, 50, 13), "0000.000")
        'Valor do Acréscimo
        xString = Format(x_valor_acrescimo, "00000000.00")
        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        'Valor do Desconto
        xString = Format(x_valor_desconto, "00000000.00")
        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        
        'Desconsidera Descontos ou Acréscimos
        If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
            x_valor_acrescimo = 0
            x_valor_desconto = 0
            ValorAcrescimo = "0000000000"
            ValorDesconto = "0000000000"
        End If
        
        
        'Departamento
        Departamento = Format(1, "00")
        'Unidade de Medida
        Un = Mid(x_linha, 46, 2)
        If lImpBematech Then
            
            'Tratamento de Truncamento
            If lEcfTruncamento = True Then
                xTruncaValor = Format(Mid(x_linha, 63, 15), "000000.0000")
                If lEcfQtdCasasDecimais = 2 Then
                    xTruncaQuantidade = Format(Mid(x_linha, 50, 12), "0000.000")
                Else
                    xTruncaQuantidade = Format(Mid(x_linha, 50, 13), "0000.000")
                End If
                xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
                ValorAcrescimo = "0000000000"
                ValorDesconto = "0000000000"
                If fValidaValor(Mid(x_linha, 78, 13)) > xTruncaTotalCalculado Then
                    x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - xTruncaTotalCalculado
                    Call CriaLogCupom("Acrescimo Truncamento  valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    xString = Format(x_valor_acrescimo, "00000000.00")
                    ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                ElseIf fValidaValor(Mid(x_linha, 78, 13)) < xTruncaTotalCalculado Then
                    x_valor_desconto = xTruncaTotalCalculado - fValidaValor(Mid(x_linha, 78, 13))
                    Call CriaLogCupom("Desconto Truncamento   valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    xString = Format(x_valor_desconto, "00000000.00")
                    ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                End If
            End If
            
            BemaRetorno = Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
            If BemaRetorno <> 1 Then
                Call AnalizaRetornoBematech(BemaRetorno)
            End If
        ElseIf lImpQuick Then
            Call EcfQuickVendeItem(True, -2, 0, CodigoProduto, "", Trim(NomeProduto), 0, (CCur(Valor) / 1000), (CCur(Quantidade) / 1000), Produto.Unidade)
        ElseIf lImpDaruma Then
            'Valor Unitário
            Valor = Format(Mid(x_linha, 64, 15), "000000.000")
            'Quantidade
            Quantidade = Format(Mid(x_linha, 51, 13), "0000.000")
            'Valor do Acréscimo
            ValorAcrescimo = Format(x_valor_acrescimo, "00000000.00")
            'Valor do Desconto
            ValorDesconto = Format(x_valor_desconto, "00000000.00")
            
            'Tratamento de Truncamento
            If lEcfTruncamento = True Then
                xTruncaValor = Format(Mid(x_linha, 63, 15), "000000.0000")
                If lEcfQtdCasasDecimais = 2 Then
                    xTruncaQuantidade = Format(Mid(x_linha, 50, 12), "0000.000")
                Else
                    xTruncaQuantidade = Format(Mid(x_linha, 50, 13), "0000.000")
                End If
                xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
                ValorAcrescimo = "0000000000"
                ValorDesconto = "0000000000"
                If fValidaValor(Mid(x_linha, 78, 13)) > xTruncaTotalCalculado Then
                    x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - xTruncaTotalCalculado
                    Call CriaLogCupom("Acrescimo Truncamento  valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    xString = Format(x_valor_acrescimo, "00000000.00")
                    ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                ElseIf fValidaValor(Mid(x_linha, 78, 13)) < xTruncaTotalCalculado Then
                    x_valor_desconto = xTruncaTotalCalculado - fValidaValor(Mid(x_linha, 78, 13))
                    Call CriaLogCupom("Desconto Truncamento   valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    xString = Format(x_valor_desconto, "00000000.00")
                    ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                End If
            End If
            
            BemaRetorno = Daruma_FI_VendeItem(CodigoProduto, NomeProduto, xAliquota, Un, Quantidade, 3, Valor, "$", ValorDesconto)
        End If

        'Grava Cupom Complementar
        xValorTotal = Format(xValorUnitario2 * xQuantidade2, "0000000000.00") - x_valor_desconto + x_valor_acrescimo
        Call AtualizaTabelaCupomFiscal(lNumeroCupom, lOrdemCupom, lDataCupom, lHoraCupom, xCodigoProduto, xValorUnitario2, xQuantidade2, xValorTotal, xCodigoAliquota, x_linha)
        Call AtualizaTabelaCupomFiscalItem(xCodigoAliquota)
        
        'Desconto para o Cupom Fiscal
        xString = Mid(Format(fValidaValor(0), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(0), "000000000000.00"), 14, 2)
        If lImpBematech Then
            BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
        ElseIf lImpQuick Then
            'estudar e testar desconto ou acrescimo
            'Valor do Acréscimo/Desconto
            'If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
            '    Call EcfQuickAcresceItemFiscal(lOrdem, False, 0, (CCur(xString) / 1000))
            'End If
        ElseIf lImpDaruma Then
            xString = Format(fValidaValor(0), "000000000000.00")
            BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", xString)
        End If
        
        'Efetua Forma de Pagamento
        xString = "Dinheiro        "
        xString2 = Mid(Format(xValorTotal, "000000000000.00"), 1, 12) + Mid(Format(xValorTotal, "000000000000.00"), 14, 2)
        xDescricao = ""
        If lImpBematech Then
            BemaRetorno = Bematech_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
            'Fecha Cupom Fiscal
            xString = "Cerrado Informatica - (062) 3277-1017           Sistemas para Automacao Comercial               "
            BemaRetorno = Bematech_FI_TerminaFechamentoCupom(xString)
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Impresso o C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
        ElseIf lImpQuick Then
            If EcfQuickPagaCupom(0, Trim(xString), "", (CCur(xString2) / 100)) Then
                If EcfQuickEncerraDocumento(g_nome_usuario, "Cerrado Informatica (62) 3277-1017") Then
                    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Impresso o C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
                Else
                    MsgBox "Erro ao finalizar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
                End If
            Else
                MsgBox "Erro ao pagar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        ElseIf lImpDaruma Then
            xString2 = Format(xValorTotal, "000000000000.00")
            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
            'Fecha Cupom Fiscal
            xString = "Cerrado Informatica - (062) 3277-1017           Sistemas para Automacao Comercial               "
            BemaRetorno = Daruma_FI_TerminaFechamentoCupom(xString)
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Impresso o C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
        End If
        l_flag_cupom_fiscal = "F"
        xStringEmail = xStringEmail & vbCrLf
    Loop
    
    'Testa se tem Automação
    'Quando for necessário enviar email
    'Retirar o comentário logo abaixo
    If Mid(Configuracao.OutrasConfiguracoes, 5, 1) = "S" Then
        xStringEmail = xStringEmail & "Finalizado a impressão em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS")
        'Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Cupom Complementar!", xStringEmail, True, gNumeroEmailInicial)
    End If
    
    Exit Sub
ErroImpCupomComplementar:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro ImpCupomComplementar - " & x_linha)
    Exit Sub
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cmd_visualizar.SetFocus
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
    On Error GoTo ErroImprimir
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Pedido a Impressão do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    
    'Cria o Arquivo CUPOM_COMPLEMENTAR.TXT
    lNomeArquivoTXT = "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT"
    Set gArquivoTXT = gArqTxt.CreateTextFile(lNomeArquivoTXT, True)
    
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "Data Inicial:" & msk_data_i.Text & " Data Final:" & msk_data_f.Text)
            Call AtivaBotoes(False)
            g_string = "imprimiu|@|"
            Relatorio
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Impresso")
        End If
    End If
    If ValidaCampos Then
        'Sleep 10000
        If TestaImprimeEncerrante Then
            ImprimeEncerrante ("I")
        End If
        Call GravaAuditoria(1, Me.name, 26, "Imprime reducao Z Automatica=" & Configuracao.ImprimirReducaoZ)
        If Configuracao.ImprimirReducaoZ = True Then
            Call GravaAuditoria(1, Me.name, 26, "Hora Atual=" & Time & " BuscaUltimaHora=" & BuscaUltimaHora)
            If Time >= BuscaUltimaHora And SeImprimeReducaoZ = True Then
                Call GravaAuditoria(1, Me.name, 26, "Existe Reducao Z ?=" & ReducaoZ.LocalizarCodigo(Date))
                'Codigo comentado pelo motivo de um estabelecimento
                'Ter mais de 1 ECF, e a Tabela ReducaoZ, ser projetada
                'Para apenas 1 ECF
                'If Not ReducaoZ.LocalizarCodigo(Date) Then
                    Call CriaLogCupom(Time & " - Emissao Cupom Complementar: Foi Acionado a Emissão da Redução Z Automaticamente.")
                    Call ImprimeReducaoZ
                    Call CriaLogCupom(Time & " - Emissao Cupom Complementar: Foi Finalizado a Emissão da Redução Z.")
                    ReducaoZ.Data = Date
                    If Not ReducaoZ.Incluir Then
                        MsgBox "Não foi possível incluir ReducaoZ", vbInformation, "Erro de Integridade!"
                    End If
                'End If
                If lAutomacao Then
                    MsgBox "Só poderá ser tirado cupom fiscal após as 00:00 horas.", vbInformation, "Atenção!"
                    Call CriaLogCupom(Time & " - Emissao Cupom Complementar: Só poderá ser tirado cupom fiscal após as 00:00 horas.")
                Else
                    Call CriaLogCupom(Time & " - Emissao Cupom Complementar: O Usuário Está Sendo Informado do Fechamento Automatico do SGP.")
                    MsgBox "Este programa será fechado e somente funcionará após as 00:00 horas.", vbInformation, "Fechamento para Reprogramação"
                    Call CriaLogCupom(Time & " - Emissao Cupom Complementar: O Sistema Gerenciador de Posto Está Sendo Fechado Automaticamente.")
                    End
                End If
            End If
        End If
    End If
    gArquivoTXT.Close
    Exit Sub
ErroImprimir:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro ao abrir o arquivo CUPOM_COMPLEMENTAR.TXT")
    Call AtivaBotoes(True)
    Exit Sub
End Sub
Private Function SeImprimeReducaoZ() As Boolean
    ' Se empresa = LG AUTO POSTO
    If g_nome_empresa Like "*LG AUTO POSTO*" Or g_nome_empresa Like "*TEIXEIRA E PINHEIRO LTDA*" Then
        SeImprimeReducaoZ = False
        'If LiberacaoDigitacao.PeriodoInicial = 2 Then
        If LiberacaoDigitacao.PeriodoInicial = 3 Then
            Dim xDiaSemana As Integer
            xDiaSemana = Weekday(LiberacaoDigitacao.DataInicial) '1-Domingo, 2-Segunda ... 6-Sexta, 7-Sabado
            ' Segunda a Quinta, vai ter 2 periodos
            If xDiaSemana >= 2 And xDiaSemana <= 5 Then
                SeImprimeReducaoZ = True
            End If
        'ElseIf LiberacaoDigitacao.PeriodoInicial = 3 Then
        ElseIf LiberacaoDigitacao.PeriodoInicial = 4 Then
            SeImprimeReducaoZ = True
        End If
    Else
        SeImprimeReducaoZ = True
    End If
End Function
Function TestaImprimeEncerrante() As Boolean
    Dim dados As String
    TestaImprimeEncerrante = False
    
    On Error GoTo FileError
    
    dados = ReadINI("CUPOM FISCAL", "Imprime Encerrante", gArquivoIni)
    If dados = "SIM" Then
        TestaImprimeEncerrante = True
    End If
    
    Exit Function
FileError:
    Exit Function
End Function
Private Sub TotalizaAcertoVendaECF()
    Dim i As Integer
    
    For i = 1 To 6
        If AcertoVendaECF.LocalizarCodigo(g_empresa, CDate(msk_data_i.Text), lTipoCombustivel(i)) Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: AcertoVendaECF:" & msk_data_i.Text & " Comb:" & lTipoCombustivel(i) & " " & AcertoVendaECF.Operacao & " Qtd:" & AcertoVendaECF.Quantidade)
            If AcertoVendaECF.Operacao = "+" Then
                lQtdBombaV(i) = lQtdBombaV(i) + AcertoVendaECF.Quantidade
                lTotalBombaV(i) = lTotalBombaV(i) + AcertoVendaECF.ValorTotal
            ElseIf AcertoVendaECF.Operacao = "-" Then
                lQtdBombaV(i) = lQtdBombaV(i) - AcertoVendaECF.Quantidade
                lTotalBombaV(i) = lTotalBombaV(i) - AcertoVendaECF.ValorTotal
            End If
        End If
    Next
End Sub
Private Sub TotalizaMovimentoAfericao()
    Dim i As Integer
    
    For i = 1 To 6
        lQtdAfericaoV(i) = lQtdAfericaoV(i) + MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 1, 9, lTipoCombustivel(i), "")
        lTotalAfericaoV(i) = lTotalAfericaoV(i) + MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 1, 9, lTipoCombustivel(i), "")
    Next
    
    'subtrai as aferições nas vendas de bomba
    For i = 1 To 6
        If lQtdAfericaoV(i) > 0 Then
            lQtdBombaV(i) = lQtdBombaV(i) - lQtdAfericaoV(i)
            lTotalBombaV(i) = lTotalBombaV(i) - lTotalAfericaoV(i)
            Call CriaLogCupom(Time & " - [Emissão do Cupom Complementar - AFERICAO ] À Vista - Combustivel:" & lTipoCombustivel(i) & " - Lts.Aferição:" & lQtdAfericaoV(i))
        End If
    Next
End Sub
Private Sub TotalizaMovimentoAbastecimento()
    Dim i As Integer
    With rstMovimentoAbastecimento
        .MoveFirst
        Do Until .EOF
            For i = 1 To 6
                If lTipoCombustivel(i) = ![Tipo de Combustivel] Then
                    lQtdBombaV(i) = lQtdBombaV(i) + !Quantidade
                    lTotalBombaV(i) = lTotalBombaV(i) + !ValorTotal
                    Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
End Sub
Private Sub TotalizaMovimentoBombaCupom()
    Dim i As Integer
    With rstMovimentoBomba
        .MoveFirst
        Do Until .EOF
            For i = 1 To 6
                If lTipoCombustivel(i) = ![Tipo de Combustivel] Then
                    lQtdBombaV(i) = lQtdBombaV(i) + ![Quantidade da Saida]
                    lTotalBombaV(i) = lTotalBombaV(i) + (![Quantidade da Saida] * ![Preco de Venda])
                    Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
End Sub
Private Sub TotalizaMovimentoCupomFiscal2(ByVal pTipoCombustivel As String)
    Dim i As Integer
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Total]) AS TotalCupom, SUM(Quantidade) AS QtdCupom"
    lSQL = lSQL & "  FROM Movimento_Cupom_Fiscal"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data do Cupom] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND [Data do Cupom] <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    Set rst = Conectar.RsConexao(lSQL)
    If Not rst.EOF Then
        For i = 1 To 6
            If lTipoCombustivel(i) = pTipoCombustivel Then
                If Not IsNull(rst!QtdCupom) Then
                    lQtdCupomV(i) = rst!QtdCupom
                    lTotalCupomV(i) = rst!TotalCupom
                End If
                Exit For
            End If
        Next
    End If
    rst.Close
End Sub
Function ValidaCampos() As Boolean
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
        MsgBox "Data final deve ser maior ou igual a " & msk_data_i.Text & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Not ValidaECF Then
        cmd_sair.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Function ValidaECF() As Boolean
    Dim i As Integer
    
    ValidaECF = False
    If lImpBematech Then
        BemaRetorno = Bematech_FI_FlagsFiscais(i)
        If i <> 0 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro na Rotina: ValidaECF - Bematech_FI_FlagsFiscais(i):" & i)
        End If
        If i = 33 Then
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro na Rotina: ValidaECF - Cupom Fiscal Aberto - Bematech_FI_FlagsFiscais(i):" & i)
            MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Else
            ValidaECF = True
        End If
    Else
        ValidaECF = True
    End If
End Function
Private Sub VerificaSeEcfTruncamento()
    Dim x_string As String
    
    lEcfTruncamento = False
    If lImpBematech Then
        x_string = Space(1)
        'Call CriaLogCupom("Bematech_FI_VerificaTruncamento(x_string) - x_string" & x_string)
        BemaRetorno = Bematech_FI_VerificaTruncamento(x_string)
        'Call CriaLogCupom("Bematech_FI_VerificaTruncamento - x_string" & x_string & " - BemaRetorno=" & BemaRetorno)
        If x_string = "1" Then
            lEcfTruncamento = True
        End If
    ElseIf lImpDaruma Then
        x_string = Space(2)
        Call CriaLogCupom("Daruma_FI_VerificaTruncamento(x_string) - x_string" & x_string)
        BemaRetorno = Daruma_FI_VerificaTruncamento(x_string)
        Call CriaLogCupom("Daruma_FI_VerificaTruncamento - x_string" & x_string & " - BemaRetorno=" & BemaRetorno)
        If Mid(x_string, 1, 1) = "1" Then
            lEcfTruncamento = True
        End If
    End If

    'Busca no ini quantidade de casas decimais da ECF para "Quantidade"
    lEcfQtdCasasDecimais = 3
    x_string = ReadINI("CUPOM FISCAL", "Quantidade Casa Decimal", gArquivoIni)
    If Val(x_string) > 0 Then
        lEcfQtdCasasDecimais = Val(x_string)
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Pedido a Visualização do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "Data Inicial:" & msk_data_i.Text & " Data Final:" & msk_data_f.Text)
            Call AtivaBotoes(False)
            Relatorio
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Visualizado")
            If TestaImprimeEncerrante Then
                If (MsgBox("Deseja imprimir encerrante de bombas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Relatório na ECF") = vbYes) Then
                    ImprimeEncerrante ("V")
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdImprimiEncerranteAtual_Click()
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Foi Pedido a Visualização do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    lLocal = 0
    If ValidaCampos Then
        ImprimeEncerranteAutomacaoAtual ("V")
    End If
End Sub

'Private Sub Command1_Click()
'
'    Dim NFCe As New NFCeDLL.ProcessaNFCEFronteira
'
'
'
'    If NFCe.ProcessaSolicitacaoFuncaoNFCe(3081, 8, 1) Then
'        MsgBox "Sucesso"
'    Else
'        MsgBox "Falha"
'    End If
'
'End Sub

Private Sub Form_Activate()
Dim rstCupomCompProd As New adodb.Recordset
Dim i As Integer

    Call GravaAuditoria(1, Me.name, 1, "")
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar (5): A Emissão do Cupom Complementar Foi Aberta")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(Date, "dd/mm/yyyy")
        msk_data_f.Text = Format(Date, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
    Screen.MousePointer = 1

'    lSQL = ""
'    lSQL = lSQL & "SELECT * "
'    lSQL = lSQL & "  FROM ConfiguracaoDiversa"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Nome like '%CUPOM COMPLEMENTAR PROD%'"
'
'    Set rstCupomCompProd = Conectar.RsConexao(lSQL)
'
'    If Not rstCupomCompProd.EOF Then
'        rstCupomCompProd.MoveFirst
'
'        For i = 1 To rstCupomCompProd.RecordCount
'            MsgBox rstCupomCompProd("Nome").Value, vbInformation, "Informação!"
'            rstCupomCompProd.MoveNext
'        Next
'    End If

    ''Set rstCupomCompProd = Conectar.RsConexao(lSQL)
    
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
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar(1): A Emissão do Cupom Complementar Foi Aberta")
    CentraForm Me
    
    MovimentoAfericao.NomeTabela = "Movimento_Afericao"
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar(4): A Emissão do Cupom Complementar Foi Aberta")
    AtualizaConstantes
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar(5): A Emissão do Cupom Complementar Foi Aberta")

    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    lCodigoFuncionario = 0
    If lCaixaIndividual And chbPorFuncionario.Value = 1 Then
        If g_string <> "" Then
            lCodigoFuncionario = Val(RetiraGString(1))
        End If
    End If
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
        cmd_visualizar.SetFocus
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
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
