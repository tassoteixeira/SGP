VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmissaoNFCeComplementarAuto 
   Caption         =   "Emissão de NFCe Complementar Automação"
   ClientHeight    =   3270
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoNFCeComplementarAuto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoNFCeComplementarAuto.frx":030A
   ScaleHeight     =   3270
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdImprimiEncerranteAtual 
      Caption         =   "Imprimir Encerrante Atual"
      Height          =   615
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoNFCeComplementarAuto.frx":0350
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
      Picture         =   "EmissaoNFCeComplementarAuto.frx":1A6A
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
      Picture         =   "EmissaoNFCeComplementarAuto.frx":3074
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
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "EmissaoNFCeComplementarAuto.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoNFCeComplementarAuto.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoNFCeComplementarAuto.frx":6CBA
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
   Begin VB.Frame FrameAguarde 
      Height          =   3255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "lblTitulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         Caption         =   "lblMensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   6675
      End
   End
End
Attribute VB_Name = "EmissaoNFCeComplementarAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lParar As Integer

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
Dim lNomeRelatorio As String
'Fim de variáveis padrão para relatório

Dim BemaRetorno As Integer
Dim lTipoCombustivel(1 To 7) As String
Dim lQtdBombaV(1 To 7) As Currency
Dim lTotalBombaV(1 To 7) As Currency
Dim lQtdAfericaoV(1 To 7) As Currency
Dim lTotalAfericaoV(1 To 7) As Currency
Dim lQtdCupomV(1 To 7) As Currency
Dim lTotalCupomV(1 To 7) As Currency

Dim lExisteProgramacaoADiminuir As Boolean
Dim lExisteProgramacaoASomar As Boolean

Dim lQtdBico As Integer
Dim lPeriodo As Integer
Dim lNumeroNFCe As Long
Dim lOrdemNFCe As Integer
Dim lDataNFCe As Date
Dim lHoraNFCe As Date
Dim lSerieNFCe As String
Dim lNumeroPDV As Integer
Dim lSerieECF As String
Dim lAutomacao As Boolean
Dim lImprimeNFCeComplementarAuto As Boolean
Dim lCaixaIndividual As Boolean
Dim lCodigoFuncionario As Integer
Dim lEcfTruncamento As Boolean
Dim lEcfQtdCasasDecimais As Integer
Dim lNFCeComplementarSemRetorno As Boolean
Dim lDataHoraInicioProcessamento As Date
Dim lTerminoDefinido As Boolean
Dim lPetromovelAutorizaNFCe As Boolean


'--- CRIADO PARA NFCE ---
Const MODELO_NFCE As String = "65"
Const PROGRAMA_ORIGEM As String = "COMPLEMENTAR_AUTO"


Private AcertoVendaECF As New cAcertoVendaECF
Private Aliquota As New cAliquota
Private BaixaAbastecimento As New cBaixaAbastecimento
Private Bomba As New cBomba
Private CidadeIBGE As New cCidadeIBGE
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
'Private ECF As New cEcf
Private Estoque As New cEstoque
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private MovimentoAbastecimento As New cMovimentoAbastecimento
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
Private MovimentoCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovDocEletronicoCabecalho As New cMovDocEletronicoCabecalho
Private MovDocEletronicoItem As New cMovDocEletronicoItem
Private MovNotaFiscalSaidaItem As New cMovNotaFiscalSaidaItem
Private PercentualImposto As New cPercentualImposto
Private PrevisaoVendaPrazo As New cPrevisaoVendaPrazo
Private Produto As New cProduto
Private ReducaoZ As New cReducaoZ
Private MovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
Private lProcessamentoNFCeComplementar As New cProcessaNFCeComplementar

Private Sub AlteraAbastecimentoParaConcluido(ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer, ByVal pNumeroNFCe As Long, ByVal pBaixado As Boolean)

    On Error GoTo trata_erro
            
    Call CriaLogAutomacao("AlteraAbastecimentoParaConcluido: Definindo abastecimento para Acerto=true. pData: " & pData & " - pHora: " & pHora & " - pBico: " & pBico)
    If pBaixado = True Then
        If BaixaAbastecimento.LocalizarCodigo(g_empresa, pBico, pData, pHora) Then
            BaixaAbastecimento.Acerto = True
            BaixaAbastecimento.NumeroCupom = pNumeroNFCe
            BaixaAbastecimento.CodigoECF = lNumeroPDV
            BaixaAbastecimento.DocumentoGerado = "NFCe"
            If Not BaixaAbastecimento.Alterar(g_empresa, pBico, pData, pHora) Then
                MsgBox "Não foi possível alterar a baixa de abastecimento!", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível localizar a baixa de abastecimento!", vbInformation, "Erro de Integridade!"
        End If
    Else
        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, pData, pHora, pBico) Then
            MovimentoAbastecimento.Acerto = True
            MovimentoAbastecimento.NumeroCupom = pNumeroNFCe
            MovimentoAbastecimento.CodigoECF = lNumeroPDV
            MovimentoAbastecimento.DocumentoGerado = "NFCe"
            'AFTEMP - Abastecimento selecionado para vincular a aferição
            'AFERICAO - Abastecimento vinculado à Afericao
            'AVP  - Acerto de Venda Programada
            'CF   - Cupom Fiscal
            'NFCe - Nota Fiscal do Consumidor Eletrônica
            'NT   - Nota Abastecimento
            'CP   - Cupom Complementar
            'AF   - Afericao
            'CHVIS- Cheque A Vista
            'CHPRE- Cheque Pre-Datado
            'CRT  - Cartao de Credito
            'DIN  - Dinheiro
            'DESPC- Despesa de Caixa
            'VALEF- Vale de Funcionario
            'VLABR- Vale Abastecimento Recebido
            'CRAR - Credito Antecipado Recebido
            If Not MovimentoAbastecimento.Alterar(g_empresa, pData, pHora, pBico) Then
                Call CriaLogAutomacao("AlteraAbastecimentoParaConcluido: Erro ao definir abastecimento para Acerto=true.")
                MsgBox "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
            End If
        Else
            Call CriaLogAutomacao("AlteraAbastecimentoParaConcluido: Erro ao localizar abastecimento para definir Acerto=true.")
            MsgBox "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
        End If
    End If
    Exit Sub

trata_erro:
    Call CriaLogAutomacao("Erro AlteraAbastecimentoParaConcluido: Erro=" & Err.Number & " - " & Err.Description)
End Sub
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
    
    dados = ReadINI("NFCe", "Emite NFCe", gArquivoIni)
    Me.Caption = Me.Caption & " - Emite NFCe: " & dados
    If dados = "SIM" Then
        cmd_imprimir.Enabled = True
    Else
        cmd_imprimir.Enabled = False
    End If
    
    lNumeroPDV = 901
    dados = ReadINI("NFCe", "Numero do PDV", gArquivoIni)
    If Len(dados) > 0 Then
        lNumeroPDV = Val(dados)
    End If
    
    lImprimeNFCeComplementarAuto = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Imprimir Cupom Complementar") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lImprimeNFCeComplementarAuto = True
        End If
    End If

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

    lNFCeComplementarSemRetorno = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCe Complementar Sem Retorno") Then
        lNFCeComplementarSemRetorno = ConfiguracaoDiversa.Verdadeiro
    End If

    lPetromovelAutorizaNFCe = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
        lPetromovelAutorizaNFCe = ConfiguracaoDiversa.Verdadeiro
    End If

    lEcfTruncamento = False
    lEcfQtdCasasDecimais = 3
End Sub
Private Sub AtualizaTabelaCupomFiscal(ByVal pNumeroCupom As Long, ByVal pOrdem As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pCodigoProduto As Long, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pCodigoAliquota As Integer, ByVal pTipoCombustivel As String, ByVal pCodigoGrupo, ByVal pCodigoBico As Integer, ByVal pEncerrante As Currency)
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
    MovimentoCupomFiscal.NumeroCheque = "010101"
    If Bomba.LocalizarCodigo(g_empresa, pCodigoBico) Then
        MovimentoCupomFiscal.NumeroCheque = Format(Bomba.CodigoFisicoBomba, "00") & Format(Bomba.Codigo, "00") & Format(Bomba.NumeroTanque, "00")
    End If
    MovimentoCupomFiscal.Telefone = pEncerrante
    MovimentoCupomFiscal.operador = 1
    MovimentoCupomFiscal.CupomCancelado = False
    MovimentoCupomFiscal.ItemCancelado = False
    MovimentoCupomFiscal.CodigoAliquota = pCodigoAliquota
    MovimentoCupomFiscal.ValorDesconto = 0
    MovimentoCupomFiscal.Nome = ""
    MovimentoCupomFiscal.CPFCNPJ = ""
    MovimentoCupomFiscal.ValorDescontoEmbutido = 0
    MovimentoCupomFiscal.TipoCombustivel = pTipoCombustivel
    MovimentoCupomFiscal.CodigoECF = lNumeroPDV
    MovimentoCupomFiscal.CodigoGrupo = pCodigoGrupo
    MovimentoCupomFiscal.TipoSubEstoque = 2 'Pista
    
    If Not MovimentoCupomFiscal.Incluir Then
        Call CriaLogCupom(Time & " - ERRO Emissão de NFCe Complementar Automação: Erro na Rotina: AtualizaTabelaCupomFiscal")
        MsgBox "Não foi possível incluir registro", vbInformation, "Erro de Integridade!"
    End If
    
    Exit Sub
    
FileError:
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Erro na Rotina: AtualizaTabelaCupomFiscal")
    MsgBox "Erro Gravando Cupom: " & Error
    Exit Sub
End Sub
Private Sub AtualizaTabelaCupomFiscalItem()
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
    MovimentoCupomFiscalItem.CodigoAliquota = MovimentoCupomFiscal.CodigoAliquota
    MovimentoCupomFiscalItem.CodigoGrupo = MovimentoCupomFiscal.CodigoGrupo
    If Not MovimentoCupomFiscalItem.Incluir Then
        Call CriaLogCupom(Time & " - ERRO Emissão de NFCe Complementar Automação: Erro na Rotina: AtualizaTabelaCupomFiscalItem - ECF:" & MovimentoCupomFiscalItem.NumeroCupom)
        MsgBox "Não foi possível incluir registro de item", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub
    
FileError:
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Erro na Rotina: AtualizaTabelaCupomFiscalItem - ECF:" & MovimentoCupomFiscalItem.NumeroCupom)
    MsgBox "Erro Gravando Cupom: " & Error
    Exit Sub
End Sub
Private Sub AtualizaTabelaDocumentoEletronicoCabecalho(ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pHora As Date, ByVal pValorTotal As Currency)
    
    On Error GoTo trata_erro
    
    If lOrdemNFCe = 1 Then
        MovDocEletronicoCabecalho.IdEstabelecimento = g_empresa
        MovDocEletronicoCabecalho.DataEmissao = pData
        MovDocEletronicoCabecalho.Entrada = False
        MovDocEletronicoCabecalho.Saida = True
        MovDocEletronicoCabecalho.Modelo = MODELO_NFCE
        MovDocEletronicoCabecalho.Serie = lSerieNFCe
        MovDocEletronicoCabecalho.numero = pNumeroCupom
        MovDocEletronicoCabecalho.HoraSaida = pHora
        MovDocEletronicoCabecalho.DataEntradaSaida = pData
        MovDocEletronicoCabecalho.EmissaoPropria = True
        MovDocEletronicoCabecalho.IdClienteFornecedor = 0
        MovDocEletronicoCabecalho.CodigoSituacao = 0
        MovDocEletronicoCabecalho.ChaveAcesso = ""
        MovDocEletronicoCabecalho.FormaPagamento = "1"
        MovDocEletronicoCabecalho.ValorTotal = pValorTotal
        MovDocEletronicoCabecalho.ValorDesconto = 0
        MovDocEletronicoCabecalho.ValorAbatimentoNaoTributado = 0
        MovDocEletronicoCabecalho.ValorProdutos = pValorTotal
        MovDocEletronicoCabecalho.TipoFrete = "1"
        MovDocEletronicoCabecalho.ValorFrete = 0
        MovDocEletronicoCabecalho.ValorSeguro = 0
        MovDocEletronicoCabecalho.OutrasDespesas = 0
        MovDocEletronicoCabecalho.ValorBCICMS = 0
        MovDocEletronicoCabecalho.AliquotaICMS = 0
        MovDocEletronicoCabecalho.ValorICMS = 0
        MovDocEletronicoCabecalho.ValorBCICMSST = 0
        MovDocEletronicoCabecalho.AliquotaIcmsSt = 0
        MovDocEletronicoCabecalho.ValorICMSST = 0
        MovDocEletronicoCabecalho.ValorIPI = 0
        MovDocEletronicoCabecalho.ValorPis = 0
        MovDocEletronicoCabecalho.ValorCofins = 0
        MovDocEletronicoCabecalho.ValorPisSt = 0
        MovDocEletronicoCabecalho.ValorCofinsSt = 0
        MovDocEletronicoCabecalho.Combustivel = True
        MovDocEletronicoCabecalho.Cancelado = False
        MovDocEletronicoCabecalho.AguaEnergiaGasTelefone = ""
        MovDocEletronicoCabecalho.Inutilizada = False
        MovDocEletronicoCabecalho.IncidePisConfis = False
        MovDocEletronicoCabecalho.DataDigitacao = pData
        MovDocEletronicoCabecalho.DataAlteracao = "00:00:00"
        MovDocEletronicoCabecalho.IdUsuario = 1
        MovDocEletronicoCabecalho.NumeroLote = 0
        MovDocEletronicoCabecalho.NumeroRecepcao = 0
        MovDocEletronicoCabecalho.NumeroProtocolo = 0
        MovDocEletronicoCabecalho.EtapaConcluida = 0
        MovDocEletronicoCabecalho.CodigoUltimoEvento = Val(EVENTO_NFCE.NENHUM_EVENTO)
        MovDocEletronicoCabecalho.ObservacaoEvento = ""
        MovDocEletronicoCabecalho.ProgramaOrigem = PROGRAMA_ORIGEM
        
        MovDocEletronicoCabecalho.Periodo = lPeriodo
        MovDocEletronicoCabecalho.TipoMovimento = 2 'pista
        MovDocEletronicoCabecalho.TipoSubEstoque = 2 'pista
        
        If Not MovDocEletronicoCabecalho.Incluir Then
            Call CriaLogCupom(Time & " - Não foi possível incluir registro NFCe Cabecalho")
            MsgBox "Não foi possível incluir registro NFCe Cabecalho", vbCritical, "Erro de Integridade!"
        Else
           Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.ABERTA_COMPLEMENTAR_AUTO)
        End If
    Else
        If MovDocEletronicoCabecalho.LocalizarCodigo(g_empresa, pData, False, True, MODELO_NFCE, lSerieNFCe, pNumeroCupom) = True Then
            MovDocEletronicoCabecalho.ValorProdutos = pValorTotal
            MovDocEletronicoCabecalho.ValorTotal = pValorTotal
            If Not MovDocEletronicoCabecalho.Alterar(g_empresa, pData, False, True, MODELO_NFCE, lSerieNFCe, pNumeroCupom) Then
                Call CriaLogCupom(Time & " - Não foi possível alterar registro NFCe Cabecalho")
                MsgBox "Não foi possível alterar registro NFCe Cabecalho", vbCritical, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível localizar o registro NFCe Cabecalho pra alterar", vbCritical, "Erro de Integridade!"
        End If
    End If
    Exit Sub
    
trata_erro:
    Call CriaLogCupom(Time & "Erro não Identificado ao gravar NFCe Cabecalho Erro:" & Err.Description)
End Sub
Private Sub AtualizaTabelaDocumentoEletronicoItem(ByVal pNumeroCupom As Long, ByVal pOrdem As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pCodigoProduto As Long, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pCodigoAliquota As Integer, ByVal pTipoCombustivel As String, ByVal pCodigoBico As Integer, ByVal pEncerrante As Currency)
          'Dim i As Integer
    On Error GoTo trata_erro
    
    MovDocEletronicoItem.IdEstabelecimento = g_empresa
    MovDocEletronicoItem.DataEmissao = pData
    MovDocEletronicoItem.Entrada = False
    MovDocEletronicoItem.Saida = True
    MovDocEletronicoItem.Modelo = MODELO_NFCE
    MovDocEletronicoItem.Serie = lSerieNFCe
    MovDocEletronicoItem.numero = pNumeroCupom
    MovDocEletronicoItem.Ordem = pOrdem
    MovDocEletronicoItem.IdClienteFornecedor = 0
    MovDocEletronicoItem.MovimentacaoFisica = True
    MovDocEletronicoItem.Cfop = "5656"
    MovDocEletronicoItem.IdProduto = pCodigoProduto
    MovDocEletronicoItem.ValorUnitario = pValorUnitario
    MovDocEletronicoItem.Quantidade = pQuantidade
    MovDocEletronicoItem.ValorDesconto = 0
    MovDocEletronicoItem.ValorTotalLiquido = pValorTotal
    MovDocEletronicoItem.CSTICMS = Produto.CSTICMS
    MovDocEletronicoItem.ValorBCICMS = 0
    MovDocEletronicoItem.AliquotaICMS = 0 'OBTER ALIQUOTA
    MovDocEletronicoItem.ValorICMS = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.CstIcmsSt = Produto.CSTICMS 'VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.ValorBCICMSST = 0 'VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.AliquotaIcmsSt = 0 'OBTER ALIQUOTA
    MovDocEletronicoItem.ValorICMSST = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.CSTIPI = Produto.CSTIPI
    MovDocEletronicoItem.ValorBcIpi = 0 ''VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.AliquotaIpi = 0 ''OBTER ALIQUOTA
    MovDocEletronicoItem.ValorIPI = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.ApuracaoIpiMensal = False 'VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.CodigoEnquadramentoIpi = "" 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.CstPis = Produto.CstPis
    MovDocEletronicoItem.ValorBcPis = 0 ''VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.QuantidadeBcPis = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.AliquotaPis = 0 'OBTER ALIQUOTA
    MovDocEletronicoItem.ValorPis = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.CstCofins = Produto.CstCofins
    MovDocEletronicoItem.ValorBcCofins = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.QuantidadeBcCofins = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.ValorCofins = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.NumeroTanque = 0
    MovDocEletronicoItem.Cancelado = False
    MovDocEletronicoItem.DataEntradaSaida = pData
    MovDocEletronicoItem.EncerranteFinal = Format(pEncerrante, "#######.00")
    MovDocEletronicoItem.NumeroBomba = 0
    MovDocEletronicoItem.NumeroBico = 0
    If Bomba.LocalizarCodigo(g_empresa, pCodigoBico) Then
       MovDocEletronicoItem.NumeroBomba = Format(Bomba.CodigoFisicoBomba, "00")
       MovDocEletronicoItem.NumeroBico = Format(Bomba.Codigo, "00")
       MovDocEletronicoItem.NumeroTanque = Format(Bomba.NumeroTanque, "00")
    End If
    MovDocEletronicoItem.TipoCombustivel = pTipoCombustivel
    MovDocEletronicoItem.EtapaConcluida = 0
    MovDocEletronicoItem.ProgramaOrigem = PROGRAMA_ORIGEM
    
    MovDocEletronicoItem.Periodo = lPeriodo
    
    If Not MovDocEletronicoItem.Incluir Then
        Call CriaLogCupom(Time & " - Não foi possível incluir registro NFCe Item")
        MsgBox "Não foi possível incluir registro NFCe Item", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub
    
trata_erro:
    Call CriaLogCupom(Time & "Erro não Identificado ao gravar NFCe Item Erro:" & Err.Description)
End Sub

Private Function AtualizaTabelaSolicitacaoNFCe(ByVal pTipoOperacao As String, ByVal pChaveAcessoNFe As String, ByVal pTexto As String, ByVal pNumeroDaNota As Long, ByVal pNumeroLote As String) As Boolean
    AtualizaTabelaSolicitacaoNFCe = False
    
    MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    MovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = gNumeroControleSolicitacao
    MovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy HH:mm:ss"))
    MovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = pTipoOperacao
    MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroDaNota
    MovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = pChaveAcessoNFe
    If pTipoOperacao = "STATUS SERVICO" Or pTipoOperacao = "ATV" Or pTipoOperacao = "IMPRESSAO" Then
        MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = ""
        MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = "0"
    End If
    MovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    MovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    MovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    MovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    MovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    MovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    MovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pTexto
    MovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    MovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
    AtualizaTabelaSolicitacaoNFCe = LoopIncluiRegistroSolicitacaoNFe()

End Function
Private Function AutomacaoAlteraArredondamento(ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pBaixado As Boolean) As Currency
    Dim xDiferenca As Currency
    Dim xValorTotal As Currency

    On Error GoTo trata_erro
    
    AutomacaoAlteraArredondamento = pQuantidade
    
    If pBaixado = True Then
        xValorTotal = Round(pQuantidade * pValorUnitario, 2)
        xDiferenca = pValorTotal - xValorTotal
        If xDiferenca > 0 Then
            If BaixaAbastecimento.LocalizarCodigo(g_empresa, pBico, pData, pHora) Then
                BaixaAbastecimento.Quantidade = Round(BaixaAbastecimento.ValorTotal / BaixaAbastecimento.ValorUnitario, 3)
                AutomacaoAlteraArredondamento = BaixaAbastecimento.Quantidade
                If Not BaixaAbastecimento.Alterar(g_empresa, pBico, pData, pHora) Then
                    MsgBox "Erro em: ComplementarAuto AutomacaoAlteraArredondamento BAIXADO" & vbCrLf & "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível localizar a baixa de abastecimento!", vbInformation, "Erro de Integridade!"
            End If
        End If
    Else
'        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, pData, pHora, pBico) Then
'            xValorTotal = Round(MovimentoAbastecimento.Quantidade * MovimentoAbastecimento.ValorUnitario, 2)
'            xDiferenca = MovimentoAbastecimento.ValorTotal - xValorTotal
'            If xDiferenca > 0 Then
'                MovimentoAbastecimento.Quantidade = Round(MovimentoAbastecimento.ValorTotal / MovimentoAbastecimento.ValorUnitario, 3)
'                AutomacaoAlteraArredondamento = MovimentoAbastecimento.Quantidade
'                If Not MovimentoAbastecimento.Alterar(g_empresa, pData, pHora, pBico) Then
'                    MsgBox "Erro em: ComplementarAuto AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
'                End If
'            Else
'                AutomacaoAlteraArredondamento = MovimentoAbastecimento.Quantidade
'            End If
'        Else
'            MsgBox "Erro em: ComplementarAuto AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
'        End If
'        xValorTotal = Round(MovimentoAbastecimento.Quantidade * MovimentoAbastecimento.ValorUnitario, 2)
'        xDiferenca = MovimentoAbastecimento.ValorTotal - xValorTotal
        xValorTotal = Round(pQuantidade * pValorUnitario, 2)
        xDiferenca = pValorTotal - xValorTotal
        If xDiferenca > 0 Then
            If MovimentoAbastecimento.LocalizarCodigo(g_empresa, pData, pHora, pBico) Then
                MovimentoAbastecimento.Quantidade = Round(MovimentoAbastecimento.ValorTotal / MovimentoAbastecimento.ValorUnitario, 3)
                AutomacaoAlteraArredondamento = MovimentoAbastecimento.Quantidade
                If Not MovimentoAbastecimento.Alterar(g_empresa, pData, pHora, pBico) Then
                    MsgBox "Erro em: ComplementarAuto AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Erro em: ComplementarAuto AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogAutomacao("Erro ComplementarAuto AutomacaoAlteraArredondamento: Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Sub BuscaNumeroCupom()
    Dim xString As String
    
    On Error GoTo FileError
    
    'busca numero de NFCe da impressora fiscal
    xString = ConfiguracaoDiversa.BuscaProximoCodigo(g_empresa, "NFCe: Numero", True)
    If Len(xString) > 0 Then
        lNumeroNFCe = CLng(RetiraString(1, xString))
        lSerieNFCe = RetiraString(2, xString)
    Else
        lNumeroNFCe = 1
        lSerieNFCe = "1"
    End If
    lOrdemNFCe = 1
    
    'busca data/hora da impressora fiscal
    lDataNFCe = Date
    lHoraNFCe = Time
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
Private Function CalculaImpostos(ByVal pNumeroNFCe As Long, ByVal pData As Date) As String
    Dim xBaseCalculo As Currency
    Dim xTotalCupom As Currency
    Dim xTotalImpostos As Currency
    Dim xPercentualImpostos As Currency
    Dim xDescontoCupom As Currency
    Dim xOrdem As Integer
    Dim xString As String
    
    'Call CriaLogCupom("CalculaImpostos: Fase 1=")
    CalculaImpostos = ""
    xBaseCalculo = 0
    xTotalCupom = 0
    xTotalImpostos = 0
    xPercentualImpostos = 0
    xDescontoCupom = 0
    xOrdem = 0
    
    Do Until MovDocEletronicoItem.LocalizarProximoDestaOrdem(g_empresa, pData, False, True, MODELO_NFCE, lSerieNFCe, pNumeroNFCe, xOrdem) = False
        If MovDocEletronicoItem.Cancelado = False Then
            If Produto.LocalizarCodigo(MovDocEletronicoItem.IdProduto) Then
                If LocalizarNCM(0, Produto.CodigoNCM) Then
                    xBaseCalculo = MovDocEletronicoItem.ValorTotalLiquido
                    xTotalCupom = xTotalCupom + xBaseCalculo
                    xTotalImpostos = xTotalImpostos + (Round(xBaseCalculo * PercentualImposto.AliquotaNacional / 100, 2))
                Else
                    Call CriaLogCupom("CalculaImpostos: - NCM nao localizado. Produto.CodigoNCM=" & Produto.CodigoNCM)
                End If
            Else
                'Call CriaLogCupom("CalculaImpostos: Fase 2 z - Produto nao localizado. Codigo=" & MovCupomFiscal.CodigoProduto)
            End If
        End If
        xOrdem = MovDocEletronicoItem.Ordem
    Loop
    If xTotalCupom > 0 And xTotalImpostos > 0 Then
        xPercentualImpostos = Round(xTotalImpostos / xTotalCupom * 100, 2)
        xString = "Val.Aprox.Tributos R$ " & Format(xTotalImpostos, "###,##0.00") & "(" & Format(xPercentualImpostos, "##0.00") & "%) Fonte: IBPT"
        If Len(xString) < 48 Then
            Do Until Len(xString) = 48
                xString = xString & " "
            Loop
        End If
        CalculaImpostos = Mid(xString, 1, 48)
    End If
    Call CriaLogCupom("CalculaImpostos: CalculaImpostos=" & CalculaImpostos)
End Function
Private Sub EnviaDadosParaNFCe(ByVal pNumeroCupom As Long, ByVal pDataCupom As Date, ByVal pValorTotalNFCe As Currency)
    Dim rsDadosParaNFCe As New adodb.Recordset
    Dim xTipoServico As String
    Dim xTextoSolicitacao As String
    
On Error GoTo trata_erro


 '   MsgBox "EnviaDadosParaNFCe - ALEX TESTE 2"
    xTipoServico = "NFCe 3.10"
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "VERSAO NFCE") Then
        xTipoServico = "NFCe" & " " & ConfiguracaoDiversa.Texto
    End If

    
    Set rsDadosParaNFCe = ObtenhaDadosParaNFCEDocumentoEletronico(lNumeroNFCe, pDataCupom)
  '  MsgBox "EnviaDadosParaNFCe - ALEX TESTE 3"
    
    If rsDadosParaNFCe.RecordCount > 0 Then
        rsDadosParaNFCe.MoveFirst
        
   '     MsgBox "EnviaDadosParaNFCe - ALEX TESTE 4"
        xTextoSolicitacao = MontaTextoCabecalhoSolicitacaoNFCE(rsDadosParaNFCe, pValorTotalNFCe, xTipoServico)
        
        xTextoSolicitacao = xTextoSolicitacao & MontaTextoItensSolicitacaoNFCE(rsDadosParaNFCe)
        
    '    MsgBox "EnviaDadosParaNFCe - ALEX TESTE 5"
        If (AtualizaTabelaSolicitacaoNFCe(xTipoServico, "", xTextoSolicitacao, lNumeroNFCe, "")) Then
        
           
           Dim Mensagem As String
     '       MsgBox "EnviaDadosParaNFCe - ALEX TESTE 6"
'           Set lProcessadorNFCE = New ProcessaNFCEFronteira
            'MsgBox "Vai chamar o Processamento - ALEX TESTE 7 - NSU = " & MovSolicitacaoFuncaoNFe.NSU
'           Mensagem = lProcessadorNFCE.ProcessaSolicitacaoFuncaoNFCe(MovSolicitacaoFuncaoNFe.NSU, MovSolicitacaoFuncaoNFe.CodigoEstabelecimento, GERADOR_NFCE_OOBJ)
            'MsgBox "EnviaDadosParaNFCe - ALEX TESTE 8 - Processamento Finalizado"
            
            
                            
            Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação.EnviaDadosParaNFCe: lNFCeComplementarSemRetorno=" & lNFCeComplementarSemRetorno & " - pNumeroCupom=" & pNumeroCupom)
            
            Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.FECHADA)
'            If lNFCeComplementarSemRetorno = True Then
'                Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação.EnviaDadosParaNFCe: gStringChamada=" & gStringChamada & " - pNumeroCupom=" & pNumeroCupom)
'                gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "false" & "|@|" & lNumeroNFCe & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
'                'gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "false" & "|@|" '1-Oobj TXT, 2-Oobj XML, 3-cerrado
'                Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
'                AguardaMS (500)
'            Else
'                gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "True" & "|@|" & lNumeroNFCe & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
'                'gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "true" & "|@|" '1-Oobj TXT, 2-Oobj XML, 3-cerrado
'                Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
'            End If
            
            
            If lNFCeComplementarSemRetorno = True Then
                gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "false" & "|@|" & lNumeroNFCe & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
            Else
                gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & "false" & "|@|" & gCNPJEmpresa & "|@|" & "True" & "|@|" & lNumeroNFCe & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
            End If
            
            Call CriaLogCupom("[EnviaDadosParaNFCe] - NFCE Complementar Auto: gStringChamada=" & gStringChamada)
            
            If lPetromovelAutorizaNFCe = True Then
                Call CriaLogCupom("[EnviaDadosParaNFCe] - NFCE Complementar Auto: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=True")
                If Not GravaSolicitacaoProcessamentoNFCe(MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe, gStringChamada) Then
                    Call CriaLogCupom("[EnviaDadosParaNFCe] - NFCE Complementar Auto: Não foi possível gravar a Solicitação do processamento para NFC-e!")
                    MsgBox "Não foi possível gravar a Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
                End If
                gStringChamada = ""
            Else
                Call CriaLogCupom("[EnviaDadosParaNFCe] - NFCE Complementar Auto: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=False")
                Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
            End If

                            
        
        End If
        
    End If
    Set rsDadosParaNFCe = Nothing
    Exit Sub

trata_erro:
    Dim ErroNFCE As String
    ErroNFCE = Err.Description
    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: Erro=" & Err.Number & " - " & ErroNFCE)
    MsgBox "Não foi possível gerar NFC-e. " & vbCrLf & ErroNFCE, vbCritical, "Erro Grave!"
    Exit Sub

End Sub


Private Function GravaSolicitacaoProcessamentoNFCe(ByVal pNumeroNFCe As String, ByVal pStringChamadaProcessamento As String) As Boolean
    GravaSolicitacaoProcessamentoNFCe = False
    
    On Error GoTo TrataError
    
    Dim xMovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
    
    xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    xMovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    xMovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = "PROCESSA_NFCE"
    xMovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    xMovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    xMovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroNFCe
    xMovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    xMovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    xMovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    xMovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    xMovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    xMovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    xMovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pStringChamadaProcessamento
    xMovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
    
    GravaSolicitacaoProcessamentoNFCe = xMovSolicitacaoFuncaoNFe.Incluir
    
    Exit Function
TrataError:
    Call CriaLogSGP("[GravaSolicitacaoProcessamentoNFCe]", "Erro ao tentar gravar solicitação do processamento da NFCe Complementar - " & Err.Description, "pStringChamadaProcessamento=" & pStringChamadaProcessamento)
    MsgBox "Não foi possível incluir registro de Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
End Function


Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rstMovimentoBomba = Nothing
    Set rst = Nothing
    Set rst2 = Nothing
    
    Set AcertoVendaECF = Nothing
    Set Aliquota = Nothing
    Set BaixaAbastecimento = Nothing
    Set Bomba = Nothing
    Set CidadeIBGE = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set Estoque = Nothing
    Set LiberacaoDigitacao = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set MovimentoCupomFiscalItem = Nothing
    Set MovimentoAbastecimento = Nothing
    Set MovimentoAfericao = Nothing
    Set MovNotaFiscalSaidaItem = Nothing
    Set PercentualImposto = Nothing
    Set PrevisaoVendaPrazo = Nothing
    Set Produto = Nothing
    Set ReducaoZ = Nothing
    Set MovSolicitacaoFuncaoNFe = Nothing
    Set lProcessamentoNFCeComplementar = Nothing
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
    lTipoCombustivel(7) = "GE"
    lExisteProgramacaoADiminuir = True
    lExisteProgramacaoASomar = True
    For i = 1 To 7
        lQtdBombaV(i) = 0
        lTotalBombaV(i) = 0
        lQtdAfericaoV(i) = 0
        lTotalAfericaoV(i) = 0
        lQtdCupomV(i) = 0
        lTotalCupomV(i) = 0
    Next
End Sub

Private Function LocalizarNCM(ByVal pTabela As Integer, ByVal pCodigo As String) As Boolean
    LocalizarNCM = False
    If Trim(pCodigo) = "" Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "Está cadastrado sem NCM.", vbInformation, "Produto Sem NCM!"
        Exit Function
    End If
    If Not PercentualImposto.LocalizarCodigo(pTabela, pCodigo) Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "De NCM:" & Trim(pCodigo) & vbCrLf & "Está sem NCM cadastrado.", vbInformation, "NCM Inexistente!"
        Exit Function
    End If
    LocalizarNCM = True
End Function
Private Function LoopIncluiRegistroSolicitacaoNFe() As Boolean
    Dim i As Integer
    LoopIncluiRegistroSolicitacaoNFe = False
    For i = 1 To 30
        If MovSolicitacaoFuncaoNFe.Incluir Then
            LoopIncluiRegistroSolicitacaoNFe = True
            Exit For
        End If
    Next
    If LoopIncluiRegistroSolicitacaoNFe = False Then
        MsgBox "Não foi possível incluir registro de Solicitação de Função NFe!", vbCritical, "Erro de Integridade"
    End If
End Function
Private Function MontaTextoCabecalhoSolicitacaoNFCE(ByVal pRsDadosParaNFCe As adodb.Recordset, ByVal pValorTotalNFCe As Currency, ByVal pTipoServico As String) As String

    MontaTextoCabecalhoSolicitacaoNFCE = Empty
    
    Dim xEmpresa As New cEmpresa
    Dim xStringNfce As String
    Dim xTipoServico As String
    Dim xCNPJEmpresa As String
    Dim xMensagem As String
    Dim xCodigoCidadeIbgeEmitente As String
    Dim xCodigoCidadeIbgeDestinatario As String
    
    xTipoServico = pTipoServico
    xCNPJEmpresa = Empty
    
    If xEmpresa.LocalizarCodigo(g_empresa) Then xCNPJEmpresa = xEmpresa.CGC
    If CidadeIBGE.LocalizarNome(UCase(xEmpresa.Estado), UCase(xEmpresa.Cidade)) Then
        xCodigoCidadeIbgeEmitente = CidadeIBGE.Codigo
        xCodigoCidadeIbgeDestinatario = CidadeIBGE.Codigo
    End If
    
    xStringNfce = xStringNfce & "000-000 = " & xTipoServico & "|@|" & vbCrLf
    
    'NÃO NECESSÁRIO PARA NFCE MANTIDO APENAS PARA CONSERVAR PADRÃO
    xStringNfce = xStringNfce & "001-000 = " & Format(0, "0000000000") & "|@|" & vbCrLf 'OBTER DADOS REAIS - MÉTODO PARA GERA CONTROLE DE SOLICITAÇÃO
    xStringNfce = xStringNfce & "002-000 = " & gVersaoSGP & "|@|" & vbCrLf
    xStringNfce = xStringNfce & "040-000 = " & g_empresa & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "041-000 = " & "TEIXEIRA E PINHEIRO LTDA" & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "042-000 = " & "3" & "|@|" & vbCrLf 'CRT 1-SIMPLES NACIONAL, 3-REGIME NORMAL
    xStringNfce = xStringNfce & "045-000 = " & "1" & "|@|" & vbCrLf 'Local de Destino da operação: 1-Interna, 2-Interestadual, 3-Exterior
    xStringNfce = xStringNfce & "046-000 = " & "1" & "|@|" & vbCrLf 'FINALIDADE DA NFE

    'NUMERAÇÃO ESTÁ FORA DA SEQUENCIA POIS INICIALMENTE FOI UTILIZADA A ESTRUTURA EXISTENTE DA SOLICITAÇÃO DE NFE
    'ALGUMAS NUMERAÇÕES FORAM REMOVIDAS POR NÃO TEREM UTILIDADE NA NFCE
    'QUALQUER ALTERAÇÃO NESTAS NUMERAÇÕES IMPACTAM O FUNCIONAMENTO DA EMISSÃO NFCE PELA DLL
    xStringNfce = xStringNfce & "100-000 = INICIO" & "|@|" & vbCrLf 'VALOR FIXO
    
    xStringNfce = xStringNfce & "111-000 = " & lNumeroNFCe & "|@|" & vbCrLf

    xStringNfce = xStringNfce & "112-000 = " & lSerieNFCe & "|@|" & vbCrLf

'---- DADOS DO PAGAMENTO -----
    xStringNfce = xStringNfce & "113-001 = " & "01" & "|@|" & vbCrLf 'Forma de Pagamento: 01-Diheiro
    
    xStringNfce = xStringNfce & "114-001 = " & "|@|" & vbCrLf 'Tipo de integração do Pagamento 1=Pagamento integrado com o sistema 2=equipamento POS
    
    xStringNfce = xStringNfce & "115-001 = " & "|@|" & vbCrLf 'CNPJ OPERADORA DO CARTÃO

    xStringNfce = xStringNfce & "116-001 = " & FormatNumber(pValorTotalNFCe, 2) & "|@|" & vbCrLf

    xStringNfce = xStringNfce & "117-001 = " & "|@|" & vbCrLf 'BANDEIRA DO CARTÃO
    
    xStringNfce = xStringNfce & "118-001 = " & "|@|" & vbCrLf 'NUMERO AUTORIZAÇÃO DO CARTÃO (OBRIGATÓRIO SE INFORMAR O CNPJ DA OPERADORA)
        
    xStringNfce = xStringNfce & "119-001 = " & "FIM" & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "134-000 = " & "|@|" & vbCrLf 'NUMERO DA ECF
    
    xMensagem = CalculaImpostos(lNumeroNFCe, lDataNFCe) & "\n"
    xMensagem = xMensagem & "Cerrado Tecnologia - Soluções Inteligentes\n"
    xMensagem = xMensagem & "Fone: (62) 3277-1017\n"
    xStringNfce = xStringNfce & "135-000 = " & xMensagem & "|@|" & vbCrLf 'InfCpl - Informações complementares

'---- 'OBTER DADOS REAIS DO 210 AO 223 dados do cliente ---

    xStringNfce = xStringNfce & "210-000 = " & "|@|" & vbCrLf   '& RetiraAcentos(Cliente.RazaoSocial) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "211-000 = " & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "212-000 = " & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "213-000 = " & "|@|" & vbCrLf 'RetiraAcentos(Cliente.Cidade) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "214-000 = " & "|@|" & vbCrLf 'RetiraAcentos(Cliente.UF.ToUpper) & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "215-000 = " & "|@|" & vbCrLf 'Cliente.CEP & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "216-000 = 1058" & "|@|" & vbCrLf 'Codigo País
    
    xStringNfce = xStringNfce & "217-000 = BRASIL" & "|@|" & vbCrLf 'RetiraAcentos("BRASIL") & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "218-000 = " & "|@|" & vbCrLf '& Cliente.Telefone & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "219-000 = " & "|@|" & vbCrLf 'Cliente.CGC & "|@|" & vbCrLf 'cnpj

    xStringNfce = xStringNfce & "220-000 = " & "|@|" & vbCrLf  'inscrição estadual
    
    xStringNfce = xStringNfce & "221-000 = " & "|@|" & vbCrLf 'Codigo Suframa Cliente

    'xStringNfce = xStringNfce & "222-000 = 5208707" & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli GOIANIA
    xStringNfce = xStringNfce & "222-000 = " & xCodigoCidadeIbgeDestinatario & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli MORRINHOS
            
    xStringNfce = xStringNfce & "223-000 = " & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "224-000 = " & "9" & "|@|" & vbCrLf 'Indicador IE Destinatario. 9-Isento

'---- DADOS DA EMPRESA EMITENTE ----

     
     xStringNfce = xStringNfce & "310-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'RazaoSocialEmp

     xStringNfce = xStringNfce & "311-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'NomeFantasiaEmp
                
     xStringNfce = xStringNfce & "312-000 = " & xEmpresa.Endereco & "|@|" & vbCrLf  'Logradouro
                
     xStringNfce = xStringNfce & "313-000 = " & "0" & "|@|" & vbCrLf  'NumLgrEmp
     
     xStringNfce = xStringNfce & "314-000 = " & "" & "|@|" & vbCrLf  'CplEmp
     
     xStringNfce = xStringNfce & "315-000 = " & xEmpresa.Bairro & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Bairro) & vbCrLf 'BairroEmp
     
     xStringNfce = xStringNfce & "316-000 = " & xEmpresa.CEP & "|@|" & vbCrLf  'Empresa.CEP & vbCrLf 'CepEmp
     
     xStringNfce = xStringNfce & "317-000 = " & xEmpresa.Telefone & "|@|" & vbCrLf  '& Empresa.Telefone & vbCrLf 'TelefoneEmp
     
     xStringNfce = xStringNfce & "318-000 = " & xEmpresa.InscricaoEstadual & "|@|" & vbCrLf  'Empresa.InscricaoEstadual IEEmp
                
     xStringNfce = xStringNfce & "319-000 = " & xEmpresa.Cidade & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Cidade) & vbCrLf 'NomeMunEmp
                
     xStringNfce = xStringNfce & "320-000 = " & xCodigoCidadeIbgeEmitente & "|@|" & vbCrLf  '& xCodigoCidadeEmitente & vbCrLf 'CodMunEmp
                
     xStringNfce = xStringNfce & "321-000 = " & "" & "|@|" & vbCrLf  'CodMunIdentificacao
     
     xStringNfce = xStringNfce & "322-000 = " & "" & "|@|" & vbCrLf  'Empresa.EmailContador.ToLower) 'Email Contador(a)
     
     xStringNfce = xStringNfce & "323-000 = " & xCNPJEmpresa & "|@|" & vbCrLf ' 05577906000197 & Empresa.CGC & "|@|" & vbCrLf 'CNPJ
     
     xStringNfce = xStringNfce & "324-000 = " & xEmpresa.Estado & "|@|" & vbCrLf 'UF EMPRESA - ADICIONADO PARA NFCE

    
     '----------- NFCE 4.00
     
     xStringNfce = xStringNfce & "325-000 = 1" & "|@|" & vbCrLf 'TIPO NF (0-ENTRADA 1-SAIDA)
     
     xStringNfce = xStringNfce & "326-000 = VENDA" & "|@|" & vbCrLf 'NATUREZA OPERACAO
     
     xStringNfce = xStringNfce & "350-000 = " & "" & "|@|" & vbCrLf 'CnpjOuCpfTranspor
     xStringNfce = xStringNfce & "351-000 = " & "" & "|@|" & vbCrLf 'RazaoSocialTranspor
     xStringNfce = xStringNfce & "352-000 = " & "" & "|@|" & vbCrLf 'IETranspor
     xStringNfce = xStringNfce & "353-000 = " & "" & "|@|" & vbCrLf 'EndTranspor
     xStringNfce = xStringNfce & "354-000 = " & "" & "|@|" & vbCrLf 'NomeMunTranspor
     xStringNfce = xStringNfce & "355-000 = " & "" & "|@|" & vbCrLf 'UFTranspor
     xStringNfce = xStringNfce & "356-000 = " & "" & "|@|" & vbCrLf 'PlacaTranspor
     xStringNfce = xStringNfce & "357-000 = " & "" & "|@|" & vbCrLf 'UfPlacaTranspor
     xStringNfce = xStringNfce & "358-000 = 0" & "|@|" & vbCrLf '& FormatNumber(lTotalQtd, 0) & "|@|" & vbCrLf 'QntTranspor
     xStringNfce = xStringNfce & "359-000 = " & "|@|" & vbCrLf 'EspecieTranspor
     xStringNfce = xStringNfce & "360-000 = 0" & "|@|" & vbCrLf 'PesoLiqTranspor
     xStringNfce = xStringNfce & "361-000 = 0" & "|@|" & vbCrLf 'PesoBrutoTranspor
     xStringNfce = xStringNfce & "362-000 = 9" & "|@|" & vbCrLf 'TipoFrete
     xStringNfce = xStringNfce & "363-000 = " & "" & "|@|" & vbCrLf 'CodMunTranspor
     xStringNfce = xStringNfce & "364-000 = 0" & "|@|" & vbCrLf 'BCICMSTransp
     xStringNfce = xStringNfce & "365-000 = 0" & "|@|" & vbCrLf 'AliquotaIcmsTransp
     xStringNfce = xStringNfce & "366-000 = 0" & "|@|" & vbCrLf 'ValorICMSTransp
     xStringNfce = xStringNfce & "367-000 = " & "|@|" & vbCrLf 'CFOPTransp
     xStringNfce = xStringNfce & "368-000 = 0" & "|@|" & vbCrLf 'ValorServicoMCC

    
    MontaTextoCabecalhoSolicitacaoNFCE = xStringNfce

End Function
Private Function MontaTextoItensSolicitacaoNFCE(ByVal pRsDadosParaNFCe As adodb.Recordset) As String

    MontaTextoItensSolicitacaoNFCE = Empty
    Dim xOrdem As Integer
    Dim xStringNfce As String
     
    xOrdem = 0
    
    Do Until pRsDadosParaNFCe.EOF
    
        xOrdem = xOrdem + 1
        xStringNfce = xStringNfce & "800-" & Format(xOrdem, "000") & " = INICIO" & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "810-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo NCM").Value & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "811-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NomeProduto").Value & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "812-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Unidade").Value & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "813-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("IdProduto_MovDEItem").Value & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "820-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorUnitario_MovDEItem").Value, 4) & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "821-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("Quantidade_MovDEItem").Value, 4) & "|@|" & vbCrLf
        'xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("TotalVenda"), 4) & vbCrLf 'TotalProdServ
        xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, 2) & "|@|" & vbCrLf

        xStringNfce = xStringNfce & "823-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo de Barra").Value & "|@|" & vbCrLf '& Produto.CodigoBarra & vbCrLf
        
        xStringNfce = xStringNfce & "824-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CST PIS").Value & "|@|" & vbCrLf '& Format(Produto.CSTPIS, "00") & vbCrLf
        
        xStringNfce = xStringNfce & "825-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CST COFINS").Value & "|@|" & vbCrLf '& Format(Produto.CSTCOFINS, "00") & vbCrLf
        
        xStringNfce = xStringNfce & "826-" & Format(xOrdem, "000") & " = " & "" & "|@|" & vbCrLf 'CodListServico
        xStringNfce = xStringNfce & "827-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'ValorIva
        xStringNfce = xStringNfce & "828-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'TpIpi
        xStringNfce = xStringNfce & "829-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& lPisValor & "|@|" & vbCrLf 'ValorPisReais
        xStringNfce = xStringNfce & "830-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& lCofinsValor & "|@|" & vbCrLf 'ValorCofinsREais

        xStringNfce = xStringNfce & "831-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CFOP_MovDEItem").Value & "|@|" & vbCrLf '& dvItensNF(i)("CFOP") & "|@|" & vbCrLf 'CFOP
        xStringNfce = xStringNfce & "832-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf
        'Aliquota Cofins
        xStringNfce = xStringNfce & "833-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& xCodigoUfEmitente & "|@|" & vbCrLf 'CodEstadoIde
        
        xStringNfce = xStringNfce & "834-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& "0" & "|@|" & vbCrLf 'ValorIPIProd
        xStringNfce = xStringNfce & "835-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& lCofinsPercentual & "|@|" & vbCrLf 'ValorCofinsProd
        xStringNfce = xStringNfce & "836-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& lPisPercentual & "|@|" & vbCrLf 'ValorPISProd
        xStringNfce = xStringNfce & "837-" & Format(xOrdem, "000") & " = " & RetornaValorImpostoProdutoNFCE(pRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, pRsDadosParaNFCe("Aliquota do Imposto").Value) & "|@|" & vbCrLf '& FormatNumber(lIcmsValor, 4) & "|@|" & vbCrLf 'ValorIcmsProd
        'xStringNfce = xStringNfce & "838-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, 4) & "|@|" & vbCrLf  'ValorBCICMSProd
        xStringNfce = xStringNfce & "838-" & Format(xOrdem, "000") & " = " & FormatNumber(0, 4) & "|@|" & vbCrLf  'ValorBCICMSProd
        xStringNfce = xStringNfce & "839-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'AliquotaIPIProd
        xStringNfce = xStringNfce & "840-" & Format(xOrdem, "000") & " = 0" & FormatNumber(pRsDadosParaNFCe("Aliquota do Imposto").Value, 4) & "|@|" & vbCrLf 'AliquotaICMSProd
        
        xStringNfce = xStringNfce & "841-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorDesconto_MovDEItem").Value, 4) & "|@|" & vbCrLf  '& txtDesconto.Text & "|@|" & vbCrLf 'DescontoProd
        xStringNfce = xStringNfce & "842-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'Descontope
        xStringNfce = xStringNfce & "843-" & Format(xOrdem, "000") & " = " & "1" & "|@|" & vbCrLf 'tpNF (1-Saída)
        
        xStringNfce = xStringNfce & "844-" & Format(xOrdem, "000") & " = " & Format(pRsDadosParaNFCe("CST ICMS").Value, "00") & "|@|" & vbCrLf
        xStringNfce = xStringNfce & "845-" & Format(xOrdem, "000") & " = " & "P" & "|@|" & vbCrLf 'ServProd
        xStringNfce = xStringNfce & "846-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf 'DadosAdicionais
        xStringNfce = xStringNfce & "847-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'BCST
        xStringNfce = xStringNfce & "848-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'ValorIcmsSub
        xStringNfce = xStringNfce & "849-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualICMSSub
        xStringNfce = xStringNfce & "850-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMSSub
        xStringNfce = xStringNfce & "851-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMS
        
        
        'Valor do Desconto Total do Ítem
        xStringNfce = xStringNfce & "853-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("Valor do Desconto"), 4) & "|@|" & vbCrLf '
        'Valor do Acréscimo Total do Ítem
        xStringNfce = xStringNfce & "854-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(0, 4) & "|@|" & vbCrLf '

        'Tipo de Combustivel
        xStringNfce = xStringNfce & "855-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("TipoCombustivel_MovDEItem").Value & "|@|" & vbCrLf   '& Produto.TipoCombustivel & "|@|" & vbCrLf
        
        'Codigo da ANP
        xStringNfce = xStringNfce & "856-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo ANP").Value & "|@|" & vbCrLf  '620505001 & Produto.CodigoANP & "|@|" & vbCrLf
        
        'Codigo CEST
        xStringNfce = xStringNfce & "857-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CEST").Value & "|@|" & vbCrLf '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
        
        'ENCERRANTE INICIAL 'CALCULADO PELO ENCERRANTE FINAL - QUANTIDADE
        Dim xEncerranteFinal As Currency
        xEncerranteFinal = pRsDadosParaNFCe("EncerranteFinal_MovDEItem").Value
        xStringNfce = xStringNfce & "858-" & Format(xOrdem, "000") & " = " & RetornaEncerranteInicial(xEncerranteFinal, pRsDadosParaNFCe("Quantidade_MovDEItem").Value) & "|@|" & vbCrLf    '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
        
        'ENCERRANTE FINAL 'ESTÁ SENDO GRAVADO NO CAMPO DE TELEFONE DA TABELA MOVIMENTO_CUPOOM_FISCAL
        xStringNfce = xStringNfce & "859-" & Format(xOrdem, "000") & " = " & xEncerranteFinal & "|@|" & vbCrLf
        
        'Bomba ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 1 e 2) DA TABELA MOVIMENTO_CUPOOM_FISCAL
        xStringNfce = xStringNfce & "860-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBomba_MovDEItem").Value & "|@|" & vbCrLf  'Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 1, 2) & "|@|" & vbCrLf
        
        'Bico ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 3 e 4) DA TABELA MOVIMENTO_CUPOOM_FISCAL
        xStringNfce = xStringNfce & "861-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBico_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 3, 2) & "|@|" & vbCrLf
        
        'tanque ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 5 e 6) DA TABELA MOVIMENTO_CUPOOM_FISCAL
        xStringNfce = xStringNfce & "862-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroTanque_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 5, 2) & "|@|" & vbCrLf
        
        'CSTCSOSN
        xStringNfce = xStringNfce & "863-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf

        xStringNfce = xStringNfce & "899-" & Format(xOrdem, "000") & " = FIM" & 1 & "|@|" & vbCrLf
        
        pRsDadosParaNFCe.MoveNext
     Loop
     
     xStringNfce = xStringNfce & "999-999 = 0" & "|@|" & vbCrLf
     
     MontaTextoItensSolicitacaoNFCE = xStringNfce


End Function
'Private Function ObtenhaDadosParaNFCE(ByVal pNumeroCupom As Long, ByVal pDataCupom As Date) As adodb.Recordset
'
'    Dim rsDadosParaNFCe As New adodb.Recordset
'
'    Dim i As Integer
'    Dim xSQL As String
'    Dim xTextoSolicitacao As String
'
'    'O campo Telefone está sendo preenchido com Encerrante Final para utilização na NFCE
'    'O campo Numero do Cheque está sendo preenchido com valores concatenados da bomba, Bico e Tanque
'
'    xSQL = ""
'    xSQL = xSQL & "SELECT Movimento_Cupom_Fiscal.Empresa, Movimento_Cupom_Fiscal.[Numero do Cupom],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Data do Cupom], Movimento_Cupom_Fiscal.Ordem,"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Hora, Movimento_Cupom_Fiscal.Periodo,"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo do Movimento], Movimento_Cupom_Fiscal.[Codigo do Cliente],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Codigo do Conveniado], Movimento_Cupom_Fiscal.[Codigo do Produto],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Unitario], Movimento_Cupom_Fiscal.Quantidade,"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Total], Movimento_Cupom_Fiscal.[Forma de Pagamento],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Recebido], Movimento_Cupom_Fiscal.[Valor do Desconto],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Desconto Embutido], Movimento_Cupom_Fiscal.[Numero do Cheque],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Operador], Movimento_Cupom_Fiscal.[Codigo da Aliquota],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Nome, Movimento_Cupom_Fiscal.Telefone, "
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[CPF CNPJ], Movimento_Cupom_Fiscal.[Tipo de Combustivel],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Codigo da Ecf], Movimento_Cupom_Fiscal.[Codigo do Grupo],"
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo do SubEstoque],"
'
'    xSQL = xSQL & "       Movimento_Cupom_Fiscal_Item.[Valor do Desconto], Movimento_Cupom_Fiscal_Item.[Valor do Acrescimo],"
'
'    xSQL = xSQL & "       Produto.Nome AS NomeProduto, Produto.Unidade, Produto.[Codigo de Barra],"
'    xSQL = xSQL & "       Produto.[CST ICMS], Produto.[CST PIS], Produto.[CST COFINS],"
'    xSQL = xSQL & "       Produto.[Codigo NCM], Produto.[Codigo ANP], Produto.CEST,"
'    xSQL = xSQL & "       Aliquota.[Codigo Fiscal] , Aliquota.[Aliquota do Imposto]"
'    xSQL = xSQL & "  FROM [movimento_cupom_fiscal], Movimento_Cupom_Fiscal_Item, Produto, Aliquota"
'    xSQL = xSQL & " WHERE movimento_cupom_fiscal.Empresa = " & g_empresa
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Data do Cupom] = " & preparaData(pDataCupom)
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = " & pNumeroCupom
'
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
'
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Empresa = Movimento_Cupom_Fiscal_Item.Empresa"
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Data do Cupom] = Movimento_Cupom_Fiscal_Item.Data"
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = Movimento_Cupom_Fiscal_Item.[Numero do Cupom]"
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Ordem = Movimento_Cupom_Fiscal_Item.Ordem"
'
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"
'
'    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo da Aliquota] = Aliquota.Codigo"
'    xSQL = xSQL & "   AND Aliquota.[Serie ECF] = " & preparaTexto(lSerieECF)
'
'    xSQL = xSQL & " ORDER BY Movimento_Cupom_Fiscal.Ordem"
'
'    'Abre RecordSet
'    Set rsDadosParaNFCe = New adodb.Recordset
'    Set rsDadosParaNFCe = Conectar.RsConexao(xSQL)
'
'
'    Set ObtenhaDadosParaNFCE = rsDadosParaNFCe
'
'End Function
Private Function ObtenhaDadosParaNFCEDocumentoEletronico(ByVal pNumeroNFCe As Long, ByVal pDataEmissao As Date) As adodb.Recordset

    Dim rsDadosParaNFCe As New adodb.Recordset
    
    Dim i As Integer
    Dim xSQL As String
    Dim xTextoSolicitacao As String
   
    'O campo Telefone está sendo preenchido com Encerrante Final para utilização na NFCE
    'O campo Numero do Cheque está sendo preenchido com valores concatenados da bomba, Bico e Tanque
    
    xSQL = ""
    xSQL = xSQL & "SELECT IdEstabelecimento_MovDECabecalho, Numero_MovDECabecalho,"
    xSQL = xSQL & "DataEmissao_MovDECabecalho,"
    xSQL = xSQL & "Ordem_MovDEItem,"
    xSQL = xSQL & "HoraSaida_MovDECabecalho,"
    xSQL = xSQL & "IdClienteFornecedor_MovDECabecalho,"
    xSQL = xSQL & "IdProduto_MovDEItem,"
    xSQL = xSQL & "ValorUnitario_MovDEItem,"
    xSQL = xSQL & "Quantidade_MovDEItem,"
    xSQL = xSQL & "ValorTotal_MovDECabecalho,"
    xSQL = xSQL & "FormaPagamento_MovDECabecalho,"
    xSQL = xSQL & "ValorTotalLiquido_MovDEItem,"
    xSQL = xSQL & "ValorDesconto_MovDECabecalho,"
    xSQL = xSQL & "IdUsuario_MovDECabecalho,"
    xSQL = xSQL & "ValorDesconto_MovDEItem,"
    xSQL = xSQL & "EncerranteFinal_MovDEItem,"
    xSQL = xSQL & "NumeroBomba_MovDEItem,"
    xSQL = xSQL & "NumeroBico_MovDEItem,"
    xSQL = xSQL & "NumeroTanque_MovDEItem,"
    xSQL = xSQL & "TipoCombustivel_MovDEItem,"
    xSQL = xSQL & "CFOP_MovDEItem,"
    xSQL = xSQL & "Produto.Nome As NomeProduto,"
    xSQL = xSQL & "Produto.Unidade,"
    xSQL = xSQL & "Produto.[Codigo de Barra],"
    xSQL = xSQL & "Produto.[CST ICMS],"
    xSQL = xSQL & "Produto.[CST PIS],"
    xSQL = xSQL & "Produto.[CST COFINS],"
    xSQL = xSQL & "Produto.[Codigo NCM],"
    xSQL = xSQL & "Produto.[Codigo ANP],"
    xSQL = xSQL & "Produto.CEST,"
    xSQL = xSQL & "Aliquota.[Codigo Fiscal] ,"
    xSQL = xSQL & "Aliquota.[Aliquota do Imposto]"
    xSQL = xSQL & " FROM Produto, MovimentoDocumentoEletronicoCabecalho,"
    xSQL = xSQL & " Aliquota, MovimentoDocumentoEletronicoItem"
    xSQL = xSQL & " WHERE IdEstabelecimento_MovDECabecalho =" & g_empresa
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    xSQL = xSQL & " AND Numero_MovDECabecalho = " & pNumeroNFCe
    xSQL = xSQL & " AND Cancelado_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Cancelado_MovDEItem = " & preparaBooleano(False)
    xSQL = xSQL & " AND IdEstabelecimento_MovDECabecalho = IdEstabelecimento_MovDEItem   "
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = DataEmissao_MovDEItem   "
    xSQL = xSQL & " AND Numero_MovDECabecalho = Numero_MovDEItem   "
    xSQL = xSQL & " AND IdProduto_MovDEItem = Produto.Codigo   "
    xSQL = xSQL & " AND Produto.[Codigo da Aliquota] = Aliquota.Codigo   "
    xSQL = xSQL & " AND Aliquota.[Serie ECF] = " & preparaTexto(lSerieECF)
    xSQL = xSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(True)
    xSQL = xSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(MODELO_NFCE)
    xSQL = xSQL & " AND Serie_MovDECabecalho = " & preparaTexto(lSerieNFCe)
    xSQL = xSQL & "ORDER BY Ordem_MovDEItem"
    

    'Abre RecordSet
    Set rsDadosParaNFCe = New adodb.Recordset
    Set rsDadosParaNFCe = Conectar.RsConexao(xSQL)
    
    
    Set ObtenhaDadosParaNFCEDocumentoEletronico = rsDadosParaNFCe

End Function
Private Sub Relatorio()
    ZeraVariaveis
    
    TotalizaAcertoVendaNFCe
    'Rotina abaixo, nao precisou terminar o desenvolvimento
    'LoopVerificaDuplicidadeBaixa
    LoopVerificaArredondamentoAutomacao
    
    If lCaixaIndividual And chbPorFuncionario.Value = 1 Then
        'Verifica Movimento_Abastecimento
        lSQL = ""
        lSQL = lSQL & "SELECT Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
        lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 0 AS Baixado"
        lSQL = lSQL & "  FROM Movimento_Abastecimento"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text) - 5)
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
        'lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
        lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
'        If g_nome_empresa = "AUTO POSTO MOREIRA COSTA LTDA" Then
'
'            lSQL = lSQL & " UNION "
'
'            lSQL = lSQL & "SELECT TOP 1000 Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
'            lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 1 AS Baixado"
'            lSQL = lSQL & "  FROM BaixaAbastecimento"
'            lSQL = lSQL & " WHERE Empresa = " & g_empresa
'            lSQL = lSQL & "   AND Data >= " & preparaData(CDate("01/01/2017")) 'preparaData(CDate(msk_data_i.Text))
'            lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
'            lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
'        End If
        lSQL = lSQL & " ORDER BY Data, Hora, Bico"
        Set rstMovimentoAbastecimento = Conectar.RsConexao(lSQL)
        If Not rstMovimentoAbastecimento.EOF Then
            ImpDados
            If g_automacao = True And lLocal = 1 Then
                If Not MovimentoAbastecimento.DescarregarAbastecimentoFuncionario(g_empresa, CDate(msk_data_i.Text), lNumeroPDV, "CP", lCodigoFuncionario) Then
                    MsgBox "Não foi possível descarregar os abastecimentos deste funcionário!", vbInformation, "Erro de Descarregamento"
                End If
            End If
            If lDataHoraInicioProcessamento <> CDate("00:00:00") Then
                lProcessamentoNFCeComplementar.DataHoraTermino = Now
                Call lProcessamentoNFCeComplementar.DefinirDataHoraTermino(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV, "FINALIZADO COM SUCESSO")
                lTerminoDefinido = True
            End If
        Else
            Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Informado ao Usuário que Não Existia abastecimento deste funcionário no Período.")
            If lDataHoraInicioProcessamento <> CDate("00:00:00") Then
                lProcessamentoNFCeComplementar.DataHoraTermino = Now
                Call lProcessamentoNFCeComplementar.DefinirDataHoraTermino(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV, "FINALIZADO SEM DADOS A EMITIR")
                lTerminoDefinido = True
            End If
            MsgBox "Não existe movimento de abastecimento deste funcionário no período.", vbInformation, "Erro de Verificação!"
        End If
        rstMovimentoAbastecimento.Close
    Else
        'Verifica Movimento_Abastecimento
'        lSQL = ""
'        lSQL = lSQL & "SELECT Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
'        lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 0 AS Baixado"
'        lSQL = lSQL & "  FROM Movimento_Abastecimento"
'        lSQL = lSQL & " WHERE Empresa = " & g_empresa
'        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text) - 5)
'        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
'        'lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
'        lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
        If g_nome_empresa = "xxxxAUTO POSTO MOREIRA COSTA LTDA" Then
            lSQL = ""
'            lSQL = lSQL & "SELECT Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
'            lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 0 AS Baixado"
'            lSQL = lSQL & "  FROM Movimento_Abastecimento"
'            lSQL = lSQL & " WHERE Empresa = " & g_empresa
'            lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text) + 5)
'            lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text) + 5)
'            'lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
'            lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
'
'            lSQL = lSQL & " UNION "
'
            lSQL = lSQL & "SELECT TOP 1000 Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
            lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 1 AS Baixado"
            lSQL = lSQL & "  FROM BaixaAbastecimento"
            lSQL = lSQL & " WHERE Empresa = " & g_empresa
            lSQL = lSQL & "   AND Data >= " & preparaData(CDate("01/12/2016")) 'preparaData(CDate(msk_data_i.Text))
            lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
            'lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
            lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
        Else
            lSQL = ""
            lSQL = lSQL & "SELECT Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
            lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario], 0 AS Baixado"
            lSQL = lSQL & "  FROM Movimento_Abastecimento"
            lSQL = lSQL & " WHERE Empresa = " & g_empresa
            lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text) - 5)
            lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
            'lSQL = lSQL & "   AND [Codigo do Funcionario] = " & lCodigoFuncionario
            lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
        End If
        lSQL = lSQL & " ORDER BY Data, Hora, Bico"
        Set rstMovimentoAbastecimento = Conectar.RsConexao(lSQL)
        If Not rstMovimentoAbastecimento.EOF Then
            ImpDados
            If lDataHoraInicioProcessamento <> CDate("00:00:00") Then
                lProcessamentoNFCeComplementar.DataHoraTermino = Now
                Call lProcessamentoNFCeComplementar.DefinirDataHoraTermino(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV, "FINALIZADO COM SUCESSO")
                lTerminoDefinido = True
            End If
            'Abaixo melhorar codigo, deve mudar abastecimento na medida que montar a nfce
'            If g_automacao = True And lLocal = 1 Then
'                If Not MovimentoAbastecimento.DescarregarAbastecimento(g_empresa, CDate(msk_data_i.Text), lNumeroPDV, "CP") Then
'                    MsgBox "Não foi possível descarregar os abastecimentos!", vbInformation, "Erro de Descarregamento"
'                End If
'            End If
        Else
            Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Informado ao Usuário que Não Existia Movimento de Bomba Digitado no Período.")
            
            If lDataHoraInicioProcessamento <> CDate("00:00:00") Then
                lProcessamentoNFCeComplementar.DataHoraTermino = Now
                Call lProcessamentoNFCeComplementar.DefinirDataHoraTermino(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV, "FINALIZADO SEM DADOS A EMITIR")
                lTerminoDefinido = True
            End If
            
            MsgBox "Não existe movimento de abastecimento no período.", vbInformation, "Erro de Verificação!"
        End If
        rstMovimentoAbastecimento.Close
    End If
    Call AtivaBotoes(True)
    cmd_sair.SetFocus
End Sub
Private Function RetornaEncerranteInicial(pEncerranteFinal As Currency, pQuantidadeSaida As Currency)
    If pEncerranteFinal <= 0 Then
        RetornaEncerranteInicial = 0
        Exit Function
    End If
    RetornaEncerranteInicial = pEncerranteFinal - pQuantidadeSaida
End Function
Private Function RetornaValorImpostoProdutoNFCE(ByVal pValorBaseCalculo As Currency, ByVal pAliquotaImposto As Currency) As Currency
    If Val(pAliquotaImposto) = 0 Then
        RetornaValorImpostoProdutoNFCE = pAliquotaImposto
        Exit Function
    End If
    
    Dim xAliquota As Currency
    
    xAliquota = pAliquotaImposto / 100
    
    RetornaValorImpostoProdutoNFCE = pValorBaseCalculo * pAliquotaImposto
End Function
'Private Sub LoopVerificaDuplicidadeBaixa()
'    Dim rstBaixaDuplicadas As New adodb.Recordset
'    Dim rstBaixaAbast As New adodb.Recordset
'
'    lSQL = "SELECT Data, Hora, Bico, Count(1) AS Qtd"
'    lSQL = lSQL & "  FROM BaixaAbastecimento"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data = " & preparaData(CDate("29/01/2017"))
'    lSQL = lSQL & " GROUP BY Data, Hora, Bico"
'    lSQL = lSQL & " ORDER BY Qtd Desc, Data, Hora, Bico"
'    Set rstBaixaDuplicadas = Conectar.RsConexao(lSQL)
'    If rstBaixaDuplicadas.RecordCount > 0 Then
'        rstBaixaDuplicadas.MoveFirst
'        Do Until rstBaixaDuplicadas.EOF
'            If rstBaixaDuplicadas!qtd = 2 Then
'
'
'                lSQL = "SELECT *"
'                lSQL = lSQL & "  FROM BaixaAbastecimento"
'                lSQL = lSQL & " WHERE Empresa = " & g_empresa
'                lSQL = lSQL & "   AND Data = " & preparaData(CDate("29/01/2017"))
'                lSQL = lSQL & " GROUP BY Data, Hora, Bico"
'                lSQL = lSQL & " ORDER BY Qtd Desc, Data, Hora, Bico"
'                Set rstBaixaAbast = Conectar.RsConexao(lSQL)
'                If rstBaixaAbast.RecordCount > 0 Then
'                    rstBaixaAbast.MoveFirst
'                    Do Until rstBaixaAbast.EOF
'                        If rstBaixaAbast!qtd = 2 Then
'                        End If
'                        rstBaixaAbast.MoveNext
'                    Loop
'                End If
'
'
'
'
'            End If
'            rstBaixaDuplicadas.MoveNext
'        Loop
'    End If
'    rstBaixaDuplicadas.Close
'    Set rstBaixaDuplicadas = Nothing
'
'
'
'            End If
'
'            rstBaixaDuplicadas.MoveNext
'        Loop
'    End If
'    rstBaixaDuplicadas.Close
'    Set rstBaixaDuplicadas = Nothing
'End Sub
Private Sub LoopVerificaArredondamentoAutomacao()
    Dim xDiferenca As Currency
    Dim xValorTotal As Currency
    
    lSQL = ""
    lSQL = lSQL & "SELECT Data, Periodo, Hora, [Codigo do Produto], Bico, [Tipo de Combustivel], [Valor Unitario], Quantidade, "
    lSQL = lSQL & "[Valor Total], [Encerrante Inicial], Encerrante, [Codigo do Funcionario]"
    lSQL = lSQL & "  FROM Movimento_Abastecimento"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Acerto = " & preparaBooleano(False)
    lSQL = lSQL & " ORDER BY Data, Hora, Bico"
    Set rstMovimentoAbastecimento = Conectar.RsConexao(lSQL)
    If rstMovimentoAbastecimento.RecordCount > 0 Then
        rstMovimentoAbastecimento.MoveFirst
        Do Until rstMovimentoAbastecimento.EOF
            xValorTotal = Round(rstMovimentoAbastecimento!Quantidade * rstMovimentoAbastecimento![Valor Unitario], 2)
            xDiferenca = rstMovimentoAbastecimento![Valor Total] - xValorTotal
            If xDiferenca <> 0 Then
                If MovimentoAbastecimento.LocalizarCodigo(g_empresa, rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico) Then
                    MovimentoAbastecimento.Quantidade = Round(MovimentoAbastecimento.ValorTotal / MovimentoAbastecimento.ValorUnitario, 3)
                    If Not MovimentoAbastecimento.Alterar(g_empresa, rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico) Then
                        MsgBox "Erro em: LoopVerificaArredondamentoAutomacao" & vbCrLf & "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    MsgBox "Erro em: LoopVerificaArredondamentoAutomacao" & vbCrLf & "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
                End If
            End If
            If lExisteProgramacaoADiminuir = True Then
                Call ExecutaAcertoVendaProgramada(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, rstMovimentoAbastecimento!Quantidade, rstMovimentoAbastecimento![Valor Total], rstMovimentoAbastecimento![Tipo de Combustivel])
            End If
            rstMovimentoAbastecimento.MoveNext
        Loop
    End If
End Sub
Private Sub ExecutaAcertoVendaProgramada(ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pTipoCombustivel As String)
    Dim i As Integer
    
    For i = 1 To 7
        If lTipoCombustivel(i) = pTipoCombustivel Then
            If lQtdBombaV(i) < 0 Then
                If pQuantidade < (lQtdBombaV(i) * -1) Then
                    lQtdBombaV(i) = lQtdBombaV(i) + pQuantidade
                Else
                    lQtdBombaV(i) = 0
                End If
                If MovimentoAbastecimento.LocalizarCodigo(g_empresa, pData, pHora, pBico) Then
                    MovimentoAbastecimento.Acerto = True
                    MovimentoAbastecimento.NumeroCupom = 2
                    MovimentoAbastecimento.CodigoECF = 222
                    MovimentoAbastecimento.DocumentoGerado = "AVP"
                    'AFTEMP - Abastecimento selecionado para vincular a aferição
                    'AFERICAO - Abastecimento vinculado à Afericao
                    'AVP  - Acerto de Venda Programada
                    'CF   - Cupom Fiscal
                    'NFCe - Nota Fiscal do Consumidor Eletrônica
                    'NT   - Nota Abastecimento
                    'CP   - Cupom Complementar
                    'AF   - Afericao
                    'CHVIS- Cheque A Vista
                    'CHPRE- Cheque Pre-Datado
                    'CRT  - Cartao de Credito
                    'DIN  - Dinheiro
                    'DESPC- Despesa de Caixa
                    'VALEF- Vale de Funcionario
                    'VLABR- Vale Abastecimento Recebido
                    'CRAR - Credito Antecipado Recebido
                    If Not MovimentoAbastecimento.Alterar(g_empresa, pData, pHora, pBico) Then
                        MsgBox "Não foi possível alterar o abastecimento AVP!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    MsgBox "Não foi possível localizar abastecimento AVP!", vbInformation, "Erro de Integridade!"
                End If
                Exit For
                'lQtdBombaV(i) = lQtdBombaV(i) + !Quantidade
                'lTotalBombaV(i) = lTotalBombaV(i) + !ValorTotal
            End If
        End If
    Next
End Sub
Private Sub ImpDados()
    'Dim i As Integer
    Dim xPrecoCusto As Currency
    Dim xPrecoVenda As Currency
    Dim xPrecoUsado As Currency
    Dim xValorTotalCupom As Currency
    Dim xTotalItem As Currency
    Dim xQuantidadeAutomacao As Currency
    
    
    
    'Dim xQtd As Currency
    'Dim xValor As Currency
    'Dim xStrConfiguracaoDiversa As String
    'Dim xStrLogAuditoria As String
    'Dim xQtdLimite As Currency
    'Dim xQtdNaoImpressa As Currency
    'Dim xQtdDiferenca As Currency
    'Dim xTotalDiferenca As Currency
    
    xValorTotalCupom = 0
    rstMovimentoAbastecimento.MoveFirst
    Do Until rstMovimentoAbastecimento.EOF
        xQuantidadeAutomacao = rstMovimentoAbastecimento!Quantidade
        xQuantidadeAutomacao = AutomacaoAlteraArredondamento(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, rstMovimentoAbastecimento![Valor Unitario], xQuantidadeAutomacao, rstMovimentoAbastecimento![Valor Total], rstMovimentoAbastecimento!Baixado)
        
        
    
        'Limita em R$ 9.500,00
        'Limita em 30 Ítens por NFCe
'        If g_nome_empresa = "++++++++++++AUTO POSTO BRISA LTDA+++++++++++" And rstMovimentoAbastecimento![Tipo de Combustivel] = "A " Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "++++++++++++++AUTO POSTO CLASSE A LTDA+++++++++++++++" Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "+++++++++++++++MARQUES DE CASTRO & GABRIEL LTDA+++++++++++++++" And (rstMovimentoAbastecimento![Tipo de Combustivel] = "A " Or rstMovimentoAbastecimento![Tipo de Combustivel] = "DA" Or rstMovimentoAbastecimento![Tipo de Combustivel] = "GA") Then 'ESMERALDA
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "++++++++++++++++POSTO NOVO HORIZONTE LTDA+++++++++++++++++" Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "++++++++++++++++AUTO POSTO CRISTO REI E CONVENIENCIA ME+++++++++++++++++" And (rstMovimentoAbastecimento![Tipo de Combustivel] = "G " Or rstMovimentoAbastecimento![Tipo de Combustivel] = "D ") Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "+++++++++++++++++++AUTO POSTO T13 LTDA++++++++++++++++++++" And (rstMovimentoAbastecimento![Tipo de Combustivel] = "A " Or rstMovimentoAbastecimento![Tipo de Combustivel] = "G ") Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "++++++++++++++AUTO POSTO MT LTDA+++++++++++++++++" And rstMovimentoAbastecimento![Tipo de Combustivel] = "G " Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        ElseIf g_nome_empresa = "++++++++++++++VALPOSTO COMBUSTIVEIS LTDA++++++++++++++++" Then
'            If lLocal = 1 Then
'                Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, 1, rstMovimentoAbastecimento!Baixado)
'            End If
'        Else
            If lLocal = 1 Then
                Call CriaLogCupom(Time & " - ImpDados: lOrdemNFCe: " & lOrdemNFCe)
            
                If (xValorTotalCupom + rstMovimentoAbastecimento![Valor Total]) > 9500 Or lOrdemNFCe = 30 Then
                    Call CriaLogCupom(Time & " - ImpDados: EnviaDadosParaNFCe - Inicio")
                    Call EnviaDadosParaNFCe(lNumeroNFCe, lDataNFCe, xValorTotalCupom)
                    Call CriaLogCupom(Time & " - ImpDados: EnviaDadosParaNFCe - Fim")
                    
    '                lParar = lParar + 1
    '                If lParar < 5 Then
    '                    MsgBox "Teste para verificar autorização NFCe N. " & lNumeroNFCe
    '                End If
                    
                    If lNFCeComplementarSemRetorno = False Then
                        AguardaProcessamentoNFCe MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, MovDocEletronicoItem.TipoCombustivel
                    End If
                    
                    lNumeroNFCe = 0
                    lOrdemNFCe = 0
                    xValorTotalCupom = 0
                End If
            End If
            xPrecoCusto = 0
            xPrecoVenda = 0
            xPrecoUsado = 0
            If Produto.LocalizarCodigo(rstMovimentoAbastecimento![Codigo do Produto]) = True Then
                If Estoque.LocalizarCodigo(g_empresa, rstMovimentoAbastecimento![Codigo do Produto]) Then
                    xPrecoCusto = Produto.PrecoCusto
                    xPrecoVenda = Estoque.PrecoVenda
                    xPrecoUsado = Estoque.PrecoVenda
                Else
                    MsgBox "Não foi possível localizar o Preço do Produto!", vbCritical, "Erro Fatal!"
                    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Erro ao localizar preço de venda do produto: " & rst2!Codigo)
                    Call GravaAuditoria(1, Me.name, 26, "Erro ao localizar preço do produto=" & rst2!Codigo)
                    End
                End If
                xPrecoUsado = rstMovimentoAbastecimento![Valor Unitario]
                If Combustivel.LocalizarCodigo(g_empresa, rstMovimentoAbastecimento![Tipo de Combustivel]) = True Then
                    If Combustivel.PrecoMedio > xPrecoCusto And Combustivel.PrecoMedio < xPrecoVenda Then
                        If Combustivel.PrecoMedio < rstMovimentoAbastecimento![Valor Unitario] Then
                            xPrecoUsado = Combustivel.PrecoMedio
                        Else
                            xPrecoUsado = rstMovimentoAbastecimento![Valor Unitario]
                        End If
                    End If
                End If
               
                xTotalItem = rstMovimentoAbastecimento![Valor Total]
                'Recalcula Total do Item
                If xPrecoUsado < rstMovimentoAbastecimento![Valor Unitario] Then
                    xTotalItem = Round(xQuantidadeAutomacao * xPrecoUsado, 2)
                End If
                xValorTotalCupom = xValorTotalCupom + xTotalItem
                If lLocal = 1 Then
                    Call GravaDadosNFCe(rstMovimentoAbastecimento![Codigo do Produto], rstMovimentoAbastecimento![Tipo de Combustivel], rstMovimentoAbastecimento!Bico, rstMovimentoAbastecimento!Encerrante, xPrecoUsado, xQuantidadeAutomacao, xTotalItem, xValorTotalCupom)
                    Call AlteraAbastecimentoParaConcluido(rstMovimentoAbastecimento!Data, rstMovimentoAbastecimento!Hora, rstMovimentoAbastecimento!Bico, lNumeroNFCe, rstMovimentoAbastecimento!Baixado)
                End If
            End If
'        End If
        
        rstMovimentoAbastecimento.MoveNext
    Loop
    If lNumeroNFCe > 0 Then
        If lLocal = 1 Then
            Call EnviaDadosParaNFCe(lNumeroNFCe, lDataNFCe, xValorTotalCupom)
            lNumeroNFCe = 0
            lOrdemNFCe = 0
            xValorTotalCupom = 0
            
            If lNFCeComplementarSemRetorno = False Then
                Call AguardaProcessamentoNFCe(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, MovDocEletronicoItem.TipoCombustivel)
            End If
        End If
    End If
    
End Sub

Private Function GravaDocumentoEletronicoEvento(ByVal pDocumentoEltronicoCabecalho As cMovDocEletronicoCabecalho, ByVal pDescricaoEvento As EVENTO_NFCE) As Boolean

    On Error GoTo FileError

    Dim xMovEvento As New cMovDocEletronicoEvento
    Dim xPassoProcesso As String
    
    xPassoProcesso = ""
    
    xPassoProcesso = "1"
    With xMovEvento
        .IdEstabelecimento = pDocumentoEltronicoCabecalho.IdEstabelecimento
        .Modelo = pDocumentoEltronicoCabecalho.Modelo
        .numero = Val(lNumeroNFCe)
        .Serie = pDocumentoEltronicoCabecalho.Serie
        .DataEmissao = pDocumentoEltronicoCabecalho.DataEmissao
        .Sequencia = xMovEvento.ProximaSequencia(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, lNumeroNFCe)
        .DataHora = Now
        .CodigoTipoEvento = Val(pDescricaoEvento)
        .Descricao = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
    End With
    
    xPassoProcesso = "2"
    If xMovEvento.Incluir Then
        xPassoProcesso = "3"
        pDocumentoEltronicoCabecalho.CodigoUltimoEvento = Val(pDescricaoEvento)
        pDocumentoEltronicoCabecalho.ObservacaoEvento = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
        Call pDocumentoEltronicoCabecalho.DefinirUtlimoEventoDocumento(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, pDocumentoEltronicoCabecalho.numero)
    Else
        Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & " xPassoProcesso=" & xPassoProcesso)
        MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ".", vbInformation, "Erro de Integridade."
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: " & Error & " xPassoProcesso=" & xPassoProcesso)
    MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ". ERRO= " & Err.Description, vbCritical, "Erro de Integridade."

End Function


Private Sub AguardaProcessamentoNFCe(ByVal pNSU As Long, ByVal pTipoCombustivel As String)
    Dim xHoraInicial As Date
    Dim i As Integer
    Dim xProcessamentoConcluido As Boolean

    lblTitulo.Caption = "Aguarde! Processando NFCe... NSU(" & pNSU & ")"
    lblMensagem.Caption = "Emissão NFC-e Complementar: " & pTipoCombustivel
    FrameAguarde.ZOrder 0
    FrameAguarde.Visible = True
    
    DoEvents
    xProcessamentoConcluido = False
    
 
    xHoraInicial = Time
    'Fica até 60 segundos
    Do Until DateDiff("s", xHoraInicial, Time) >= 60
        Call AguardaMS(1000)
        If MovSolicitacaoFuncaoNFe.LocalizarNSU(g_empresa, pNSU) Then
            If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Or MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
                xProcessamentoConcluido = True
                Exit Do
            End If
        End If
        DoEvents
    Loop
    
    DoEvents
    FrameAguarde.ZOrder 1
    FrameAguarde.Visible = False
    DoEvents
    
    If xProcessamentoConcluido = True Then
        If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            Dim xMensagemNFCe As String
                        
            xMensagemNFCe = MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe
            'MsgBox "NFCe Complementar AUTORIZADA!" & vbCrLf & "Mensagem: " & xMensagemNFCe, vbInformation, "Processamento Concluído!"
        ElseIf MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            MsgBox "NFCe (Nº" & MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe & ") NÃO foi Autorizada!" & vbCrLf & "Mensagem: " & MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe, vbCritical, "ERRO ao Processar NFCe Complementar!"
        Else
            MsgBox "Não será possível definir o processamento da NFCe Complementar.", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Tempo de solicitação de Processamento de NFCe Complementar excedido." & vbCrLf & "Não obtivemos retorno esperado!", vbCritical, "Tempo Excedido!"
    End If
End Sub
Private Sub ImpTermicaAbreRelatorio()
    lNomeArquivo = BioCriaImprime
    'seleciona medidas para centímetros
    BioImprime "@@Printer.ScaleMode = 7"
    BioImprime "@@Printer.PaperSize = 1"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    'teste para imprimir letra correta
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    Printer.FontName = "Lucida Console 7cpi"
    BioImprime "@@Printer.FontName = Lucida Console 7cpi"
    BioImprime "@@Printer.CurrentY = 0"
End Sub
Private Sub ImpTermicaImprimeDados(ByVal pLinhaDados As String, ByVal pNegrito As Boolean)
    Dim xNegrito As String
    
    If pNegrito = True Then
        xNegrito = "True"
    Else
        xNegrito = "False"
    End If
    BioImprime "@Printer.Print " & pLinhaDados
    BioImprime "@@Printer.FontBold = " & xNegrito
End Sub
Private Sub ImpTermicaCabecalhoPosto()
    Dim xEmpresa As New cEmpresa
    Dim xLinhaDados As String
    Dim xDados As String
    Dim i As Integer
    
    If xEmpresa.LocalizarCodigo(g_empresa) = False Then
        Exit Sub
    End If
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call ImpTermicaImprimeDados(" ", True)
'    Call ImpTermicaImprimeDados(xEmpresa.Nome, True)
'    Call ImpTermicaImprimeDados("CNPJ: " & fMascaraCNPJ(xEmpresa.CGC) & "  IE: " & xEmpresa.InscricaoEstadual, True)
'    Call ImpTermicaImprimeDados(xEmpresa.Endereco, True)
'    Call ImpTermicaImprimeDados(xEmpresa.Cidade & "-" & xEmpresa.Estado, True)
'    Call ImpTermicaImprimeDados("CEP: " & fMascaraCEP(xEmpresa.CEP) & "  FONE: " & fMascaraTelefone(xEmpresa.Telefone), True)
    
    xLinhaDados = Space(48)
    i = Len(Trim(xEmpresa.Nome))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xEmpresa.Nome)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CNPJ: " & fMascaraCNPJ(xEmpresa.CGC) & "  IE: " & xEmpresa.InscricaoEstadual
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Endereco
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Cidade & "-" & xEmpresa.Estado
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CEP: " & fMascaraCEP(xEmpresa.CEP) & "  FONE: " & fMascaraTelefone(xEmpresa.Telefone)
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    xLinhaDados = lNomeRelatorio
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    Call ImpTermicaImprimeDados("DATA: " & Format(Date, "dd/MM/yyyy") & " AS " & Format(Now, "HH:mm:ss"), True)
    
    xLinhaDados = "Usuário....: 000                                "
    Mid(xLinhaDados, 14, 3) = Format(g_usuario, "000")
    Mid(xLinhaDados, 18, Len(g_nome_usuario)) = g_nome_usuario
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
End Sub
Private Sub ImpTermicaFechaRelatorio(ByVal pNomeRelatorio)
    BioImprime "@Printer.Print  "
    BioImprime "@Printer.Print  "
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|" & pNomeRelatorio & "|@|"
    frm_preview.Show 1
End Sub
Private Sub LoopImprimeEncerrante()
    'Relatório Gerencial
    Dim i As Integer
    Dim i2 As Integer
    Dim xString As String
    Dim xLinha As String
    Dim xValor As Currency
    Dim xQuantidade As Currency
    Dim xValorTotal As Currency
    Dim xUltimoPeriodo As Integer
    
    Call CriaLogSGP(Time & " - Emissão da NFCe Complementar Automação: Início da Impressão dos Encerrantes", "", "")
    xValor = 0
    xQuantidade = 0
    xUltimoPeriodo = 0
    
    lNomeRelatorio = "Relatório de Encerrantes"
    ImpTermicaAbreRelatorio
    Call ImpTermicaCabecalhoPosto
    
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
                'xString = xString & xLinha
                Call ImpTermicaImprimeDados(xLinha, True)
            End If
            rstMovimentoBomba.MoveNext
        Next
    End If
    rstMovimentoBomba.Close
    xLinha = "          TOTAL DO CAIXA:        0,00       0,00"
    i2 = Len(Format(xQuantidade, "###,##0.00"))
    Mid(xLinha, 28 + 10 - i2, i2) = Format(xQuantidade, "###,##0.00")
    i2 = Len(Format(xValor, "###,##0.00"))
    Mid(xLinha, 39 + 10 - i2, i2) = Format(xValor, "###,##0.00")
    'xString = xString & xLinha
    Call ImpTermicaImprimeDados(xLinha, True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call CriaLogSGP(Time & " - Emissão do NFCe Complementar: TOTAL=" & xLinha, "", "")
    Call ImpTermicaFechaRelatorio(lNomeRelatorio)
    Call CriaLogSGP(Time & " - Emissão do Cupom Complementar: Foi Concluído a Impressão dos Encerrantes", "", "")
End Sub
Private Sub GravaDadosNFCe(ByVal pCodigoProduto As Integer, ByVal pTipoCombustivel As String, ByVal pCodigoBico As Integer, ByVal pEncerrante As Currency, ByVal pValorUnitario As Currency, ByVal pQuantidade As Currency, ByVal pTotalItem As Currency, ByVal pTotalNFCE As Currency)
    If lNumeroNFCe = 0 Then
        BuscaNumeroCupom
    Else
        lOrdemNFCe = lOrdemNFCe + 1
    End If
    
    Call AtualizaTabelaCupomFiscal(lNumeroNFCe, lOrdemNFCe, lDataNFCe, lHoraNFCe, pCodigoProduto, pValorUnitario, pQuantidade, pTotalItem, Produto.CodigoAliquota, pTipoCombustivel, Produto.CodigoGrupo, pCodigoBico, pEncerrante)
    Call AtualizaTabelaCupomFiscalItem

    Call AtualizaTabelaDocumentoEletronicoCabecalho(lNumeroNFCe, lDataNFCe, lHoraNFCe, pTotalNFCE)
    Call AtualizaTabelaDocumentoEletronicoItem(lNumeroNFCe, lOrdemNFCe, lDataNFCe, lHoraNFCe, pCodigoProduto, pValorUnitario, pQuantidade, pTotalItem, Produto.CodigoAliquota, pTipoCombustivel, pCodigoBico, pEncerrante)
End Sub
'Private Sub ImpCupomComplementar()
'    Dim x_linha As String
'    Dim xString As String
'    Dim xString2 As String
'    Dim xDescricao As String
'    Dim i As Integer
'
'    Dim CodigoProduto As String
'    Dim NomeProduto As String
'    Dim xAliquota As String
'    Dim Quantidade As String
'    Dim Valor As String
'    Dim ValorDesconto As String
'    Dim ValorAcrescimo As String
'    Dim Departamento As String
'    Dim Un As String
'
'    Dim x_valor_acrescimo As Currency
'    Dim x_valor_desconto As Currency
'    Dim x_total As Currency
'    Dim xValorUnitario As String * 9
'    Dim xValorUnitario2 As Currency
'    Dim xValorTotal As Currency
'    Dim xQuantidade As String * 7
'    Dim xQuantidade2 As Currency
'    Dim xCodigoProduto As Long
'    Dim xCodigoAliquota As Integer
'    Dim xCodigoFiscal As String * 2
'    Dim xStringEmail As String
'
'    Dim xTruncaValor As Double
'    Dim xTruncaQuantidade As Double
'    Dim xTruncaTotalCalculado As Currency
'
'    On Error GoTo ErroImpCupomComplementar
'
'    'Close #3
'    gArquivoTXT.Close
'    'Open "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT" For Input As #3
'    Set gArquivoTXT = gArqTxt.OpenTextFile(lNomeArquivoTXT, ForReading)
'    'Do Until EOF(3)
'    xStringEmail = "Iniciada a impressão em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & "   " & vbCrLf
'    Do Until gArquivoTXT.AtEndOfStream
'        'Line Input #3, x_linha
'        x_linha = gArquivoTXT.ReadLine
'        If Mid(x_linha, 1, 3) = "FIM" Then
'            Exit Do
'        End If
'        Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Acionado a Emissão do C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
'        Call BuscaNumeroCupom
'        x_total = 0
'
'        'Codigo da Aliquota
'        'xCodigoAliquota = Mid(x_linha, 48, 2)
'        'Venda de Item com entrada de departamento,
'        'Verifica se há diferença do total
'        xString = Format(Format(fValidaValor(Mid(x_linha, 63, 15)) * fValidaValor(Mid(x_linha, 50, 13)), "###,##0.0000"), "###,##0.0000")
'        i = Len(xString)
'        xString = Mid(xString, 1, i - 2)
'        x_valor_acrescimo = 0
'        x_valor_desconto = 0
'        If fValidaValor(Mid(x_linha, 78, 13)) > fValidaValor(xString) Then
'            x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - fValidaValor(xString)
'        ElseIf fValidaValor(Mid(x_linha, 78, 13)) < fValidaValor(xString) Then
'            x_valor_desconto = fValidaValor(xString) - fValidaValor(Mid(x_linha, 78, 13))
'        Else
'        End If
'
'
'        'código do produto
'        xCodigoProduto = Mid(x_linha, 1, 5)
'        CodigoProduto = Format(Mid(x_linha, 1, 5), "####0")
'        'nome do produto
'        NomeProduto = Mid(x_linha, 6, 40)
'        xStringEmail = xStringEmail & Mid(NomeProduto, 1, 30)
'        'tipo de tributação
'        xCodigoAliquota = Mid(x_linha, 48, 2)
'        lSQL = "SELECT [Codigo Fiscal] FROM Aliquota WHERE Codigo = " & xCodigoAliquota
'        Set rst = Conectar.RsConexao(lSQL)
'        If Not rst.EOF Then
'            xAliquota = rst![Codigo Fiscal]
'        Else
'            xAliquota = "II"
'        End If
'        If Produto.LocalizarCodigo(CodigoProduto) Then
'            If Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
'                xCodigoAliquota = Aliquota.Codigo
'                xAliquota = Aliquota.CodigoFiscal
'            Else
'                Call CriaLogCupom(Time & " - ERRO Cupom Complementar: Aliquota não encontrada=" & Produto.CodigoAliquota & " -SerieECF=" & lSerieECF)
'            End If
'        Else
'            Call CriaLogCupom(Time & " - ERRO Cupom Complementar: Produto não encontrada=" & CodigoProduto)
'        End If
'        rst.Close
'        'Valor Unitário
'        xString = Format(Mid(x_linha, 63, 15), "000000.000")
'        xStringEmail = xStringEmail & " Vlr:" & xString
'        Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
'        xValorUnitario2 = xString
'        'Quantidade
'        'xString = Format(Mid(x_linha, 50, 13), "0000.000")
'        'Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
'        Quantidade = Mid(x_linha, 55, 4) + Mid(x_linha, 60, 3)
'        xStringEmail = xStringEmail & " Qtd:" & Quantidade
'        xQuantidade2 = Format(Mid(x_linha, 50, 13), "0000.000")
'        'Valor do Acréscimo
'        xString = Format(x_valor_acrescimo, "00000000.00")
'        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'        'Valor do Desconto
'        xString = Format(x_valor_desconto, "00000000.00")
'        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'
'        'Desconsidera Descontos ou Acréscimos
'        If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
'            x_valor_acrescimo = 0
'            x_valor_desconto = 0
'            ValorAcrescimo = "0000000000"
'            ValorDesconto = "0000000000"
'        End If
'
'
'        'Departamento
'        Departamento = Format(1, "00")
'        'Unidade de Medida
'        Un = Mid(x_linha, 46, 2)
'
'        'Tratamento de Truncamento
'        If lEcfTruncamento = True Then
'            xTruncaValor = Format(Mid(x_linha, 63, 15), "000000.0000")
'            If lEcfQtdCasasDecimais = 2 Then
'                xTruncaQuantidade = Format(Mid(x_linha, 50, 12), "0000.000")
'            Else
'                xTruncaQuantidade = Format(Mid(x_linha, 50, 13), "0000.000")
'            End If
'            xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
'            ValorAcrescimo = "0000000000"
'            ValorDesconto = "0000000000"
'            If fValidaValor(Mid(x_linha, 78, 13)) > xTruncaTotalCalculado Then
'                x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - xTruncaTotalCalculado
'                Call CriaLogCupom("Acrescimo Truncamento  valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                xString = Format(x_valor_acrescimo, "00000000.00")
'                ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'            ElseIf fValidaValor(Mid(x_linha, 78, 13)) < xTruncaTotalCalculado Then
'                x_valor_desconto = xTruncaTotalCalculado - fValidaValor(Mid(x_linha, 78, 13))
'                Call CriaLogCupom("Desconto Truncamento   valor total=" & Mid(x_linha, 78, 13) & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                xString = Format(x_valor_desconto, "00000000.00")
'                ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'            End If
'        End If
'
'        'BemaRetorno = Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
'
'        'Grava Cupom Complementar
'        xValorTotal = Format(xValorUnitario2 * xQuantidade2, "0000000000.00") - x_valor_desconto + x_valor_acrescimo
'        Call AtualizaTabelaCupomFiscal(lNumeroNFCe, lOrdemNFCe, lDataNFCe, lHoraNFCe, xCodigoProduto, xValorUnitario2, xQuantidade2, xValorTotal, xCodigoAliquota, x_linha)
'        Call AtualizaTabelaCupomFiscalItem(xCodigoAliquota)
'
'
'        Call AtualizaTabelaDocumentoEletronicoCabecalho(lNumeroNFCe, lDataNFCe, lHoraNFCe, xValorTotal)
'        Call AtualizaTabelaDocumentoEletronicoItem(lNumeroNFCe, lOrdemNFCe, lDataNFCe, lHoraNFCe, xCodigoProduto, xValorUnitario2, xQuantidade2, xValorTotal, xCodigoAliquota, x_linha)
'
'        'Desconto para o Cupom Fiscal
'        xString = Mid(Format(fValidaValor(0), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(0), "000000000000.00"), 14, 2)
'        'BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
'
'        'Efetua Forma de Pagamento
'        xString = "Dinheiro        "
'        xString2 = Mid(Format(xValorTotal, "000000000000.00"), 1, 12) + Mid(Format(xValorTotal, "000000000000.00"), 14, 2)
'        xDescricao = ""
'        'BemaRetorno = Bematech_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
'        'Fecha Cupom Fiscal
'        xString = "Cerrado Tecnologia - (62) 3277-1017             Soluções Inteligentes                           "
'        'BemaRetorno = Bematech_FI_TerminaFechamentoCupom(xString)
'
'
'        Call EnviaDadosParaNFCe(lNumeroNFCe, lDataNFCe)
'        MsgBox "Aguarde o Retorno e Tecle Enter", vbInformation, "Aguarde"
'
'
'        Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Impresso o C.F. de " & Format(Mid(x_linha, 51, 13), "0000.000") & " " & Mid(x_linha, 46, 3) & " de " & Mid(x_linha, 6, 40))
'        xStringEmail = xStringEmail & vbCrLf
'    Loop
'
'    'Testa se tem Automação
'    'Quando for necessário enviar email
'    'Retirar o comentário logo abaixo
'    If Mid(Configuracao.OutrasConfiguracoes, 5, 1) = "S" Then
'        xStringEmail = xStringEmail & "Finalizado a impressão em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS")
'        'Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Cupom Complementar!", xStringEmail, True, gNumeroEmailInicial)
'    End If
'
'    Exit Sub
'ErroImpCupomComplementar:
'    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Erro ImpCupomComplementar - " & x_linha)
'    Exit Sub
'End Sub
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
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Pedido a Impressão do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    'DefineImpressoraTermicaComoPadrao
    
    lLocal = 1
    If ValidaCampos Then
        Call GravaAuditoria(1, Me.name, 7, "Data Inicial:" & msk_data_i.Text & " Data Final:" & msk_data_f.Text)
        Call AtivaBotoes(False)
        
        If lDataHoraInicioProcessamento <> CDate("00:00:00") Then
           lProcessamentoNFCeComplementar.DataHoraEmissao = Now
           Call lProcessamentoNFCeComplementar.DefinirDataHoraEmissao(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV)
        End If
        
        g_string = "imprimiu|@|"
        Relatorio
        Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Impresso")
        If TestaImprimeEncerrante Then
            If (MsgBox("Deseja imprimir encerrante de bombas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Relatório Gerencial na Impressora Térmica.") = vbYes) Then
                LoopImprimeEncerrante
            End If
        End If
    End If
    Exit Sub
ErroImprimir:
    MsgBox (Err.Description & " - " & Err.Number)
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Erro ao abrir o arquivo CUPOM_COMPLEMENTAR.TXT")
    Call AtivaBotoes(True)
    Exit Sub
End Sub

Private Sub GravaRegistroProcessamento()

    Set lProcessamentoNFCeComplementar = New cProcessaNFCeComplementar
    
    With lProcessamentoNFCeComplementar
    
        .IdEstabelecimento = g_empresa
        .PDV = lNumeroPDV
        .DataHoraInicio = lDataHoraInicioProcessamento
        .DataHoraEmissao = CDate("00:00:00")
        .DataHoraTermino = CDate("00:00:00")
        .Observacao = ""
    
    End With
    
    Call lProcessamentoNFCeComplementar.Incluir

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
Private Sub TotalizaAcertoVendaNFCe()
    Dim i As Integer
    
    For i = 1 To 7
        If AcertoVendaECF.LocalizarCodigo(g_empresa, CDate(msk_data_i.Text), lTipoCombustivel(i)) Then
            Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: AcertoVendaECF:" & msk_data_i.Text & " Comb:" & lTipoCombustivel(i) & " " & AcertoVendaECF.Operacao & " Qtd:" & AcertoVendaECF.Quantidade)
            If AcertoVendaECF.Operacao = "+" Then
                lExisteProgramacaoASomar = True
                lQtdBombaV(i) = lQtdBombaV(i) + AcertoVendaECF.Quantidade
                lTotalBombaV(i) = lTotalBombaV(i) + AcertoVendaECF.ValorTotal
            ElseIf AcertoVendaECF.Operacao = "-" Then
                lExisteProgramacaoADiminuir = True
                lQtdBombaV(i) = lQtdBombaV(i) - AcertoVendaECF.Quantidade
                lTotalBombaV(i) = lTotalBombaV(i) - AcertoVendaECF.ValorTotal
            End If
        End If
    
        lSQL = ""
        lSQL = lSQL & "SELECT Sum(Quantidade) As Quantidade, Sum([Valor Total]) As ValorTotal"
        lSQL = lSQL & "  FROM Movimento_Abastecimento"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & "   AND Acerto = " & preparaBooleano(True)
        lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel(i))
        lSQL = lSQL & "   AND [Documento Gerado] = " & preparaTexto("AVP")
        Set rstMovimentoAbastecimento = Conectar.RsConexao(lSQL)
        If rstMovimentoAbastecimento.RecordCount > 0 Then
            rstMovimentoAbastecimento.MoveFirst
            If Not IsNull(rstMovimentoAbastecimento!Quantidade) Then
                lQtdBombaV(i) = lQtdBombaV(i) + rstMovimentoAbastecimento!Quantidade
                lTotalBombaV(i) = lTotalBombaV(i) + rstMovimentoAbastecimento!ValorTotal
            End If
        End If
        rstMovimentoAbastecimento.Close
    Next
End Sub
Private Sub TotalizaMovimentoAfericao()
    Dim i As Integer
    
    For i = 1 To 7
        lQtdAfericaoV(i) = lQtdAfericaoV(i) + MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 1, 9, lTipoCombustivel(i), "")
        lTotalAfericaoV(i) = lTotalAfericaoV(i) + MovimentoAfericao.ValorTotalPeriodoCombustivel(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 1, 9, lTipoCombustivel(i), "")
    Next
    
    'subtrai as aferições nas vendas de bomba
    For i = 1 To 7
        If lQtdAfericaoV(i) > 0 Then
            lQtdBombaV(i) = lQtdBombaV(i) - lQtdAfericaoV(i)
            lTotalBombaV(i) = lTotalBombaV(i) - lTotalAfericaoV(i)
            Call CriaLogCupom(Time & " - [Emissão de NFCe Complementar Automação - AFERICAO ] À Vista - Combustivel:" & lTipoCombustivel(i) & " - Lts.Aferição:" & lQtdAfericaoV(i))
        End If
    Next
End Sub
'Private Sub TotalizaMovimentoAbastecimento()
'    Dim i As Integer
'    With rstMovimentoAbastecimento
'        .MoveFirst
'        Do Until .EOF
'            For i = 1 To 7
'                If lTipoCombustivel(i) = ![Tipo de Combustivel] Then
'                    lQtdBombaV(i) = lQtdBombaV(i) + !Quantidade
'                    lTotalBombaV(i) = lTotalBombaV(i) + !ValorTotal
'                    Exit For
'                End If
'            Next
'            .MoveNext
'        Loop
'    End With
'End Sub
'Private Sub TotalizaMovimentoBombaCupom()
'    Dim i As Integer
'    With rstMovimentoBomba
'        .MoveFirst
'        Do Until .EOF
'            For i = 1 To 7
'                If lTipoCombustivel(i) = ![Tipo de Combustivel] Then
'                    lQtdBombaV(i) = lQtdBombaV(i) + ![Quantidade da Saida]
'                    lTotalBombaV(i) = lTotalBombaV(i) + (![Quantidade da Saida] * ![Preco de Venda])
'                    Exit For
'                End If
'            Next
'            .MoveNext
'        Loop
'    End With
'End Sub
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
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Pedido a Visualização do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    lLocal = 0
   ' DefineImpressoraTermicaComoPadrao
    If ValidaCampos Then
        Call GravaAuditoria(1, Me.name, 6, "Data Inicial:" & msk_data_i.Text & " Data Final:" & msk_data_f.Text)
        Call AtivaBotoes(False)
        Relatorio
        Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Visualizado")
        If TestaImprimeEncerrante Then
            If (MsgBox("Deseja imprimir encerrante de bombas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Relatório Gerencial na Impressora Térmica.") = vbYes) Then
                LoopImprimeEncerrante
            End If
        End If
    End If
End Sub

Private Sub cmdImprimiEncerranteAtual_Click()
'    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: Foi Pedido a Visualização do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
'    lLocal = 0
'    If ValidaCampos Then
'        ImprimeEncerranteAutomacaoAtual ("V")
'    End If
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
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação (5): A Emissão de NFCe Complementar Automação Foi Aberta")
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
    Call CriaLogCupom(Time & " - Emissão de NFCe Complementar Automação: A Emissão de NFCe Complementar Automação Foi Aberta")
    CentraForm Me
    lDataHoraInicioProcessamento = CDate("00:00:00")
    Call lProcessamentoNFCeComplementar.DefinirDataHoraTerminoPendentes(g_empresa, Now, "FINALIZADO NO PROCESSAMENTO DE PENDENTES")
    
    MovimentoAfericao.NomeTabela = "Movimento_Afericao"
    AtualizaConstantes
    
    lTerminoDefinido = False
    lDataHoraInicioProcessamento = Now
    Call GravaRegistroProcessamento

End Sub
Private Sub Form_Unload(Cancel As Integer)
    If lDataHoraInicioProcessamento <> CDate("00:00:00") And lTerminoDefinido = False Then
        lProcessamentoNFCeComplementar.DataHoraTermino = Now
        Call lProcessamentoNFCeComplementar.DefinirDataHoraTermino(g_empresa, lDataHoraInicioProcessamento, lNumeroPDV, "FINALIZADO SEM IMPRIMIR")
    End If
    
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
