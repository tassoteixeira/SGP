VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form movimento_mapa_resumo 
   Caption         =   "Movimentação do Mapa Resumo E.C.F."
   ClientHeight    =   6795
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   8760
   Icon            =   "movimento_mapa_resumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   8760
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_mapa_resumo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Imprime o Mapa Resumo."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_mapa_resumo.frx":1914
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Cria um novo registro."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_mapa_resumo.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Altera o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_mapa_resumo.frx":44A0
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_mapa_resumo.frx":5B32
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4620
      Picture         =   "movimento_mapa_resumo.frx":6FA4
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5880
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Caption         =   "Movimentação do Mapa Resumo E.C.F."
      Enabled         =   0   'False
      Height          =   5715
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8595
      Begin VB.TextBox txtICMS3 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   32
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txtIcms19 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   42
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txtIcms13 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   38
         Top             =   3840
         Width           =   1275
      End
      Begin VB.TextBox txtIcms25 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   44
         Top             =   4560
         Width           =   1275
      End
      Begin VB.TextBox txtIcms7 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   34
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txtNaoIncidencia 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   30
         Top             =   3120
         Width           =   1275
      End
      Begin VB.TextBox txtAcrescimoIcms 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   24
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtDescontoIcms 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   20
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtContagemReinicioOperacao 
         Height          =   285
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1320
         Width           =   555
      End
      Begin VB.TextBox txtIcms12 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   36
         Top             =   3840
         Width           =   1275
      End
      Begin VB.TextBox txt_Contador_Reducoes_Z 
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   46
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_numero_reducao 
         Height          =   285
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_observacao_2 
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   49
         Top             =   5280
         Width           =   5415
      End
      Begin VB.TextBox txt_observacao_1 
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   48
         Top             =   4920
         Width           =   5415
      End
      Begin VB.TextBox txt_icms_17 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   40
         Top             =   4200
         Width           =   1275
      End
      Begin VB.TextBox txt_substituicao 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   26
         Top             =   2760
         Width           =   1275
      End
      Begin VB.TextBox txt_isentas 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   28
         Top             =   2760
         Width           =   1275
      End
      Begin VB.TextBox txt_valor_contabil 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   22
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txt_cancelamento_item 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   18
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txt_gt_inicial 
         Height          =   285
         Left            =   2880
         MaxLength       =   14
         TabIndex        =   14
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txt_gt_final 
         Height          =   285
         Left            =   7200
         MaxLength       =   14
         TabIndex        =   16
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txt_contador_final 
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_contador_inicial 
         Height          =   285
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_ecf_numero 
         Height          =   285
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   6
         Top             =   600
         Width           =   555
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "IC&MS 3%"
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   2715
      End
      Begin VB.Label Label13 
         Caption         =   "IC&MS 19%"
         Height          =   285
         Index           =   2
         Left            =   4440
         TabIndex        =   41
         Top             =   4200
         Width           =   2715
      End
      Begin VB.Label Label17 
         Caption         =   "IC&MS 13%"
         Height          =   285
         Left            =   4440
         TabIndex        =   37
         Top             =   3840
         Width           =   2715
      End
      Begin VB.Label Label13 
         Caption         =   "IC&MS 25%"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   4560
         Width           =   2715
      End
      Begin VB.Label Label16 
         Caption         =   "IC&MS 17%"
         Height          =   165
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Width           =   2715
      End
      Begin VB.Label Label15 
         Caption         =   "Não Incidência"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   2715
      End
      Begin VB.Label Label9 
         Caption         =   "Acréscimo de ICMS"
         Height          =   285
         Left            =   4440
         TabIndex        =   23
         Top             =   2400
         Width           =   2715
      End
      Begin VB.Label Label7 
         Caption         =   "Desconto de ICMS"
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Contagem de Reinício de Operação"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label lblIcms12 
         Caption         =   "IC&MS 7%"
         Height          =   285
         Left            =   4440
         TabIndex        =   33
         Top             =   3480
         Width           =   2715
      End
      Begin VB.Label Label12 
         Caption         =   "Contador de &Reduções Z"
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   45
         Top             =   4560
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "&Número da Redução Z"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2715
      End
      Begin VB.Label Label14 
         Caption         =   "&Observações"
         Height          =   285
         Index           =   24
         Left            =   120
         TabIndex        =   47
         Top             =   4920
         Width           =   2715
      End
      Begin VB.Label Label13 
         Caption         =   "IC&MS 12%"
         Height          =   285
         Index           =   22
         Left            =   120
         TabIndex        =   35
         Top             =   3840
         Width           =   2715
      End
      Begin VB.Label Label12 
         Caption         =   "&Substituição Tributária"
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   2715
      End
      Begin VB.Label Label11 
         Caption         =   "Isent&as"
         Height          =   285
         Index           =   18
         Left            =   4440
         TabIndex        =   27
         Top             =   2760
         Width           =   2715
      End
      Begin VB.Label Label10 
         Caption         =   "&Valor Contábil"
         Height          =   285
         Index           =   16
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   2715
      End
      Begin VB.Label lblCancelamentoItem 
         Caption         =   "&Cancelamento de ICMS"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label8 
         Caption         =   "Totalizador &Geral (GT) Inicial"
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2715
      End
      Begin VB.Label Label6 
         Caption         =   "&Totalizador Geral (GT) Final"
         Height          =   285
         Index           =   10
         Left            =   4440
         TabIndex        =   15
         Top             =   1680
         Width           =   2715
      End
      Begin VB.Label Label3 
         Caption         =   "Cont. de Ordem de Operação &Final"
         Height          =   285
         Index           =   8
         Left            =   4440
         TabIndex        =   9
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "Cont. de Ordem de Operação &Inicial"
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label Label5 
         Caption         =   "&Data"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "&ECF Número"
         Height          =   285
         Index           =   4
         Left            =   4440
         TabIndex        =   5
         Top             =   600
         Width           =   2715
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6540
      TabIndex        =   58
      Top             =   5760
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_mapa_resumo.frx":8636
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_mapa_resumo.frx":9B30
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_mapa_resumo.frx":B02A
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_mapa_resumo.frx":C49C
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "movimento_mapa_resumo.frx":DA1E
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7920
      Picture         =   "movimento_mapa_resumo.frx":F028
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   6060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "movimento_mapa_resumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lNumero As Long
Dim lColuna(0 To 1) As Currency
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lLinha As Currency
Dim lLinhaDepoisTotal As Currency
Dim lLinhaTab As Currency
Dim lLocal As Integer
Dim lEcfNumero As Integer
Dim lCodigoEcfImprimir As Integer
Dim lSairDiferenca As Boolean
Dim lImprimeTodasReducoesDia As Boolean
Dim lQtdReducao As Integer
Dim lCancelamentoItem As Currency
Dim lAcrescimo As Currency
Dim lValorContabil As Currency
Dim lIsentas As Currency
Dim lNaoIncidencia As Currency
Dim lSubstituicaoTributaria As Currency
Dim lICMS12 As Currency
Dim lICMS17 As Currency
Dim lImpostoDebitado As Currency

Private Empresa As New cEmpresa
Private MovMapaResumo As New cMovimentoMapaResumo
Private MovCupomFiscal As New cMovimentoCupomFiscal
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem

Private Sub AtualTabe()
    MovMapaResumo.Empresa = g_empresa
    MovMapaResumo.Data = msk_data.Text
    MovMapaResumo.numero = CLng(txt_numero_reducao.Text)
    MovMapaResumo.ECFNumero = txt_ecf_numero.Text
    MovMapaResumo.ContagemOperacaoInicial = CLng(txt_contador_inicial.Text)
    MovMapaResumo.ContagemOperacaoFinal = CLng(txt_contador_final.Text)
    MovMapaResumo.ContagemReinicioOperacao = Val(txtContagemReinicioOperacao.Text)
    MovMapaResumo.TotalizadorGeralFinal = fValidaValor(txt_gt_final.Text)
    MovMapaResumo.TotalizadorGeralInicial = fValidaValor(txt_gt_inicial.Text)
    MovMapaResumo.CancelamentoItem = fValidaValor(txt_cancelamento_item.Text)
    MovMapaResumo.Desconto = fValidaValor(txtDescontoIcms.Text)
    MovMapaResumo.Acrescimo = fValidaValor(txtAcrescimoIcms.Text)
    MovMapaResumo.ValorContabil = fValidaValor(txt_valor_contabil.Text)
    MovMapaResumo.SubstituicaoTributaria = fValidaValor(txt_substituicao.Text)
    MovMapaResumo.Isentas = fValidaValor(txt_isentas.Text)
    MovMapaResumo.NaoIncidencia = fValidaValor(txtNaoIncidencia.Text)
    MovMapaResumo.ICMS3 = fValidaValor(txtICMS3.Text) 'insere o valor do txtIcms7 no objeto
    MovMapaResumo.ICMS7 = fValidaValor(txtIcms7.Text) 'insere o valor do txtIcms7 no objeto
    MovMapaResumo.ICMS12 = fValidaValor(txtIcms12.Text)
    MovMapaResumo.ICMS17 = fValidaValor(txt_icms_17.Text)
    MovMapaResumo.ICMS13 = fValidaValor(txtIcms13.Text)
    MovMapaResumo.ICMS25 = fValidaValor(txtIcms25.Text) 'insere o valor do txtIcms25 no objeto
    MovMapaResumo.ICMS19 = fValidaValor(txtIcms19.Text)
    MovMapaResumo.ContadorReducoesZ = CLng(txt_Contador_Reducoes_Z.Text)
    MovMapaResumo.Observacao1 = txt_observacao_1.Text
    MovMapaResumo.Observacao2 = txt_observacao_2.Text
End Sub
Private Sub AtualTela()
    lData = MovMapaResumo.Data
    lNumero = MovMapaResumo.numero
    lEcfNumero = MovMapaResumo.ECFNumero
    
    msk_data.Text = Format(MovMapaResumo.Data, "dd/mm/yyyy")
    txt_numero_reducao.Text = Format(MovMapaResumo.numero, "##,###,##0")
    txt_ecf_numero.Text = Format(MovMapaResumo.ECFNumero, "#0")
    txt_contador_inicial.Text = Format(MovMapaResumo.ContagemOperacaoInicial, "##,###,##0")
    txt_contador_final.Text = Format(MovMapaResumo.ContagemOperacaoFinal, "##,###,##0")
    txtContagemReinicioOperacao.Text = Format(MovMapaResumo.ContagemReinicioOperacao, "##0")
    txt_gt_final.Text = Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00")
    txt_gt_inicial.Text = Format(MovMapaResumo.TotalizadorGeralInicial, "###,###,##0.00")
    txt_cancelamento_item.Text = Format(MovMapaResumo.CancelamentoItem, "###,###,##0.00")
    txtDescontoIcms.Text = Format(MovMapaResumo.Desconto, "###,###,##0.00")
    txtAcrescimoIcms.Text = Format(MovMapaResumo.Acrescimo, "###,###,##0.00")
    txt_valor_contabil.Text = Format(MovMapaResumo.ValorContabil, "###,###,##0.00")
    txt_substituicao.Text = Format(MovMapaResumo.SubstituicaoTributaria, "###,###,##0.00")
    txt_isentas.Text = Format(MovMapaResumo.Isentas, "###,###,##0.00")
    txtNaoIncidencia.Text = Format(MovMapaResumo.NaoIncidencia, "###,###,##0.00")
    txtICMS3.Text = Format(MovMapaResumo.ICMS3, "###,###,##0.00")
    txtIcms7.Text = Format(MovMapaResumo.ICMS7, "###,###,##0.00") 'o campo txtIcms7 recebe o valor vindo do MovMapaResumo.ICMS7 ja formatado
    txtIcms12.Text = Format(MovMapaResumo.ICMS12, "###,###,##0.00")
    txtIcms13.Text = Format(MovMapaResumo.ICMS13, "###,###,##0.00")
    txt_icms_17.Text = Format(MovMapaResumo.ICMS17, "###,###,##0.00")
    txtIcms19.Text = Format(MovMapaResumo.ICMS19, "###,###,##0.00")
    txtIcms25.Text = Format(MovMapaResumo.ICMS25, "###,###,##0.00") 'o campo txtIcms25 recebe o valor vindo do MovMapaResumo.ICMS25 ja formatado
    txt_Contador_Reducoes_Z = Format(MovMapaResumo.ContadorReducoesZ, "##,###,##0")
    txt_observacao_1.Text = MovMapaResumo.Observacao1
    txt_observacao_2.Text = MovMapaResumo.Observacao2
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set MovMapaResumo = Nothing
    Set MovCupomFiscal = Nothing
    Set MovCupomFiscalItem = Nothing
End Sub
Private Sub ImpDados()
    Dim xValor As Currency
    Dim xValorCombustivel As Currency
    Dim xValorLubrificante As Currency
    Dim xValorGas As Currency
    Dim xValorCancelamento As Currency
    Dim xStringLog As String
    Dim i As Integer
        
    Call Empresa.LocalizarCodigo(g_empresa)
    
    Printer.FontSize = 12
    Printer.FontBold = False
    If Empresa.Nome Like "*POSTO DO RATINHO*" Then
    Else
        ImprimeCentralizado Format(MovMapaResumo.numero, "###,###,##0"), lColunaI + 17, lColunaI + 22, lLinhaI + 0.4, lLocal
    End If
    ImprimeCentralizado msk_data.Text, lColunaI + 22, lColunaI + 26, lLinhaI + 0.4, lLocal
    
    ImprimeTexto Empresa.Nome, lColunaI + 0.2, lColunaI + 20, lLinhaI + 1.4, lLocal
    ImprimeTexto Empresa.InscricaoEstadual, lColunaI + 20.2, lColunaI + 26, lLinhaI + 1.4, lLocal
    
    ImprimeTexto Trim(Empresa.Endereco) & ", " & Trim(Empresa.Bairro), lColunaI + 0.2, lColunaI + 11, lLinhaI + 2.4, lLocal
    ImprimeTexto Empresa.Cidade, lColunaI + 11.2, lColunaI + 19, lLinhaI + 2.4, lLocal
    ImprimeCentralizado Empresa.Estado, lColunaI + 19, lColunaI + 20.1, lLinhaI + 2.4, lLocal
    ImprimeTexto fMascaraCNPJ(Empresa.CGC), lColunaI + 20.2, lColunaI + 26, lLinhaI + 2.4, lLocal
    
    Printer.FontSize = 9
    Printer.FontBold = False
    lLinha = 3.95
    For i = 1 To lQtdReducao
        lLinha = lLinha + 0.5
        ImprimeCentralizado Format(MovMapaResumo.ECFNumero, "00"), lColunaI + 0, lColunaI + 1, lLinhaI + lLinha, lLocal
        ImprimeCentralizado Format(MovMapaResumo.ContagemOperacaoInicial, "###,###,##0"), lColunaI + 1, lColunaI + 2.4, lLinhaI + lLinha, lLocal
        ImprimeCentralizado Format(MovMapaResumo.ContagemOperacaoFinal, "###,###,##0"), lColunaI + 2.4, lColunaI + 3.9, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00"), lColunaI + 3.9, lColunaI + 6.6, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.TotalizadorGeralInicial, "###,###,##0.00"), lColunaI + 6.6, lColunaI + 9.3, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.CancelamentoItem + MovMapaResumo.Desconto - MovMapaResumo.Acrescimo, "###,###,##0.00"), lColunaI + 9.3, lColunaI + 11.3, lLinhaI + lLinha, lLocal
        If UCase(g_nome_empresa) Like "*VENTANIA*" Then
            xValor = Format(MovMapaResumo.TotalizadorGeralFinal - MovMapaResumo.TotalizadorGeralInicial, "0000000000.00")
            ImprimeValor Format(xValor, "###,###,##0.00"), lColunaI + 11.3, lColunaI + 13.5, lLinhaI + lLinha, lLocal
        Else
            ImprimeValor Format(MovMapaResumo.ValorContabil, "###,###,##0.00"), lColunaI + 11.3, lColunaI + 13.5, lLinhaI + lLinha, lLocal
        End If
        ImprimeValor Format(MovMapaResumo.SubstituicaoTributaria, "###,###,##0.00"), lColunaI + 13.5, lColunaI + 15.3, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.Isentas, "###,###,##0.00"), lColunaI + 15.3, lColunaI + 17.18, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.NaoIncidencia, "###,###,##0.00"), lColunaI + 17.18, lColunaI + 19, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.ICMS12, "###,###,##0.00"), lColunaI + 19, lColunaI + 20.82, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.ICMS17, "###,###,##0.00"), lColunaI + 20.82, lColunaI + 22.6, lLinhaI + lLinha, lLocal
        xValor = 0
        If MovMapaResumo.ICMS12 > 0 Or MovMapaResumo.ICMS17 > 0 Then
            'Quando for ventania, truncar valor o imposto
            If UCase(g_nome_empresa) Like "*VENTANIA*" Then
                xValor = CCur(Mid(Format((MovMapaResumo.ICMS12 * 12 / 100), "0000000000.0000"), 1, 13)) + CCur(Mid(Format((MovMapaResumo.ICMS17 * 17 / 100), "0000000000.0000"), 1, 13))
            Else
                xValor = (MovMapaResumo.ICMS12 * 12 / 100) + (MovMapaResumo.ICMS17 * 17 / 100)
            End If
        End If
        ImprimeValor Format(xValor, "###,###,##0.00"), lColunaI + 22.65, lColunaI + 24.5, lLinhaI + lLinha, lLocal
        ImprimeValor Format(MovMapaResumo.ContadorReducoesZ, "###,###,##0"), lColunaI + 24.5, lColunaI + 26, lLinhaI + lLinha, lLocal
        
        lCancelamentoItem = lCancelamentoItem + MovMapaResumo.CancelamentoItem + MovMapaResumo.Desconto
        lAcrescimo = lAcrescimo + MovMapaResumo.Acrescimo
        lValorContabil = lValorContabil + MovMapaResumo.ValorContabil
        lIsentas = lIsentas + MovMapaResumo.Isentas
        lNaoIncidencia = lNaoIncidencia + MovMapaResumo.NaoIncidencia
        lSubstituicaoTributaria = lSubstituicaoTributaria + MovMapaResumo.SubstituicaoTributaria
        lICMS12 = lICMS12 + MovMapaResumo.ICMS12
        lICMS17 = lICMS17 + MovMapaResumo.ICMS17
        lImpostoDebitado = lImpostoDebitado + xValor
        
        If i < lQtdReducao Then
            MovMapaResumo.LocalizarProximo
        End If
    Next
    
    'ImprimeValor Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00"), lColunaI + 3.9, lColunaI + 6.6, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    'ImprimeValor Format(MovMapaResumo.TotalizadorGeralInicial, "###,###,##0.00"), lColunaI + 6.6, lColunaI + 9.3, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lCancelamentoItem - lAcrescimo, "###,###,##0.00"), lColunaI + 9.3, lColunaI + 11.3, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    If UCase(g_nome_empresa) Like "*VENTANIA*" Then
        xValor = Format(MovMapaResumo.TotalizadorGeralFinal - MovMapaResumo.TotalizadorGeralInicial, "0000000000.00")
        ImprimeValor Format(xValor, "###,###,##0.00"), lColunaI + 11.3, lColunaI + 13.5, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    Else
        ImprimeValor Format(lValorContabil, "###,###,##0.00"), lColunaI + 11.3, lColunaI + 13.5, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    End If
    ImprimeValor Format(lSubstituicaoTributaria, "###,###,##0.00"), lColunaI + 13.5, lColunaI + 15.3, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lIsentas, "###,###,##0.00"), lColunaI + 15.3, lColunaI + 17.18, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lNaoIncidencia, "###,###,##0.00"), lColunaI + 17.18, lColunaI + 19, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lICMS12, "###,###,##0.00"), lColunaI + 19, lColunaI + 20.82, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lICMS17, "###,###,##0.00"), lColunaI + 20.82, lColunaI + 22.6, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    ImprimeValor Format(lImpostoDebitado, "###,###,##0.00"), lColunaI + 22.65, lColunaI + 24.5, lLinhaI + lLinhaDepoisTotal - 0.45, lLocal
    
    ImprimeTexto MovMapaResumo.Observacao1, lColunaI + 0, lColunaI + 13, lLinhaI + lLinhaDepoisTotal + 0.7, lLocal
    ImprimeTexto MovMapaResumo.Observacao2, lColunaI + 0, lColunaI + 13, lLinhaI + lLinhaDepoisTotal + 1.2, lLocal
    ImprimeCentralizado Trim(Empresa.ResponsavelLegal), lColunaI + 13, lColunaI + 26, lLinhaI + lLinhaDepoisTotal + 1.2, lLocal

    xValorCombustivel = MovCupomFiscal.ValorCombustiveisVendaData(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, lCodigoEcfImprimir)
    xValorGas = MovCupomFiscal.ValorProdutoGasVendaData(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, lCodigoEcfImprimir)
    xValorCombustivel = xValorCombustivel + xValorGas
    xValorLubrificante = MovCupomFiscalItem.TotValorProdutosSubstVendaData(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, lCodigoEcfImprimir)
    'soma Venda de Produtos Tributado 17
    xValorLubrificante = xValorLubrificante - xValorGas + MovCupomFiscalItem.TotValorProdutosTribVendaData(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, lCodigoEcfImprimir, 17)
    'soma Venda de Produtos Tributado 12
    xValorLubrificante = xValorLubrificante + MovCupomFiscalItem.TotValorProdutosTribVendaData(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, lCodigoEcfImprimir, 12)
    'MovMapaResumo.ICMS17
    
    If UCase(g_nome_empresa) Like "*VENTANIA*" Then
        ImprimeTexto "Combustível + Gás......: ", lColunaI + 0, lColunaI + 3.5, lLinhaI + lLinhaDepoisTotal + 1.7, lLocal
        ImprimeValor Format(xValorCombustivel, "###,###,#0.00"), lColunaI + 3.5, lColunaI + 5.5, lLinhaI + lLinhaDepoisTotal + 1.7, lLocal
        ImprimeTexto "Produtos/Lubrificantes.: ", lColunaI + 0, lColunaI + 3.5, lLinhaI + lLinhaDepoisTotal + 2.2, lLocal
        ImprimeValor Format(xValorLubrificante, "###,###,#0.00"), lColunaI + 3.5, lColunaI + 5.5, lLinhaI + lLinhaDepoisTotal + 2.2, lLocal
    End If
    xStringLog = "Comb:" & Format(xValorCombustivel, "###,###,#0.00")
    xStringLog = xStringLog & " Prod:" & Format(xValorLubrificante, "###,###,#0.00")
    If lSairDiferenca Then
        ImprimeTexto "Diferença..............: ", lColunaI + 0, lColunaI + 3.5, lLinhaI + lLinhaDepoisTotal + 2.7, lLocal
        ImprimeValor Format(lValorContabil - (xValorCombustivel + xValorLubrificante), "###,###,#0.00"), lColunaI + 3.5, lColunaI + 5.5, lLinhaI + lLinhaDepoisTotal + 2.7, lLocal
        xStringLog = xStringLog & " Dif:" & Format(lValorContabil - (xValorCombustivel + xValorLubrificante), "###,###,#0.00")
    End If
    
    'Descontos e Cancelamentos
    xValorCancelamento = MovCupomFiscalItem.TotalCancelamento(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), 1, 9, 0, lCodigoEcfImprimir)
    xValor = lCancelamentoItem - xValorCancelamento
    If UCase(g_nome_empresa) Like "*VENTANIA*" Then
        ImprimeTexto "Desconto..........: ", lColunaI + 7.5, lColunaI + 10, lLinhaI + lLinhaDepoisTotal + 1.7, lLocal
        ImprimeValor Format(xValor, "###,###,#0.00"), lColunaI + 10, lColunaI + 12, lLinhaI + lLinhaDepoisTotal + 1.7, lLocal
        ImprimeTexto "Cancelamento..: ", lColunaI + 7.5, lColunaI + 10, lLinhaI + lLinhaDepoisTotal + 2.2, lLocal
        ImprimeValor Format(xValorCancelamento, "###,###,#0.00"), lColunaI + 10, lColunaI + 12, lLinhaI + lLinhaDepoisTotal + 2.2, lLocal
        ImprimeTexto "Acrescimo........: ", lColunaI + 7.5, lColunaI + 10, lLinhaI + lLinhaDepoisTotal + 2.7, lLocal
        ImprimeValor Format(lAcrescimo, "###,###,#0.00"), lColunaI + 10, lColunaI + 12, lLinhaI + lLinhaDepoisTotal + 2.7, lLocal
    End If
    xStringLog = xStringLog & " Desc:" & Format(xValor, "###,###,#0.00")
    xStringLog = xStringLog & " Canc:" & Format(xValorCancelamento, "###,###,#0.00")
    Call GravaAuditoria(1, Me.name, 7, xStringLog)
    
    Printer.EndDoc
End Sub
Private Sub ImpGrade()
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    Dim i As Integer
    Dim xLinhaFinal As Currency
    
    'seleciona medidas para centímetros
    Printer.ColorMode = 2
    Printer.ScaleMode = 7
    Printer.PrintQuality = vbPRPQLow
    
    On Error Resume Next
    Printer.PaperSize = vbPRPSLegal
    On Error GoTo ErrorRotina
    Printer.Orientation = 2
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    'teste para imprimir letra correta
    Printer.FontBold = False
    ImprimeTexto "  ", lColuna(0), lColuna(1), lLinhaTab, lLocal
    Printer.FontBold = True
    xR = 0
    xG = 0
    xB = 0
    Printer.DrawWidth = 1
    Printer.ForeColor = RGB(xR, xG, xB)
    
    'Bordas Externas
    xLinhaFinal = 8 + (lQtdReducao * 0.5)
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 26, lLinhaI + 0), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 0, lLinhaI + xLinhaFinal), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 26, lLinhaI + 0)-(lColunaI + 26, lLinhaI + xLinhaFinal), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + xLinhaFinal)-(lColunaI + 26, lLinhaI + xLinhaFinal), RGB(xR, xG, xB)
    
    'Linhas Horizontais do Cabecalho
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeCentralizado "MAPA RESUMO E.C.F. (EQUIPAMENTO EMISSOR DE CUPOM FISCAL) - MR", lColunaI + 0, lColunaI + 17, lLinhaI + 0.25, lLocal
    Printer.FontBold = False
    Printer.FontSize = 8
    ImprimeTexto "NÚMERO", lColunaI + 17 + 0.2, lColunaI + 22, lLinhaI + 0.1, lLocal
    ImprimeTexto "DATA", lColunaI + 22 + 0.2, lColunaI + 26, lLinhaI + 0.1, lLocal
    ImprimeTexto "NOME", lColunaI + 0.2, lColunaI + 20, lLinhaI + 1.1, lLocal
    ImprimeTexto "INSCRIÇÃO ESTADUAL", lColunaI + 20 + 0.2, lColunaI + 26, lLinhaI + 1.1, lLocal
    ImprimeTexto "ENDEREÇO", lColunaI + 0.2, lColunaI + 11, lLinhaI + 2.1, lLocal
    ImprimeTexto "MUNICÍPIO", lColunaI + 11 + 0.2, lColunaI + 19, lLinhaI + 2.1, lLocal
    ImprimeTexto "UF", lColunaI + 19 + 0.2, lColunaI + 20.1, lLinhaI + 2.1, lLocal
    ImprimeTexto "CNPJ", lColunaI + 20.1 + 0.2, lColunaI + 26, lLinhaI + 2.1, lLocal
    Printer.DrawWidth = 1
    Printer.Line (lColunaI + 0, lLinhaI + 1)-(lColunaI + 26, lLinhaI + 1), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 2)-(lColunaI + 26, lLinhaI + 2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 3)-(lColunaI + 26, lLinhaI + 3), RGB(xR, xG, xB)
    
    'Linhas Verticais do Cabecalho
    ImprimeCentralizado "ECF", lColunaI, lColunaI + 1, lLinhaI + 3.4, lLocal
    ImprimeCentralizado "Nº", lColunaI, lColunaI + 1, lLinhaI + 3.8, lLocal
    ImprimeCentralizado "CONT.DE ORDEM", lColunaI + 1, lColunaI + 3.9, lLinhaI + 3.05, lLocal
    ImprimeCentralizado "DE OPERAÇÃO", lColunaI + 1, lColunaI + 3.9, lLinhaI + 3.4, lLocal
    ImprimeCentralizado "INICIAL", lColunaI + 1, lColunaI + 2.4, lLinhaI + 3.9, lLocal
    ImprimeCentralizado "FINAL", lColunaI + 2.4, lColunaI + 3.9, lLinhaI + 3.9, lLocal
    ImprimeCentralizado "TOTALIZADOR GERAL (GT)", lColunaI + 3.9, lColunaI + 9.3, lLinhaI + 3.2, lLocal
    ImprimeCentralizado "FINAL", lColunaI + 3.9, lColunaI + 6.6, lLinhaI + 3.9, lLocal
    ImprimeCentralizado "INICIAL", lColunaI + 6.6, lColunaI + 9.3, lLinhaI + 3.9, lLocal
    ImprimeCentralizado "CANCELAM.", lColunaI + 9.3, lColunaI + 11.3, lLinhaI + 3.4, lLocal
    ImprimeCentralizado "DE ITEM", lColunaI + 9.3, lColunaI + 11.3, lLinhaI + 3.8, lLocal
    ImprimeCentralizado "VALOR", lColunaI + 11.3, lColunaI + 13.5, lLinhaI + 3.4, lLocal
    ImprimeCentralizado "CONTÁBIL", lColunaI + 11.3, lColunaI + 13.5, lLinhaI + 3.8, lLocal
    ImprimeCentralizado "BASE DE CÁLCULO", lColunaI + 18.99, lColunaI + 22.65, lLinhaI + 3.2, lLocal
    ImprimeCentralizado "SUBSTIT.", lColunaI + 13.5, lColunaI + 15.33, lLinhaI + 3.75, lLocal
    ImprimeCentralizado "TRIBUTÁRIA", lColunaI + 13.5, lColunaI + 15.33, lLinhaI + 4.05, lLocal
    
    ImprimeCentralizado "ISENÇÃO", lColunaI + 15.33, lColunaI + 17.16, lLinhaI + 3.75, lLocal
    ImprimeCentralizado "TRIBUTÁRIA", lColunaI + 15.33, lColunaI + 17.16, lLinhaI + 4.05, lLocal
    
    ImprimeCentralizado "NÃO", lColunaI + 17.16, lColunaI + 18.99, lLinhaI + 3.75, lLocal
    ImprimeCentralizado "INCIDÊNCIA", lColunaI + 17.16, lColunaI + 18.99, lLinhaI + 4.05, lLocal
    
    ImprimeCentralizado "12%", lColunaI + 18.99, lColunaI + 20.82, lLinhaI + 3.9, lLocal
    ImprimeCentralizado "17%", lColunaI + 20.82, lColunaI + 22.65, lLinhaI + 3.9, lLocal
    Printer.Line (lColunaI + 17.16, lLinhaI + 3.7)-(lColunaI + 17.16, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.99, lLinhaI + 3.7)-(lColunaI + 18.99, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20.82, lLinhaI + 3.7)-(lColunaI + 20.82, lLinhaI + lLinha), RGB(xR, xG, xB)
    
    
    
    ImprimeCentralizado "IMPOSTO", lColunaI + 22.65, lColunaI + 24.48, lLinhaI + 3.4, lLocal
    ImprimeCentralizado "DEBITADO", lColunaI + 22.65, lColunaI + 24.48, lLinhaI + 3.8, lLocal
    
    ImprimeCentralizado "CONT.", lColunaI + 24.48, lColunaI + 25.98, lLinhaI + 3.1, lLocal
    ImprimeCentralizado "RED.", lColunaI + 24.48, lColunaI + 25.98, lLinhaI + 3.5, lLocal
    ImprimeCentralizado "Z", lColunaI + 24.48, lColunaI + 25.98, lLinhaI + 3.9, lLocal
    
    Printer.Line (lColunaI + 11, lLinhaI + 2)-(lColunaI + 11, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 17, lLinhaI + 0)-(lColunaI + 17, lLinhaI + 1), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 19, lLinhaI + 2)-(lColunaI + 19, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20.1, lLinhaI + 1)-(lColunaI + 20.1, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22, lLinhaI + 0)-(lColunaI + 22, lLinhaI + 1), RGB(xR, xG, xB)
    
    
    'Linhas do Detalhe
    Printer.Line (lColunaI + 1, lLinhaI + 3.7)-(lColunaI + 9.3, lLinhaI + 3.7), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13.5, lLinhaI + 3.7)-(lColunaI + 22.65, lLinhaI + 3.7), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 4.4)-(lColunaI + 26, lLinhaI + 4.4), RGB(xR, xG, xB)
    
    
    lLinha = 4.4
    'Imprime linha que sai logo abaixo dos dados da Reducao Z
    For i = 1 To lQtdReducao
        lLinha = lLinha + 0.5
        Printer.Line (lColunaI + 0, lLinhaI + lLinha)-(lColunaI + 26, lLinhaI + lLinha), RGB(xR, xG, xB)
    Next
    ImprimeCentralizado "TOTAIS DO DIA", lColunaI, lColunaI + 3.9, lLinhaI + lLinha + 0.1, lLocal
    
    'Imprime linha que sai logo abaixo do "Totais do dia"
    lLinha = lLinha + 0.5
    lLinhaDepoisTotal = lLinha
    Printer.Line (lColunaI + 0, lLinhaI + lLinha)-(lColunaI + 26, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 3.9, lLinhaI + 3)-(lColunaI + 3.9, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 6.6, lLinhaI + 3.7)-(lColunaI + 6.6, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 9.3, lLinhaI + 3)-(lColunaI + 9.3, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 11.3, lLinhaI + 3)-(lColunaI + 11.3, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13.5, lLinhaI + 3)-(lColunaI + 13.5, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 15.33, lLinhaI + 3.7)-(lColunaI + 15.33, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 17.16, lLinhaI + 3.7)-(lColunaI + 17.16, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.99, lLinhaI + 3)-(lColunaI + 18.99, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20.82, lLinhaI + 3.7)-(lColunaI + 20.82, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22.65, lLinhaI + 3)-(lColunaI + 22.65, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 24.48, lLinhaI + 3)-(lColunaI + 24.48, lLinhaI + lLinha), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 1, lLinhaI + 3)-(lColunaI + 1, lLinhaI + lLinha - 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 2.4, lLinhaI + 3.7)-(lColunaI + 2.4, lLinhaI + lLinha - 0.5), RGB(xR, xG, xB)
    
    'Linhas do Rodapé
    ImprimeCentralizado "OBSERVAÇÕES", lColunaI, lColunaI + 13, lLinhaI + lLinha + 0.1, lLocal
    ImprimeCentralizado "RESPONSÁVEL PELO ESTABELECIMENTO", lColunaI + 13, lColunaI + 26, lLinhaI + lLinha + 0.1, lLocal
    
    
    'Imprime linha que sai logo abaixo do "Observações"
    lLinha = lLinha + 0.5
    Printer.Line (lColunaI + 0, lLinhaI + lLinha)-(lColunaI + 26, lLinhaI + lLinha), RGB(xR, xG, xB)
    ImprimeTexto "NOME", lColunaI + 13.2, lColunaI + 26, lLinhaI + lLinha + 0.1, lLocal
    'Imprime linha que sai logo abaixo do "Nome xxxxxxxxxx  do Responsável pela empresa"
    lLinha = lLinha + 1.3
    Printer.Line (lColunaI + 13, lLinhaI + lLinha)-(lColunaI + 26, lLinhaI + lLinha), RGB(xR, xG, xB)
    ImprimeTexto "FUNÇÃO", lColunaI + 13.2, lColunaI + 19, lLinhaI + lLinha + 0.1, lLocal
    ImprimeTexto "ASSINATURA", lColunaI + 19.2, lColunaI + 26, lLinhaI + lLinha + 0.1, lLocal
    
    Printer.Line (lColunaI + 13, lLinhaI + lLinhaDepoisTotal)-(lColunaI + 13, lLinhaI + lLinhaDepoisTotal + 3.1), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 19, lLinhaI + lLinhaDepoisTotal + 1.8)-(lColunaI + 19, lLinhaI + lLinhaDepoisTotal + 3.1), RGB(xR, xG, xB)
    
    Printer.FontBold = False
    Printer.FontSize = 6
    ImprimeValor "V." & gVersaoSGP, lColunaI + 24, lColunaI + 26, lLinhaI + lLinhaDepoisTotal + 3.1, lLocal
    Printer.FontBold = False
    Printer.FontSize = 8
    
    Exit Sub
    
    'Printer.DrawWidth = 1
    'Printer.EndDoc
ErrorRotina:
    MsgBox "Erro Não Identificado!", vbCritical, "Erro na Rotina: ImpGrade"
End Sub
Private Sub RecalculaContadorInicialFinal()
    On Error GoTo FileError
    
    If MovMapaResumo.LocalizarPrimeiro Then
        Do Until MovMapaResumo.LocalizarProximo = False
            If MovMapaResumo.Data >= CDate("01/04/2007") Then
                If MovCupomFiscal.LocalizarPrimeiroData(g_empresa, 0, MovMapaResumo.Data) Then
                    MovMapaResumo.ContagemOperacaoInicial = MovCupomFiscal.NumeroCupom
                End If
                If MovCupomFiscal.LocalizarUltimoData(g_empresa, 0, MovMapaResumo.Data) Then
                    MovMapaResumo.ContagemOperacaoFinal = MovCupomFiscal.NumeroCupom
                End If
                If Not MovMapaResumo.Alterar(g_empresa, 1, MovMapaResumo.Data, MovMapaResumo.numero) Then
                    MsgBox "Nao foi possível alterar o mapa resumo.", vbInformation, "Erro de Integridade!"
                End If
            End If
        Loop
        MsgBox "Processamento concluido com sucesso.", vbInformation, "Operacao concluida!"
    Else
        MsgBox "Nao foi possível localizar registro.", vbInformation, "Erro!"
    End If
    Exit Sub
    
FileError:
    MsgBox "Erro ao processar", vbInformation, "RecalculaContadorInicialFinal"
End Sub
Private Sub RecalculaSubstituicao()
    Dim xValor As Currency
    
    xValor = 0
    If fValidaValor(txt_gt_inicial.Text) <> 0 And fValidaValor(txt_gt_final.Text) Then
        xValor = fValidaValor(txt_gt_final.Text) - fValidaValor(txt_gt_inicial.Text)
        xValor = xValor - fValidaValor(txt_cancelamento_item.Text)
        xValor = xValor - fValidaValor(txtDescontoIcms.Text)
        xValor = xValor + fValidaValor(txtAcrescimoIcms.Text)
    End If
    If lOpcao = 1 Then
        txt_substituicao.Text = Format(xValor, "###,###,##0.00")
    End If
End Sub
Private Sub RecalculaValorContabil()
    Dim xValor As Currency
    
    xValor = 0
    If fValidaValor(txt_gt_inicial.Text) <> 0 And fValidaValor(txt_gt_final.Text) Then
        xValor = fValidaValor(txt_gt_final.Text) - fValidaValor(txt_gt_inicial.Text)
        xValor = xValor - fValidaValor(txt_cancelamento_item.Text)
        xValor = xValor - fValidaValor(txtDescontoIcms.Text)
        'xValor = xValor + fValidaValor(txtAcrescimoIcms.Text)
    End If
    txt_valor_contabil.Text = Format(xValor, "###,###,##0.00")
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    lSairDiferenca = False
    If UCase(g_nome_empresa) Like "*VENTANIA*" Then
        If (MsgBox("Sair Diferenca no mapa resumo?", vbQuestion + vbYesNo + vbDefaultButton1, "Sair Diferenca?")) = vbYes Then
            lSairDiferenca = True
        End If
    End If
    lImprimeTodasReducoesDia = False
    lQtdReducao = MovMapaResumo.QuantidadeMapaResumoData(g_empresa, CDate(msk_data.Text))
    If lQtdReducao > 1 Then
        If (MsgBox("Existe " & lQtdReducao & " reduções Z nesta data." & vbCrLf & "Deseja imprimir todas totalizando em um único relatório?", vbQuestion + vbYesNo + vbDefaultButton1, "Impressão Geral?")) = vbYes Then
            lImprimeTodasReducoesDia = True
            lCodigoEcfImprimir = 0
            If Not MovMapaResumo.LocalizarPrimeiroData(g_empresa, CDate(msk_data.Text)) Then
                MsgBox "Não foi possível localizar o primeiro mapa resumo desta data!", vbCritical, "Erro de Integridade!"
            End If
        Else
            lQtdReducao = 1
        End If
    End If
    ImpGrade
    ImpDados
    cmd_sair.SetFocus
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    If MovMapaResumo.LocalizarUltimo(g_empresa) Then
        msk_data.Text = Format(MovMapaResumo.Data + 1, "dd/mm/yyyy")
        txt_numero_reducao.Text = Format(MovMapaResumo.numero + 1, "##,###,##0")
        txt_ecf_numero.Text = Format(MovMapaResumo.ECFNumero, "#0")
        txt_contador_inicial.Text = Format(MovMapaResumo.ContagemOperacaoFinal, "##,###,##0")
        txt_gt_inicial.Text = Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00")
        txtContagemReinicioOperacao.Text = Format(MovMapaResumo.ContagemReinicioOperacao, "##0")
        txt_Contador_Reducoes_Z.Text = Format(MovMapaResumo.ContadorReducoesZ + 1, "##,###,##0")
        txt_contador_inicial.SetFocus
        'txt_numero_reducao.SetFocus
    Else
        msk_data.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    msk_data.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovMapaResumo.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If MovMapaResumo.LocalizarCodigo(g_empresa, lEcfNumero, lData, lNumero) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    msk_data.Text = "__/__/____"
    txt_numero_reducao.Text = ""
    txt_ecf_numero.Text = ""
    txt_contador_inicial.Text = ""
    txt_contador_final.Text = ""
    txtContagemReinicioOperacao.Text = ""
    txt_gt_final.Text = ""
    txt_gt_inicial.Text = ""
    txt_cancelamento_item.Text = ""
    txtDescontoIcms.Text = ""
    txtAcrescimoIcms.Text = ""
    txt_valor_contabil.Text = ""
    txt_substituicao.Text = ""
    txt_isentas.Text = ""
    txtNaoIncidencia.Text = ""
    txtICMS3.Text = ""
    txtIcms7.Text = ""
    txtIcms12.Text = ""
    txtIcms13.Text = ""
    txt_icms_17.Text = ""
    txtIcms19.Text = ""
    txtIcms25.Text = ""
    txt_Contador_Reducoes_Z = ""
    txt_observacao_1.Text = ""
    txt_observacao_2.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If msk_data.Text <> "00:00:00" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & MovMapaResumo.Data & " Ecf:" & MovMapaResumo.ECFNumero)
            If MovMapaResumo.Excluir(g_empresa, lEcfNumero, lData, lNumero) Then
                LimpaTela
                If MovMapaResumo.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtivaBotoes
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Registro não excluido!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_imprimir_Click()
    Call GravaAuditoria(1, Me.name, 7, "Data:" & MovMapaResumo.Data & " Ecf:" & MovMapaResumo.ECFNumero)
    If SelecionaImpressoraHP(Me) Then
        Relatorio
        If MovMapaResumo.LocalizarCodigo(g_empresa, lEcfNumero, lData, lNumero) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    frm_dados.Enabled = True
    Inclui
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 18 Then 'Crtl + R
        KeyAscii = 0
        'ZzLoopGravaMapaResumo
        'If Date <= CDate("10/05/2007") Then
        '    RecalculaContadorInicialFinal
        'End If
    ElseIf KeyAscii = 22 Then 'Crtl + V
        KeyAscii = 0
        ZzLoopGravaMapaResumoCat52
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, "Data:" & MovMapaResumo.Data & " Ecf:" & MovMapaResumo.ECFNumero)
            If MovMapaResumo.Incluir Then
                lData = msk_data.Text
                lNumero = CLng(txt_numero_reducao.Text)
                lEcfNumero = Val(txt_ecf_numero.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De Data:" & MovMapaResumo.Data & " Ecf:" & MovMapaResumo.ECFNumero)
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, "Para Data:" & MovMapaResumo.Data & " Ecf:" & MovMapaResumo.ECFNumero)
            If MovMapaResumo.Alterar(g_empresa, lEcfNumero, lData, lNumero) Then
                lData = msk_data.Text
                lNumero = CLng(txt_numero_reducao.Text)
                lEcfNumero = Val(txt_ecf_numero.Text)
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If MovMapaResumo.LocalizarCodigo(g_empresa, lEcfNumero, lData, lNumero) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
        lOpcao = 0
        cmd_imprimir.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    Dim xDiferencaGT As Currency
    Dim xSomaCanceladosContabil As Currency
    Dim xSomaBaseDeCalculo As Currency
    
    xDiferencaGT = fValidaValor(txt_gt_final.Text) - fValidaValor(txt_gt_inicial.Text)
    xSomaCanceladosContabil = fValidaValor(txt_cancelamento_item.Text) + fValidaValor(txtDescontoIcms.Text) + fValidaValor(txt_valor_contabil.Text)
    xSomaBaseDeCalculo = fValidaValor(txt_substituicao.Text) + fValidaValor(txt_isentas.Text) + fValidaValor(txtNaoIncidencia.Text) + fValidaValor(txtICMS3.Text) + fValidaValor(txtIcms7.Text) + fValidaValor(txtIcms12.Text) + fValidaValor(txtIcms13.Text) + fValidaValor(txt_icms_17.Text) + fValidaValor(txtIcms19.Text) + fValidaValor(txtIcms25.Text) + fValidaValor(txt_cancelamento_item.Text) + fValidaValor(txtDescontoIcms.Text)
    
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not Val(txt_numero_reducao.Text) > 0 Then
        MsgBox "Informe o número da redução Z.", vbInformation, "Atenção!"
        txt_numero_reducao.SetFocus
    ElseIf Not Val(txt_ecf_numero.Text) > 0 Then
        MsgBox "Informe o número do ECF.", vbInformation, "Atenção!"
        txt_ecf_numero.SetFocus
    ElseIf Not Val(txt_contador_inicial.Text) > 0 Then
        MsgBox "Informe a contagem de ordem de operação inicial.", vbInformation, "Atenção!"
        txt_contador_inicial.SetFocus
    ElseIf Not Val(txt_contador_final.Text) > 0 Then
        MsgBox "Informe a contagem de ordem de operação final.", vbInformation, "Atenção!"
        txt_contador_final.SetFocus
    ElseIf CLng(txt_contador_final.Text) < CLng(txt_contador_inicial.Text) Then
        MsgBox "A contagem de ordem de operação final está menor que a inicial.", vbInformation, "Atenção!"
        txt_contador_final.SetFocus
    ElseIf Not Val(txtContagemReinicioOperacao.Text) > 0 Then
        MsgBox "Informe o contador de reinício de operação.", vbInformation, "Atenção!"
        txtContagemReinicioOperacao.SetFocus
    ElseIf Not fValidaValor(txt_gt_final.Text) > 0 Then
        MsgBox "Informe o totalizador geral final.", vbInformation, "Atenção!"
        txt_gt_final.SetFocus
    ElseIf Not fValidaValor(txt_gt_inicial.Text) > 0 Then
        MsgBox "Informe o totalizador geral inicial.", vbInformation, "Atenção!"
        txt_gt_inicial.SetFocus
    ElseIf fValidaValor(txt_gt_inicial.Text) > fValidaValor(txt_gt_final.Text) Then
        MsgBox "O totalizador geral inicial está maior que o final.", vbInformation, "Atenção!"
        txt_gt_inicial.SetFocus
    ElseIf Not fValidaValor(txt_valor_contabil.Text) >= 0 Then
        MsgBox "Informe o valor contábil.", vbInformation, "Atenção!"
        txt_valor_contabil.SetFocus
    ElseIf Not Val(txt_Contador_Reducoes_Z.Text) > 0 Then
        MsgBox "Informe o Contador de Redução Z.", vbInformation, "Atenção!"
        txt_Contador_Reducoes_Z.SetFocus
    ElseIf Not xDiferencaGT = xSomaCanceladosContabil Then
        MsgBox "A diferença do Totalizador Geral (GT) é: " & Format(xDiferencaGT, "###,##0.00") & "." & Chr(10) & "Porém a soma de Cancelamento + Desconto + Valor Contabil - Acréscimo está em: " & Format(xSomaCanceladosContabil, "###,##0.00") & "." & Chr(10) & "Verifique e corrija a diferença.", vbInformation, "Erro de Consistência!"
        txt_valor_contabil.SetFocus
    ElseIf Not xDiferencaGT = xSomaBaseDeCalculo Then
        MsgBox "A diferença do Totalizador Geral (GT) é: " & Format(xDiferencaGT, "###,##0.00") & "." & Chr(10) & "Porém a soma de Substituição + Isentas + Não Incidência + ICMS 7% + ICMS 12% + ICMS 13% + ICMS 17% + ICMS 25% - Cancelamento - Desconto + Acrescimo, está em: " & Format(xSomaBaseDeCalculo, "###,##0.00") & "." & Chr(10) & "Verifique e corrija a diferença.", vbInformation, "Erro de Consistência!"
        txt_isentas.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLocal = 1
    lLinhaI = CCur(ReadINI("MAPA RESUMO", "Margem Superior", gArquivoIni))
    lColunaI = CCur(ReadINI("MAPA RESUMO", "Margem Esquerda", gArquivoIni))
    
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 20
    lLinhaTab = 0
    lCancelamentoItem = 0
    lAcrescimo = 0
    lValorContabil = 0
    lIsentas = 0
    lNaoIncidencia = 0
    lSubstituicaoTributaria = 0
    lICMS12 = 0
    lICMS17 = 0
    lImpostoDebitado = 0
    lCodigoEcfImprimir = Val(txt_ecf_numero.Text)
End Sub
Private Sub ZzLoopGravaMapaResumo()
    Dim xString As String
    Dim xData As Date
    Dim xDataF As Date
    Dim i As Integer
    Dim xColuna17 As Integer
    Dim xECF As Integer
    'Dim xString2 As String
        
    On Error GoTo FileError
    
    xECF = txt_ecf_numero.Text
    xColuna17 = 118
    xString = ReadINI("MAPA RESUMO", "COLUNA ICMS 17,00%", gArquivoIni)
    If Len(xString) > 0 Then
        xColuna17 = Val(xString)
    End If
    
    xData = CDate("01/08/2010")
    xDataF = CDate("31/08/2010")
    Do Until xData > xDataF
        xString = ZzLeArquivoLogReducaoZ(xData)
        If Len(xString) = 631 Then
            If Not MovMapaResumo.LocalizarDataECF(g_empresa, xData, xECF) Then
                If xData = CDate(Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2)) Then
                    If MovMapaResumo.LocalizarAnteriorDataECF(g_empresa, xECF, xData) Then
                        MovMapaResumo.numero = MovMapaResumo.numero + 1
                        MovMapaResumo.ContagemOperacaoInicial = MovMapaResumo.ContagemOperacaoFinal + 1
                        MovMapaResumo.TotalizadorGeralInicial = MovMapaResumo.TotalizadorGeralFinal
                    Else
                        MovMapaResumo.numero = 1
                        MovMapaResumo.ContagemOperacaoInicial = 0
                        MovMapaResumo.TotalizadorGeralInicial = 0
                        MovMapaResumo.ContadorReducoesZ = 0
                        MovMapaResumo.ContagemReinicioOperacao = 1
                    End If
                
                    MovMapaResumo.Data = xData
                    MovMapaResumo.ECFNumero = xECF
                    MovMapaResumo.ContagemOperacaoFinal = Mid(xString, 579, 6)
                    MovMapaResumo.TotalizadorGeralFinal = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                    MovMapaResumo.CancelamentoItem = fValidaValor(Mid(xString, 23, 12) & "," & Mid(xString, 35, 2))
                    MovMapaResumo.Desconto = fValidaValor(Mid(xString, 38, 12) & "," & Mid(xString, 50, 2))
                    MovMapaResumo.Acrescimo = fValidaValor(Mid(xString, 603, 12) & "," & Mid(xString, 615, 2))
                    MovMapaResumo.ValorContabil = 0
                    MovMapaResumo.Isentas = fValidaValor(Mid(xString, 342, 12) & "," & Mid(xString, 354, 2))
                    MovMapaResumo.NaoIncidencia = fValidaValor(Mid(xString, 356, 12) & "," & Mid(xString, 368, 2))
                    MovMapaResumo.SubstituicaoTributaria = fValidaValor(Mid(xString, 370, 12) & "," & Mid(xString, 382, 2))
                    'If UCase(g_nome_empresa) Like "*VALPOSTO*" Or UCase(g_nome_empresa) Like "*VW COMERCIO*" Then
                    '    MovMapaResumo.Isentas = fValidaValor(Mid(xString, 356, 12) & "," & Mid(xString, 368, 2))
                    'End If
                    MovMapaResumo.ICMS12 = 0
                    If UCase(g_nome_empresa) Like "*VENTANIA*" Then
                        MovMapaResumo.ICMS17 = fValidaValor(Mid(xString, 132, 12) & "," & Mid(xString, 144, 2))
                    Else
                        MovMapaResumo.ICMS17 = fValidaValor(Mid(xString, xColuna17, 12) & "," & Mid(xString, xColuna17 + 12, 2))
                    End If
                    MovMapaResumo.ValorContabil = MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.NaoIncidencia + MovMapaResumo.Isentas + MovMapaResumo.ICMS12 + MovMapaResumo.ICMS17
                    MovMapaResumo.ContadorReducoesZ = MovMapaResumo.ContadorReducoesZ + 1
                    MovMapaResumo.Observacao1 = ""
                    MovMapaResumo.Observacao2 = ""
                    If Not MovMapaResumo.Incluir Then
                        MsgBox "Não foi possível incluir o registro do Mapa Resumo!", vbInformation, "Erro de Verificação!"
                    End If
                Else
                    MsgBox "Data do arquivo não he igual a " & xData & vbCrLf & "Data do arquivo=" & Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2)
                End If
            Else
                MsgBox "Reducao Z Existente em " & xData
            End If
        End If
        xData = DateAdd("d", 1, xData)
    Loop
    Exit Sub

FileError:
    MsgBox " - ZzLoopGravaMapaResumo: Erro ao Gravar o Mapa Resumo" & xString
    Exit Sub
End Sub
Private Sub ZzLoopGravaMapaResumoCat52()
    Dim i As Integer
    Dim xString As String
    Dim xFiltro As String
    Dim xNomeArquivoCat52 As String
    Dim xExtensaoArquivo As String
    Dim xNomeDiretorio As String
    Dim xNumeroEcf As Integer
    Dim xCOO As Long
    Dim xCRO As Integer
    Dim xCRZ As Long
    Dim xData As Date
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xGrandeTotal As Currency
    Dim xVendaBruta As Currency
    Dim xAcrescimo As Currency
    Dim xDesconto As Currency
    Dim xCancelamento As Currency
    Dim xSubstituicao As Currency
    Dim xIsenta As Currency
    Dim xNaoIncidencia As Currency
    Dim xIcms12 As Currency
    Dim xIcms17 As Currency
    Dim xImpDaruma As Boolean
    
    On Error GoTo FileError
    
    xImpDaruma = False
    xString = InputBox("Informe a Data inicial no formato dd/mm/yyyy.", "Data Inicial!", "")
    If Not IsDate(xString) Then
        Exit Sub
    End If
    xDataInicial = CDate(xString)
    
    xString = InputBox("Informe a Data final no formato dd/mm/yyyy.", "Data Final!", "")
    If Not IsDate(xString) Then
        Exit Sub
    End If
    xDataFinal = CDate(xString)
    
    xFiltro = "*.*"
    CommonDialog1.Filter = xFiltro
    CommonDialog1.ShowOpen
    xNomeArquivoCat52 = CommonDialog1.Filename
    xNomeDiretorio = Mid(xNomeArquivoCat52, 1, Len(xNomeArquivoCat52) - 12)
    xNomeArquivoCat52 = Mid(xNomeArquivoCat52, Len(xNomeArquivoCat52) - 11, 9)
    If Mid(xNomeArquivoCat52, 1, 2) = "DR" Then
        xImpDaruma = True
    End If
    

    
    For xData = xDataInicial To xDataFinal
        xExtensaoArquivo = Cat52ConverteDiaBema(Format(xData, "dd"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Format(xData, "mm"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Mid(Format(xData, "yyyy"), 3, 2))
        Set gArquivoTMP = gArqTxt.OpenTextFile(xNomeDiretorio & xNomeArquivoCat52 & xExtensaoArquivo, ForReading)
        xGrandeTotal = 0
        xVendaBruta = 0
        xAcrescimo = 0
        xDesconto = 0
        xCancelamento = 0
        xSubstituicao = 0
        xIsenta = 0
        xNaoIncidencia = 0
        xIcms12 = 0
        xIcms17 = 0
        Do Until gArquivoTMP.AtEndOfStream
            xString = gArquivoTMP.ReadLine
            If Mid(xString, 1, 3) = "E01" Then
                xNumeroEcf = Val(Mid(xString, 96, 3))
                '28/04/2016
            'ElseIf Mid(xString, 1, 3) = "E14" Then
             '   xGrandeTotal = xGrandeTotal + CCur(Mid(xString, 67, 14) / 100)
              '  xAcrescimo = xAcrescimo + CCur(Mid(xString, 95, 13) / 100)
               ' xDesconto = xDesconto + CCur(Mid(xString, 81, 13) / 100)
             'ElseIf Mid(xString, 1, 3) = "E15" Then
              '  xCancelamento = xCancelamento + CCur(Mid(xString, 239, 13) / 100)
                '
            ElseIf Mid(xString, 1, 3) = "E11" Then
                xGrandeTotal = CCur(Mid(xString, 105, 18) / 100)
            ElseIf Mid(xString, 1, 3) = "E12" Then
                'xNumeroEcf = Val(Mid(xString, 45, 2))
                xCOO = CLng(Mid(xString, 53, 6))
                xCRO = Val(Mid(xString, 59, 6))
                xCRZ = CLng(Mid(xString, 47, 6))
                xData = CDate(Mid(xString, 71, 2) & "/" & Mid(xString, 69, 2) & "/" & Mid(xString, 65, 4))
                xVendaBruta = CCur(Mid(xString, 87, 14) / 100)
            ElseIf Mid(xString, 1, 3) = "E13" Then
                If Mid(xString, 53, 2) = "AT" Then
                    xAcrescimo = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 53, 5) = "Can-T" Then
                    xCancelamento = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 53, 2) = "DT" Then
                    xDesconto = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 53, 2) = "F1" Then
                    xSubstituicao = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 53, 2) = "I1" Then
                    xIsenta = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 53, 2) = "N1" Then
                    xNaoIncidencia = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 55, 3) = "T12" Then
                    xIcms12 = CCur(Mid(xString, 60, 13) / 100)
                ElseIf Mid(xString, 55, 3) = "T17" Then
                    xIcms17 = CCur(Mid(xString, 60, 13) / 100)
                End If
            End If
        Loop
        gArquivoTMP.Close
        
        If xGrandeTotal > 0 Then
            If Not MovMapaResumo.LocalizarDataECF(g_empresa, xData, xNumeroEcf) Then
    '            If xData = CDate(Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2)) Then
                If MovMapaResumo.LocalizarAnteriorDataECF(g_empresa, xNumeroEcf, xData) Then
                    MovMapaResumo.numero = MovMapaResumo.numero + 1
                    MovMapaResumo.ContagemOperacaoInicial = MovMapaResumo.ContagemOperacaoFinal + 1
                    MovMapaResumo.TotalizadorGeralInicial = MovMapaResumo.TotalizadorGeralFinal
                Else
                    MovMapaResumo.numero = 1
                    MovMapaResumo.ContagemOperacaoInicial = 0
                    MovMapaResumo.TotalizadorGeralInicial = 0
                    MovMapaResumo.ContadorReducoesZ = 0
                End If
            
                MovMapaResumo.Data = xData
                MovMapaResumo.ECFNumero = xNumeroEcf
                MovMapaResumo.ContagemOperacaoFinal = xCOO
                MovMapaResumo.ContagemReinicioOperacao = xCRO
                MovMapaResumo.TotalizadorGeralFinal = xGrandeTotal
                If xImpDaruma Then
                    MovMapaResumo.TotalizadorGeralFinal = MovMapaResumo.TotalizadorGeralInicial + xVendaBruta
                End If
                MovMapaResumo.CancelamentoItem = xCancelamento
                MovMapaResumo.Desconto = xDesconto
                MovMapaResumo.Acrescimo = xAcrescimo
                MovMapaResumo.ValorContabil = xVendaBruta - xCancelamento - xDesconto
                MovMapaResumo.Isentas = xIsenta
                MovMapaResumo.NaoIncidencia = xNaoIncidencia
                MovMapaResumo.SubstituicaoTributaria = xSubstituicao
                MovMapaResumo.ICMS12 = xIcms12
                MovMapaResumo.ICMS17 = xIcms17
                MovMapaResumo.ContadorReducoesZ = MovMapaResumo.ContadorReducoesZ + 1
                MovMapaResumo.Observacao1 = ""
                MovMapaResumo.Observacao2 = ""
                If Not MovMapaResumo.Incluir Then
                    MsgBox "Não foi possível incluir o registro do Mapa Resumo!", vbInformation, "Erro de Verificação!"
                End If
    '            Else
    '                MsgBox "Data do arquivo não he igual a " & xData & vbCrLf & "Data do arquivo=" & Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2)
    '            End If
            Else
                MsgBox "Reducao Z Existente em " & xData
            End If
        End If
    
    
    Next
    
    Exit Sub

FileError:
    MsgBox " - ZzLoopGravaMapaResumoCat52: Erro ao Gravar o Mapa Resumo" & xString
    
    Exit Sub
End Sub


Private Function ZzLeArquivoLogReducaoZ(ByVal pData As Date) As String
    Dim i As Integer
    Dim xString As String
    Dim xNomeArquivo As String
    Dim xDadosReducaoZ As String
        
    On Error GoTo FileError
    
    ZzLeArquivoLogReducaoZ = ""
    xDadosReducaoZ = ""
    i = 0
    pData = DateAdd("d", 1, pData)
    xNomeArquivo = "C:\Documents and Settings\Tasso\Meus documentos\ReducaoZ RioVerde ECF 01\"
    xNomeArquivo = xNomeArquivo & "ECF_" & Format(pData, "dd") & "_" & Format(pData, "mm") & "_" & Format(pData, "yyyy") & ".LOG"
    Set gArquivoTMP = gArqTxt.OpenTextFile(xNomeArquivo, ForReading)
    Do Until gArquivoTMP.AtEndOfStream
        xString = gArquivoTMP.ReadLine
        If Mid(xString, 12, 10) = "Reducao Z:" Then
            i = i + 1
            If i = 3 Then
                xDadosReducaoZ = Mid(xString, 23, 631)
                Exit Do
            ElseIf i > 3 Then
                Exit Do
            End If
        End If
    Loop
    gArquivoTMP.Close
    ZzLeArquivoLogReducaoZ = xDadosReducaoZ
    
    Exit Function

FileError:
    Call CriaLogCupom(Time & " - ZzLeArquivoLogReducaoZ: Erro ao acessar o arquivo de log do Mapa Resumo. Arquivo:" & xNomeArquivo)
    Exit Function
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_mapa_resumo.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lNumero = RetiraGString(2)
        lEcfNumero = RetiraGString(3)
        Call MovMapaResumo.LocalizarCodigo(g_empresa, lEcfNumero, lData, lNumero)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovMapaResumo.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovMapaResumo.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovMapaResumo.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If MovMapaResumo.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF6 And lOpcao = 0 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Call GravaAuditoria(1, Me.name, 1, "")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_imprimir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_imprimir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = Len(msk_data.Text)
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_reducao.SetFocus
    End If
End Sub
Private Sub txt_cancelamento_item_GotFocus()
    txt_cancelamento_item.SelStart = 0
    txt_cancelamento_item.SelLength = Len(txt_cancelamento_item.Text)
End Sub
Private Sub txt_cancelamento_item_LostFocus()
    If fValidaValor(Trim(txt_cancelamento_item.Text)) = 0 Then
        txt_cancelamento_item.Text = 0
    End If
    txt_cancelamento_item.Text = Format(txt_cancelamento_item.Text, "###,###,##0.00")
    RecalculaValorContabil
End Sub
Private Sub txt_contador_final_GotFocus()
    txt_contador_final.SelStart = 0
    txt_contador_final.SelLength = Len(txt_contador_final.Text)
End Sub
Private Sub txt_contador_final_LostFocus()
    txt_contador_final.Text = Format(txt_contador_final.Text, "##,###,##0")
End Sub
Private Sub txt_contador_inicial_GotFocus()
    txt_contador_inicial.SelStart = 0
    txt_contador_inicial.SelLength = Len(txt_contador_inicial.Text)
End Sub
Private Sub txt_contador_inicial_LostFocus()
    txt_contador_inicial.Text = Format(txt_contador_inicial.Text, "##,###,##0")
End Sub
Private Sub txt_Contador_Reducoes_Z_GotFocus()
    txt_Contador_Reducoes_Z.SelStart = 0
    txt_Contador_Reducoes_Z.SelLength = Len(txt_Contador_Reducoes_Z.Text)
End Sub
Private Sub txt_Contador_Reducoes_Z_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_1.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_Contador_Reducoes_Z_LostFocus()
    txt_Contador_Reducoes_Z.Text = Format(txt_Contador_Reducoes_Z.Text, "##,###,##0")
End Sub
Private Sub txt_ecf_numero_GotFocus()
    txt_ecf_numero.SelStart = 0
    txt_ecf_numero.SelLength = Len(txt_ecf_numero.Text)
End Sub
Private Sub txt_ecf_numero_LostFocus()
    txt_ecf_numero.Text = Format(txt_ecf_numero.Text, "#0")
    If lOpcao = 1 Then
        If Val(txt_ecf_numero.Text) > 0 And IsDate(msk_data.Text) Then
            If MovMapaResumo.LocalizarAnteriorDataECF(g_empresa, Val(txt_ecf_numero.Text), CDate(msk_data.Text)) Then
                msk_data.Text = Format(MovMapaResumo.Data + 1, "dd/mm/yyyy")
                txt_numero_reducao.Text = Format(MovMapaResumo.numero + 1, "##,###,##0")
                txt_ecf_numero.Text = Format(MovMapaResumo.ECFNumero, "#0")
                txt_contador_inicial.Text = Format(MovMapaResumo.ContagemOperacaoFinal + 1, "##,###,##0")
                txt_gt_inicial.Text = Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00")
                txt_Contador_Reducoes_Z.Text = Format(MovMapaResumo.ContadorReducoesZ + 1, "##,###,##0")
                txtContagemReinicioOperacao.Text = Format(MovMapaResumo.ContagemReinicioOperacao, "##0")
                txt_contador_final.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_gt_final_GotFocus()
    txt_gt_final.SelStart = 0
    txt_gt_final.SelLength = Len(txt_gt_final.Text)
End Sub
Private Sub txt_gt_final_LostFocus()
    If fValidaValor(Trim(txt_gt_final.Text)) = 0 Then
        txt_gt_final.Text = 0
    End If
    txt_gt_final.Text = Format(txt_gt_final.Text, "###,###,##0.00")
    RecalculaValorContabil
    RecalculaSubstituicao
End Sub
Private Sub txt_gt_inicial_GotFocus()
    txt_gt_inicial.SelStart = 0
    txt_gt_inicial.SelLength = Len(txt_gt_inicial.Text)
End Sub
Private Sub txt_gt_inicial_LostFocus()
    If fValidaValor(Trim(txt_gt_inicial.Text)) = 0 Then
        txt_gt_inicial.Text = 0
    End If
    txt_gt_inicial.Text = Format(txt_gt_inicial.Text, "###,###,##0.00")
End Sub


Private Sub txt_icms_17_GotFocus()
    txt_icms_17.SelStart = 0
    txt_icms_17.SelLength = Len(txt_icms_17.Text)
End Sub
Private Sub txt_icms_17_LostFocus()
    If fValidaValor(Trim(txt_icms_17.Text)) = 0 Then
        txt_icms_17.Text = 0
    End If
    txt_icms_17.Text = Format(txt_icms_17.Text, "###,###,##0.00")
End Sub
Private Sub txt_isentas_GotFocus()
    txt_isentas.SelStart = 0
    txt_isentas.SelLength = Len(txt_isentas.Text)
End Sub
Private Sub txt_isentas_LostFocus()
    If fValidaValor(Trim(txt_isentas.Text)) = 0 Then
        txt_isentas.Text = 0
    End If
    txt_isentas.Text = Format(txt_isentas.Text, "###,###,##0.00")
End Sub
Private Sub txt_numero_reducao_GotFocus()
    txt_numero_reducao.SelStart = 0
    txt_numero_reducao.SelLength = Len(txt_numero_reducao.Text)
End Sub
Private Sub txt_numero_reducao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_contador_final.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_ecf_numero_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_contador_inicial.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_contador_inicial_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_contador_final.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_contador_final_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtContagemReinicioOperacao.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_gt_final_keypress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_cancelamento_item.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_gt_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_gt_final.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_cancelamento_item_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtDescontoIcms.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_numero_reducao_LostFocus()
    txt_numero_reducao.Text = Format(txt_numero_reducao.Text, "##,###,##0")
End Sub
Private Sub txt_observacao_1_GotFocus()
    txt_observacao_1.SelStart = 0
    txt_observacao_1.SelLength = Len(txt_observacao_1.Text)
End Sub
Private Sub txt_observacao_2_GotFocus()
    txt_observacao_2.SelStart = 0
    txt_observacao_2.SelLength = Len(txt_observacao_2.Text)
End Sub
Private Sub txt_observacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_substituicao_GotFocus()
    txt_substituicao.SelStart = 0
    txt_substituicao.SelLength = Len(txt_substituicao.Text)
End Sub
Private Sub txt_substituicao_LostFocus()
    If fValidaValor(Trim(txt_substituicao.Text)) = 0 Then
        txt_substituicao.Text = 0
    End If
    txt_substituicao.Text = Format(txt_substituicao.Text, "###,###,##0.00")
End Sub
Private Sub txt_valor_contabil_GotFocus()
    txt_valor_contabil.SelStart = 0
    txt_valor_contabil.SelLength = Len(txt_valor_contabil.Text)
End Sub
Private Sub txt_valor_contabil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtAcrescimoIcms.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_isentas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtNaoIncidencia.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_substituicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_isentas.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_icms_17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtIcms19.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_observacao_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_2.SetFocus
    End If
End Sub
Private Sub txt_valor_contabil_LostFocus()
    If fValidaValor(Trim(txt_valor_contabil.Text)) = 0 Then
        txt_valor_contabil.Text = 0
    End If
    txt_valor_contabil.Text = Format(txt_valor_contabil.Text, "###,###,##0.00")
End Sub
Private Sub txtAcrescimoIcms_GotFocus()
    txtAcrescimoIcms.SelStart = 0
    txtAcrescimoIcms.SelLength = Len(txtAcrescimoIcms.Text)
End Sub
Private Sub txtAcrescimoIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_substituicao.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtAcrescimoIcms_LostFocus()
    If fValidaValor(Trim(txtAcrescimoIcms.Text)) = 0 Then
        txtAcrescimoIcms.Text = 0
    End If
    txtAcrescimoIcms.Text = Format(txtAcrescimoIcms.Text, "###,###,##0.00")
    'RecalculaValorContabil
End Sub
Private Sub txtContagemReinicioOperacao_GotFocus()
    txtContagemReinicioOperacao.SelStart = 0
    txtContagemReinicioOperacao.SelLength = Len(txtContagemReinicioOperacao.Text)
End Sub
Private Sub txtContagemReinicioOperacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_gt_final.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtContagemReinicioOperacao_LostFocus()
    txtContagemReinicioOperacao.Text = Format(txtContagemReinicioOperacao.Text, "##0")
End Sub
Private Sub txtDescontoIcms_GotFocus()
    txtDescontoIcms.SelStart = 0
    txtDescontoIcms.SelLength = Len(txtDescontoIcms.Text)
End Sub
Private Sub txtDescontoIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_contabil.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtDescontoIcms_LostFocus()
    If fValidaValor(Trim(txtDescontoIcms.Text)) = 0 Then
        txtDescontoIcms.Text = 0
    End If
    txtDescontoIcms.Text = Format(txtDescontoIcms.Text, "###,###,##0.00")
    RecalculaValorContabil
End Sub
Private Sub txtIcms12_GotFocus()
    txtIcms12.SelStart = 0
    txtIcms12.SelLength = Len(txtIcms12.Text)
End Sub
Private Sub txtIcms12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtIcms13.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtIcms12_LostFocus()
    If fValidaValor(Trim(txtIcms12.Text)) = 0 Then
        txtIcms12.Text = 0
    End If
    txtIcms12.Text = Format(txtIcms12.Text, "###,###,##0.00")
End Sub

Private Sub txtIcms25_GotFocus()
    txtIcms25.SelStart = 0 'posiciona o ponteiro no primeiro registro do campo
    txtIcms25.SelLength = Len(txtIcms25.Text) 'posiciona o ponteiro no ultimo registro do campo
End Sub
Private Sub txtIcms25_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then 'se tiver sido pressionado ponto(KeyAscii = 46)
        KeyAscii = 44 'converte o ponto para virgula (KeyAscii = 44)
    ElseIf KeyAscii = 13 Then 'se tiver sido digitado enter
        KeyAscii = 0 'anula o que foi pressicionado
        txt_Contador_Reducoes_Z.SetFocus 'muda o foco para o txt_Contador_Reducoes_Z
    End If
    Call ValidaValor(KeyAscii) 'passa pela função ValidaValor
End Sub
Private Sub txtIcms25_LostFocus()
If fValidaValor(Trim(txtIcms25.Text)) = 0 Then
        txtIcms25.Text = 0
    End If
    txtIcms25.Text = Format(txtIcms25.Text, "###,###,##0.00")
End Sub
Private Sub txtICMS3_GotFocus()
    txtICMS3.SelStart = 0 'posiciona o ponteiro no primeiro registro do campo
    txtICMS3.SelLength = Len(txtICMS3.Text) 'posiciona o ponteiro no ultimo registro do campo
End Sub
Private Sub txtICMS3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then 'se tiver sido pressionado ponto(KeyAscii = 46)
        KeyAscii = 44 'converte o ponto para virgula (KeyAscii = 44)
    ElseIf KeyAscii = 13 Then 'se tiver sido digitado enter
        KeyAscii = 0 'anula o que foi pressicionado
        txtIcms7.SetFocus 'muda o foco para o txt_Contador_Reducoes_Z
    End If
    Call ValidaValor(KeyAscii) 'passa pela função ValidaValor
End Sub

Private Sub txtICMS3_LostFocus()
    If fValidaValor(Trim(txtICMS3.Text)) = 0 Then
        txtICMS3.Text = 0
    End If
    txtICMS3.Text = Format(txtICMS3.Text, "###,###,##0.00")
End Sub

Private Sub txtIcms7_GotFocus()
    txtIcms7.SelStart = 0
    txtIcms7.SelLength = Len(txtIcms7.Text)
End Sub
Private Sub txtIcms7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtIcms12.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtIcms7_LostFocus()
    If fValidaValor(Trim(txtIcms7.Text)) = 0 Then
        txtIcms7.Text = 0
    End If
    txtIcms7.Text = Format(txtIcms7.Text, "###,###,##0.00")
End Sub

Private Sub txtIcms13_GotFocus()
    txtIcms13.SelStart = 0
    txtIcms13.SelLength = Len(txtIcms13.Text)
End Sub
Private Sub txtIcms13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_icms_17.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtIcms13_LostFocus()
    If fValidaValor(Trim(txtIcms13.Text)) = 0 Then
        txtIcms13.Text = 0
    End If
    txtIcms13.Text = Format(txtIcms13.Text, "###,###,##0.00")
End Sub
Private Sub txtIcms19_GotFocus()
    txtIcms19.SelStart = 0
    txtIcms19.SelLength = Len(txtIcms19.Text)
End Sub
Private Sub txtIcms19_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtIcms25.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtIcms19_LostFocus()
    If fValidaValor(Trim(txtIcms19.Text)) = 0 Then
        txtIcms19.Text = 0
    End If
    txtIcms19.Text = Format(txtIcms19.Text, "###,###,##0.00")
End Sub
Private Sub txtNaoIncidencia_GotFocus()
    txtNaoIncidencia.SelStart = 0
    txtNaoIncidencia.SelLength = Len(txtNaoIncidencia.Text)
End Sub
Private Sub txtNaoIncidencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtICMS3.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtNaoIncidencia_LostFocus()
    If fValidaValor(Trim(txtNaoIncidencia.Text)) = 0 Then
        txtNaoIncidencia.Text = 0
    End If
    txtNaoIncidencia.Text = Format(txtNaoIncidencia.Text, "###,###,##0.00")
End Sub
