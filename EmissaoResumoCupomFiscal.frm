VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoResumoCupomFiscal 
   Caption         =   "Resumo da Venda de Cupom Fiscal"
   ClientHeight    =   2685
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "EmissaoResumoCupomFiscal.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoResumoCupomFiscal.frx":030A
   ScaleHeight     =   2685
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "EmissaoResumoCupomFiscal.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza resumo do L.M.C."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "EmissaoResumoCupomFiscal.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime resumo do L.M.C."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "EmissaoResumoCupomFiscal.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoCupomFiscal.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoCupomFiscal.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoCupomFiscal.frx":6CBA
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
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoResumoCupomFiscal"
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

Dim lEcfSubstituicaoTrib As Currency
Dim lEcfIsencaoTrib As Currency
Dim lEcfNaoTrib As Currency
Dim lEcfTrib17 As Currency
Dim lEcfTrib12 As Currency
Dim lEcfVendaBruta As Currency
Dim lEcfDescontoCancelamento As Currency
Dim lEcfVendaLiquida As Currency
Dim lSistSubstituicaoTrib As Currency
Dim lSistIsencaoTrib As Currency
Dim lSistNaoTrib As Currency
Dim lSistTrib17 As Currency
Dim lSistTrib12 As Currency
Dim lSistVendaBruta As Currency
Dim lSistDescontoCancelamento As Currency
Dim lSistVendaLiquida As Currency
Dim lDifSubstituicaoTrib As Currency
Dim lDifIsencaoTrib As Currency
Dim lDifNaoTrib As Currency
Dim lDifTrib17 As Currency
Dim lDifTrib12 As Currency
Dim lDifVendaBruta As Currency
Dim lDifDescontoCancelamento As Currency
Dim lDifVendaLiquida As Currency

Dim Aliquota As New cAliquota
Dim MovimentoCupomFiscal As New cMovimentoCupomFiscal
Dim MovimentoMapaResumo As New cMovimentoMapaResumo
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Aliquota = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set MovimentoMapaResumo = Nothing
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    xLinha = "| RESUMO DE CUPOM FISCAL                                                                                            Goiânia, __/__/____ |"
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____                                                                                                 |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  DATA  DO  | TIPO  DO | SUBSTITUICAO|   ISENCAO   |     NAO     |   TRIBUTADA |  TRIBUTADA  | VENDA BRUTA | DESCONTO  E |    VALOR    |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  MOVIMENTO | MOVIMENTO|  TRIBUTARIA |  TRIBUTARIA |  TRIBUTADA  |         17% |        12%  |             | CANCELAMENTO|   CONTABIL  |"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDet(ByVal pData As Date, ByVal pTipoMovimento As String, ByVal pSubstituicaoTrib As Currency, ByVal pIsencaoTrib As Currency, ByVal pNaoTrib As Currency, ByVal pTrib17 As Currency, ByVal pTrib12 As Currency, ByVal pVendaBruta As Currency, ByVal pDescontoCancelamento As Currency, ByVal pVendaLiquida As Currency)
    Dim xLinha As String
    Dim i As Integer
    
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |          |             |             |             |             |             |             |             |             |"
    
    If pData = CDate("01/01/1900") Then
        Mid(xLinha, 4, 9) = " ** TOTAL"
    Else
        If pTipoMovimento = "SUB-TOTAL" Then
            Mid(xLinha, 4, 9) = "DIFERENCA"
        Else
            Mid(xLinha, 3, 10) = Format(pData, "dd/mm/yyyy")
        End If
    End If
    Mid(xLinha, 16, 9) = pTipoMovimento
    i = Len(Format(pSubstituicaoTrib, "####,##0.00"))
    Mid(xLinha, 27 + 11 - i, i) = Format(pSubstituicaoTrib, "####,##0.00")
    i = Len(Format(pIsencaoTrib, "####,##0.00"))
    Mid(xLinha, 41 + 11 - i, i) = Format(pIsencaoTrib, "####,##0.00")
    i = Len(Format(pNaoTrib, "####,##0.00"))
    Mid(xLinha, 55 + 11 - i, i) = Format(pNaoTrib, "####,##0.00")
    i = Len(Format(pTrib17, "####,##0.00"))
    Mid(xLinha, 69 + 11 - i, i) = Format(pTrib17, "####,##0.00")
    i = Len(Format(pTrib12, "####,##0.00"))
    Mid(xLinha, 83 + 11 - i, i) = Format(pTrib12, "####,##0.00")
    i = Len(Format(pVendaBruta, "####,##0.00"))
    Mid(xLinha, 97 + 11 - i, i) = Format(pVendaBruta, "####,##0.00")
    i = Len(Format(pDescontoCancelamento, "####,##0.00"))
    Mid(xLinha, 111 + 11 - i, i) = Format(pDescontoCancelamento, "####,##0.00")
    i = Len(Format(pVendaLiquida, "####,##0.00"))
    Mid(xLinha, 125 + 11 - i, i) = Format(pVendaLiquida, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    If lLinha >= 62 Then
        xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    Else
        xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
        BioImprime "@Printer.Print " & xLinha
    End If
    
    
    Call ImpDet(CDate("01/01/1900"), "ECF", lEcfSubstituicaoTrib, lEcfIsencaoTrib, lEcfNaoTrib, lEcfTrib17, lEcfTrib12, lEcfVendaBruta, lEcfDescontoCancelamento, lEcfVendaLiquida)
    Call ImpDet(CDate("01/01/1900"), "SISTEMA", lSistSubstituicaoTrib, lSistIsencaoTrib, lSistNaoTrib, lSistTrib17, lSistTrib12, lSistVendaBruta, lSistDescontoCancelamento, lSistVendaLiquida)
    
    lSistSubstituicaoTrib = lSistSubstituicaoTrib - lEcfSubstituicaoTrib
    lSistIsencaoTrib = lSistIsencaoTrib - lEcfIsencaoTrib
    lSistNaoTrib = lSistNaoTrib - lEcfNaoTrib
    lSistTrib17 = lSistTrib17 - lEcfTrib17
    lSistTrib12 = lSistTrib12 - lEcfTrib12
    lSistVendaBruta = lSistVendaBruta - lEcfVendaBruta
    lSistDescontoCancelamento = lSistDescontoCancelamento - lEcfDescontoCancelamento
    lSistVendaLiquida = lSistVendaLiquida - lEcfVendaLiquida
    Call ImpDet(CDate("01/01/1900"), "DIFERENCA", lSistSubstituicaoTrib, lSistIsencaoTrib, lSistNaoTrib, lSistTrib17, lSistTrib12, lSistVendaBruta, lSistDescontoCancelamento, lSistVendaLiquida)
    
    
    xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub Relatorio()
    Dim xData As Date
    
    ZeraVariaveis
    xData = CDate(msk_data_i.Text)
'    If chkSomenteResumo.Value = 1 Then
'        Call LoopCombustivel(CDate(msk_data_i.Text), CDate(msk_data_f.Text))
'    Else
        'Loop data
        Do Until xData > CDate(msk_data_f.Text)
            Call LoopDados(xData)
            'Call ImpDet(xData)
            'If chkDetalhadaBico.Value = 1 Then
            '    Call ImpDetBico(xData)
            'End If
            xData = xData + 1
            'ImpSubTotal
        Loop
'    End If
    ImpTotal
'    If chkImprimirDescontoUnitario.Value = 1 Then
'        ImpTotalDesconto
'    End If
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório do Resumo de Cupom Fiscal|@|"
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
    
    lEcfSubstituicaoTrib = 0
    lEcfIsencaoTrib = 0
    lEcfNaoTrib = 0
    lEcfTrib17 = 0
    lEcfTrib12 = 0
    lEcfVendaBruta = 0
    lEcfDescontoCancelamento = 0
    lEcfVendaLiquida = 0
    lSistSubstituicaoTrib = 0
    lSistIsencaoTrib = 0
    lSistNaoTrib = 0
    lSistTrib17 = 0
    lSistTrib12 = 0
    lSistVendaBruta = 0
    lSistDescontoCancelamento = 0
    lSistVendaLiquida = 0
    lDifSubstituicaoTrib = 0
    lDifIsencaoTrib = 0
    lDifNaoTrib = 0
    lDifTrib17 = 0
    lDifTrib12 = 0
    lDifVendaBruta = 0
    lDifDescontoCancelamento = 0
    lDifVendaLiquida = 0
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
End Sub
Private Sub LoopDados(ByVal pData As Date)
    Dim xLinha As String
    
    Dim xSubstituicaoTrib As Currency
    Dim xIsencaoTrib As Currency
    Dim xNaoTrib As Currency
    Dim xTrib17 As Currency
    Dim xTrib12 As Currency
    Dim xVendaBruta As Currency
    Dim xDescontoCancelamento As Currency
    Dim xVendaLiquida As Currency
    
    Dim xDifSubstituicaoTrib As Currency
    Dim xDifIsencaoTrib As Currency
    Dim xDifNaoTrib As Currency
    Dim xDifTrib17 As Currency
    Dim xDifTrib12 As Currency
    Dim xDifVendaBruta As Currency
    Dim xDifDescontoCancelamento As Currency
    Dim xDifVendaLiquida As Currency
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 62 Then
        xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "+------------+----------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    xSubstituicaoTrib = 0
    xIsencaoTrib = 0
    xNaoTrib = 0
    xTrib17 = 0
    xTrib12 = 0
    xVendaBruta = 0
    xDescontoCancelamento = 0
    xVendaLiquida = 0
    
    xDifSubstituicaoTrib = 0
    xDifIsencaoTrib = 0
    xDifNaoTrib = 0
    xDifTrib17 = 0
    xDifTrib12 = 0
    xDifVendaBruta = 0
    xDifDescontoCancelamento = 0
    xDifVendaLiquida = 0
    If MovimentoMapaResumo.LocalizarDataECF(g_empresa, pData, 1) Then
        xSubstituicaoTrib = MovimentoMapaResumo.SubstituicaoTributaria
        xIsencaoTrib = MovimentoMapaResumo.Isentas
        xNaoTrib = 0
        xTrib17 = MovimentoMapaResumo.ICMS17
        xTrib12 = 0
        xVendaBruta = MovimentoMapaResumo.TotalizadorGeralFinal - MovimentoMapaResumo.TotalizadorGeralInicial
        xDescontoCancelamento = MovimentoMapaResumo.CancelamentoItem
        xVendaLiquida = MovimentoMapaResumo.ValorContabil
    End If
    Call ImpDet(pData, "ECF", xSubstituicaoTrib, xIsencaoTrib, xNaoTrib, xTrib17, xTrib12, xVendaBruta, xDescontoCancelamento, xVendaLiquida)
    
    xDifSubstituicaoTrib = xSubstituicaoTrib
    xDifIsencaoTrib = xIsencaoTrib
    xDifNaoTrib = xNaoTrib
    xDifTrib17 = xTrib17
    xDifTrib12 = xTrib12
    xDifVendaBruta = xVendaBruta
    xDifDescontoCancelamento = xDescontoCancelamento
    xDifVendaLiquida = xVendaLiquida
    
    lEcfSubstituicaoTrib = lEcfSubstituicaoTrib + xSubstituicaoTrib
    lEcfIsencaoTrib = lEcfIsencaoTrib + xIsencaoTrib
    lEcfNaoTrib = lEcfNaoTrib + xNaoTrib
    lEcfTrib17 = lEcfTrib17 + xTrib17
    lEcfTrib12 = lEcfTrib12 + xTrib12
    lEcfVendaBruta = lEcfVendaBruta + xVendaBruta
    lEcfDescontoCancelamento = lEcfDescontoCancelamento + xDescontoCancelamento
    lEcfVendaLiquida = lEcfVendaLiquida + xVendaLiquida
    
    'Busca totalizadores do cupom fiscal no sistema
    xSubstituicaoTrib = 0
    xIsencaoTrib = 0
    xNaoTrib = 0
    xTrib17 = 0
    xTrib12 = 0
    xVendaBruta = 0
    xDescontoCancelamento = 0
    xVendaLiquida = 0
    If Aliquota.LocalizarNomeSemelhante("Substitui") Then
        xSubstituicaoTrib = MovimentoCupomFiscal.ValorAliquotaVendaData(g_empresa, pData, pData, 1, 9, Aliquota.Codigo)
    End If
    If Aliquota.LocalizarNomeSemelhante("Isen") Then
        xIsencaoTrib = MovimentoCupomFiscal.ValorAliquotaVendaData(g_empresa, pData, pData, 1, 9, Aliquota.Codigo)
    End If
    xNaoTrib = 0
    If Aliquota.LocalizarNomeSemelhante("17") Then
        xTrib17 = MovimentoCupomFiscal.ValorAliquotaVendaData(g_empresa, pData, pData, 1, 9, Aliquota.Codigo)
    End If
    xTrib12 = 0
    xDescontoCancelamento = MovimentoCupomFiscal.CancelamentoProdutoVendaData(g_empresa, 0, pData, pData, 1, 9, 0)
    xVendaLiquida = MovimentoCupomFiscal.ValorAliquotaVendaData(g_empresa, pData, pData, 1, 9, 0)
    xVendaBruta = xVendaLiquida + xDescontoCancelamento
    Call ImpDet(pData, "SISTEMA", xSubstituicaoTrib, xIsencaoTrib, xNaoTrib, xTrib17, xTrib12, xVendaBruta, xDescontoCancelamento, xVendaLiquida)
    
    lSistSubstituicaoTrib = lSistSubstituicaoTrib + xSubstituicaoTrib
    lEcfIsencaoTrib = lSistIsencaoTrib + xIsencaoTrib
    lSistNaoTrib = lSistNaoTrib + xNaoTrib
    lSistTrib17 = lSistTrib17 + xTrib17
    lSistTrib12 = lSistTrib12 + xTrib12
    lSistVendaBruta = lSistVendaBruta + xVendaBruta
    lSistDescontoCancelamento = lSistDescontoCancelamento + xDescontoCancelamento
    lSistVendaLiquida = lSistVendaLiquida + xVendaLiquida
    
    xDifSubstituicaoTrib = xSubstituicaoTrib - xDifSubstituicaoTrib
    xDifIsencaoTrib = xIsencaoTrib - xDifIsencaoTrib
    xDifNaoTrib = xNaoTrib - xDifNaoTrib
    xDifTrib17 = xTrib17 - xDifTrib17
    xDifTrib12 = xTrib12 - xDifTrib12
    xDifVendaBruta = xVendaBruta - xDifVendaBruta
    xDifDescontoCancelamento = xDescontoCancelamento - xDifDescontoCancelamento
    xDifVendaLiquida = xVendaLiquida - xDifVendaLiquida
    Call ImpDet(pData, "SUB-TOTAL", xDifSubstituicaoTrib, xDifIsencaoTrib, xDifNaoTrib, xDifTrib17, xDifTrib12, xDifVendaBruta, xDifDescontoCancelamento, xDifVendaLiquida)
    
    xDifSubstituicaoTrib = 0
    xDifIsencaoTrib = 0
    xDifNaoTrib = 0
    xDifTrib17 = 0
    xDifTrib12 = 0
    xDifVendaBruta = 0
    xDifDescontoCancelamento = 0
    xDifVendaLiquida = 0
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
