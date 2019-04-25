VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_analise_venda_cartao 
   Caption         =   "Emissão da Análise das Vendas c/ Cartão de Crédito"
   ClientHeight    =   2715
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "lst_analise_venda_cartao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_analise_venda_cartao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza análise das vendas com cartão de crédito."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "lst_analise_venda_cartao.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime análise das vendas com cartão de crédito."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "lst_analise_venda_cartao.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_venda_cartao.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_venda_cartao.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6000
         Picture         =   "lst_analise_venda_cartao.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
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
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
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
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   975
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
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_venda_cartao"
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
Dim l_faturamento(1 To 4) As Currency
Dim l_credito(1 To 4) As Currency
Dim l_administradora(1 To 4) As Currency
Dim l_liquido(1 To 4) As Currency
Dim l_total_faturamento(1 To 4) As Currency
Dim l_total_credito(1 To 4) As Currency
Dim l_total_administradora(1 To 4) As Currency
Dim l_total_liquido(1 To 4) As Currency
Dim l_dias(1 To 3) As Integer
Dim l_percentual(1 To 3) As Currency

Private CartaoCredito As New cCartaoCredito

Dim tbl_empresa As Table
Dim tbl_movimento_cartao_credito As Table
Private Sub BuscaPercentualCartoes()
    Dim i As Integer
    For i = 1 To 3
        If CartaoCredito.LocalizarCodigo(i) Then
            l_percentual(i) = CartaoCredito.TaxaCusto
            l_dias(i) = CartaoCredito.DiasPrazo
        End If
    Next
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set CartaoCredito = Nothing
    tbl_empresa.Close
    tbl_movimento_cartao_credito.Close
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 1 To 4
        l_faturamento(i) = 0
        l_credito(i) = 0
        l_administradora(i) = 0
        l_liquido(i) = 0
        l_total_faturamento(i) = 0
        l_total_credito(i) = 0
        l_total_administradora(i) = 0
        l_total_liquido(i) = 0
    Next
    For i = 1 To 3
        l_dias(i) = 0
        l_percentual(i) = 0
    Next
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
    ZeraVariaveis
    If tbl_movimento_cartao_credito.RecordCount > 0 Then
        LoopTabelaEmpresa
        If l_total_faturamento(4) > 0 Or l_total_credito(4) > 0 Then
            ImpTotal
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Análise da Venda com Cartão de Crédito|@|"
            frm_preview.Show 1
        End If
    End If
    cmd_sair.SetFocus
End Sub
Private Sub LoopMovimentoCartaoCreditoEmissao(x_empresa As Integer)
    Dim i As Integer
    With tbl_movimento_cartao_credito
        .Index = "id_data_emissao"
        .Seek ">=", x_empresa, CDate(msk_data_i), cbo_periodo_i, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                    Exit Do
                End If
                i = ![Codigo do Cartao]
                If i < 4 Then
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        l_faturamento(i) = l_faturamento(i) + !valor
                        l_faturamento(4) = l_faturamento(4) + !valor
                        l_total_faturamento(i) = l_total_faturamento(i) + !valor
                        l_total_faturamento(4) = l_total_faturamento(4) + !valor
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopMovimentoCartaoCreditoVencimento(x_empresa As Integer)
    Dim i As Integer
    With tbl_movimento_cartao_credito
        .Index = "id_data_vencimento2"
        .Seek ">=", x_empresa, CDate(msk_data_i), cbo_periodo_i, 0, CDate("01/01/1900")
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> x_empresa Or ![Data do Vencimento] > CDate(msk_data_f) Then
                    Exit Do
                End If
                i = ![Codigo do Cartao]
                If ![Codigo do Cartao] < 4 Then
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        l_credito(i) = l_credito(i) + !valor
                        l_credito(4) = l_credito(4) + !valor
                        l_total_credito(i) = l_total_credito(i) + !valor
                        l_total_credito(4) = l_total_credito(4) + !valor
                        l_administradora(i) = l_administradora(i) + Format(!valor * l_percentual(i) / 100, "00000000.00")
                        l_administradora(4) = l_administradora(4) + Format(!valor * l_percentual(i) / 100, "00000000.00")
                        l_total_administradora(i) = l_total_administradora(i) + Format(!valor * l_percentual(i) / 100, "00000000.00")
                        l_total_administradora(4) = l_total_administradora(4) + Format(!valor * l_percentual(i) / 100, "00000000.00")
                        l_liquido(i) = l_liquido(i) + Format(!valor - !valor * l_percentual(i) / 100, "00000000.00")
                        l_liquido(4) = l_liquido(4) + Format(!valor - !valor * l_percentual(i) / 100, "00000000.00")
                        l_total_liquido(i) = l_total_liquido(i) + Format(!valor - !valor * l_percentual(i) / 100, "00000000.00")
                        l_total_liquido(4) = l_total_liquido(4) + Format(!valor - !valor * l_percentual(i) / 100, "00000000.00")
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopTabelaEmpresa()
    Dim x_linha As String
    Dim i As Integer
    
    BuscaPercentualCartoes
    'loop tabela empresa
    With tbl_empresa
        .MoveFirst
        Do Until .EOF
            For i = 1 To 4
                l_faturamento(i) = 0
                l_credito(i) = 0
                l_administradora(i) = 0
                l_liquido(i) = 0
            Next
            Call LoopMovimentoCartaoCreditoEmissao(!Codigo)
            Call LoopMovimentoCartaoCreditoVencimento(!Codigo)
            If l_faturamento(4) > 0 Or l_credito(4) > 0 Then
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 55 Then
                    x_linha = "+------------------------------------------+----------------------+-------------+-------------+-------------+-------------+-------------+"
                    Mid(x_linha, 5, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                Call ImpDet(!Nome)
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+-------------+-------------+-------------+-------------+-------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = "*** TOTAL GERAL"
    Mid(x_linha, 46, 20) = "Faturamento         "
    i = Len(Format(l_total_faturamento(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_total_faturamento(3), "####,##0.00")
    i = Len(Format(l_total_faturamento(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_total_faturamento(2), "####,##0.00")
    i = Len(Format(l_total_faturamento(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_total_faturamento(1), "####,##0.00")
    i = Len(Format(l_total_faturamento(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_total_faturamento(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = "*** TOTAL GERAL"
    Mid(x_linha, 46, 20) = "Crédito             "
    i = Len(Format(l_total_credito(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_total_credito(3), "####,##0.00")
    i = Len(Format(l_total_credito(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_total_credito(2), "####,##0.00")
    i = Len(Format(l_total_credito(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_total_credito(1), "####,##0.00")
    i = Len(Format(l_total_credito(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_total_credito(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = "*** TOTAL GERAL"
    Mid(x_linha, 46, 20) = "Administradora      "
    i = Len(Format(l_total_administradora(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_total_administradora(3), "####,##0.00")
    i = Len(Format(l_total_administradora(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_total_administradora(2), "####,##0.00")
    i = Len(Format(l_total_administradora(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_total_administradora(1), "####,##0.00")
    i = Len(Format(l_total_administradora(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_total_administradora(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = "*** TOTAL GERAL"
    Mid(x_linha, 46, 20) = "Líquido             "
    i = Len(Format(l_total_liquido(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_total_liquido(3), "####,##0.00")
    i = Len(Format(l_total_liquido(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_total_liquido(2), "####,##0.00")
    i = Len(Format(l_total_liquido(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_total_liquido(1), "####,##0.00")
    i = Len(Format(l_total_liquido(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_total_liquido(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+------------------------------------------+----------------------+-------------+-------------+-------------+-------------+-------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & "* Faturamento    = Valor bruto das vendas com cartões no período"
    BioImprime "@Printer.Print " & "* Crédito        = Valor bruto dos cartões creditados no período"
    BioImprime "@Printer.Print " & "* Administradora = valor do faturamento líquido das administradoras no período"
    BioImprime "@Printer.Print " & "* Líquido        = Valor líquido dos cartões creditados no período"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(x_nome As String)
    Dim x_linha As String
    Dim i As Integer
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+-------------+-------------+-------------+-------------+-------------+"
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = x_nome
    Mid(x_linha, 46, 20) = "Faturamento         "
    i = Len(Format(l_faturamento(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_faturamento(3), "####,##0.00")
    i = Len(Format(l_faturamento(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_faturamento(2), "####,##0.00")
    i = Len(Format(l_faturamento(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_faturamento(1), "####,##0.00")
    i = Len(Format(l_faturamento(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_faturamento(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = x_nome
    Mid(x_linha, 46, 20) = "Crédito             "
    i = Len(Format(l_credito(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_credito(3), "####,##0.00")
    i = Len(Format(l_credito(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_credito(2), "####,##0.00")
    i = Len(Format(l_credito(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_credito(1), "####,##0.00")
    i = Len(Format(l_credito(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_credito(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = x_nome
    Mid(x_linha, 46, 20) = "Administradora      "
    i = Len(Format(l_administradora(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_administradora(3), "####,##0.00")
    i = Len(Format(l_administradora(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_administradora(2), "####,##0.00")
    i = Len(Format(l_administradora(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_administradora(1), "####,##0.00")
    i = Len(Format(l_administradora(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_administradora(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |             |             |             |             |             |"
    Mid(x_linha, 3, 40) = x_nome
    Mid(x_linha, 46, 20) = "Líquido             "
    i = Len(Format(l_liquido(3), "####,##0.00"))
    Mid(x_linha, 69 + 11 - i, i) = Format(l_liquido(3), "####,##0.00")
    i = Len(Format(l_liquido(2), "####,##0.00"))
    Mid(x_linha, 83 + 11 - i, i) = Format(l_liquido(2), "####,##0.00")
    i = Len(Format(l_liquido(1), "####,##0.00"))
    Mid(x_linha, 97 + 11 - i, i) = Format(l_liquido(1), "####,##0.00")
    i = Len(Format(l_liquido(4), "####,##0.00"))
    Mid(x_linha, 111 + 11 - i, i) = Format(l_liquido(4), "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 5
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
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| ANÁLISE DAS VENDAS COM CARTÕES DE CRÉDITO                       , __/__/____ |"
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
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+-------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                                  | MOVIMENTO            |  AMEX/SOLO  | CRED/DINERS |     VISA    |    TOTAL    |             |"
    x_linha = "|                                          |                      |   __  DIAS  |   __  DIAS  |   __  DIAS  |             |             |"
    Mid(x_linha, 71, 2) = Format(l_dias(3), "00")
    Mid(x_linha, 85, 2) = Format(l_dias(2), "00")
    Mid(x_linha, 99, 2) = Format(l_dias(1), "00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                          |                      |    ____%    |    ____%    |    ____%    |             |             |"
    i = Len(Format(l_percentual(3), "#0.00"))
    Mid(x_linha, 71 + 5 - i, i) = Format(l_percentual(3), "#0.00")
    i = Len(Format(l_percentual(2), "#0.00"))
    Mid(x_linha, 85 + 5 - i, i) = Format(l_percentual(2), "#0.00")
    i = Len(Format(l_percentual(1), "#0.00"))
    Mid(x_linha, 99 + 5 - i, i) = Format(l_percentual(1), "#0.00")
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_GotFocus()
    SendMessageLong cbo_periodo_i.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
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
    cbo_periodo_f.SetFocus
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
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_movimento_cartao_credito = bd_sgp.OpenTable("Movimento_Cartao_Credito")
    tbl_empresa.Index = "id_codigo"
    tbl_movimento_cartao_credito.Index = "id_data_emissao"
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
