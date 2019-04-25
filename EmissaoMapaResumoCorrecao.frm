VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoMapaResumoCorrecao 
   Caption         =   "Emissão do Mapa Resumo do ECF (Correção)"
   ClientHeight    =   2295
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoMapaResumoCorrecao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoMapaResumoCorrecao.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1740
      Picture         =   "EmissaoMapaResumoCorrecao.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime o Mapa Resumo do ECF (Correção)."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4260
      Picture         =   "EmissaoMapaResumoCorrecao.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "EmissaoMapaResumoCorrecao.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoMapaResumoCorrecao.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoMapaResumoCorrecao.frx":55A0
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
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoMapaResumoCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lColuna(0 To 1) As Currency
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lLinhaTab As Currency
Dim lLocal As Integer
Dim lSQL As String
Dim lVendaBruta  As Currency
Dim lCancelamento As Currency
Dim lContabil As Currency
Dim lIsentasNaoTributadas As Currency
Dim lSubstituicaoTributaria As Currency
Dim lICMS17  As Currency
Dim lTotalCombustivel As Currency
Dim lTotalProduto As Currency
Dim lTotal As Currency

Private Empresa As New cEmpresa
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
Private MovimentoMapaResumo As New cMovimentoMapaResumo
Private MovimentoMapaResumoCorr As New cMovimentoMapaResumoCorr
Private rstMovimentoMapaResumo As New adodb.Recordset

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set MovimentoMapaResumo = Nothing
    Set MovimentoMapaResumoCorr = Nothing
End Sub
Private Sub ImpDet()
    Dim xValor As Currency
    Dim xValorCombustivel As Currency
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    
    xR = 0
    xG = 0
    xB = 0
    
    lSQL = ""
    lSQL = lSQL & "SELECT Data, [Contador de Reducoes Z], [Contagem de Operacao Final], "
    lSQL = lSQL & "       [Totalizador Geral Inicial], [Totalizador Geral Final], [Cancelamento de Item], "
    lSQL = lSQL & "       [Valor Contabil], [Isentas Nao Tributadas], [Substituicao Tributaria], "
    lSQL = lSQL & "       [ICMS 17], [Valor Combustivel], [Valor Produto]"
    lSQL = lSQL & "  FROM MapaResumoCorrecao"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " ORDER BY Data"
    Set rstMovimentoMapaResumo = Conectar.RsConexao(lSQL)
    With rstMovimentoMapaResumo
        If .RecordCount > 0 Then
            Do Until .EOF
                Printer.FontSize = 8
                Printer.FontBold = False
                lLinhaI = lLinhaI + 0.4
                
                
                Printer.Line (lColunaI + 0, lLinhaI)-(lColunaI + 0, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 0.8, lLinhaI)-(lColunaI + 0.8, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 1.8, lLinhaI)-(lColunaI + 1.8, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 2.8, lLinhaI)-(lColunaI + 2.8, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 3.5, lLinhaI)-(lColunaI + 3.5, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 5, lLinhaI)-(lColunaI + 5, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 7, lLinhaI)-(lColunaI + 7, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 9, lLinhaI)-(lColunaI + 9, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 11, lLinhaI)-(lColunaI + 11, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 13, lLinhaI)-(lColunaI + 13, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 14.8, lLinhaI)-(lColunaI + 14.8, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 16.6, lLinhaI)-(lColunaI + 16.6, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 18.4, lLinhaI)-(lColunaI + 18.4, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 20.2, lLinhaI)-(lColunaI + 20.2, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 22, lLinhaI)-(lColunaI + 22, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 23.8, lLinhaI)-(lColunaI + 23.8, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 26, lLinhaI)-(lColunaI + 26, lLinhaI + 0.5), RGB(xR, xG, xB)
                Printer.Line (lColunaI + 0, lLinhaI + 0.5)-(lColunaI + 26, lLinhaI + 0.5), RGB(xR, xG, xB)
                
                ImprimeCentralizado Format(Day(!Data), "00"), lColunaI + 0, lColunaI + 0.8, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Contador de Reducoes Z], "#####0"), lColunaI + 0.8, lColunaI + 1.8, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Contagem de Operacao Final], "#####0"), lColunaI + 1.8, lColunaI + 2.8, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Totalizador Geral Final] - ![Totalizador Geral Inicial], "###,###,##0.00"), lColunaI + 5, lColunaI + 7, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Cancelamento de Item], "###,###,##0.00"), lColunaI + 7, lColunaI + 9, lLinhaI + 0.15, lLocal
                ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 9, lColunaI + 11, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Valor Contabil], "###,###,##0.00"), lColunaI + 11, lColunaI + 13, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Isentas Nao Tributadas], "###,###,##0.00"), lColunaI + 13, lColunaI + 14.8, lLinhaI + 0.15, lLocal
                ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 14.8, lColunaI + 16.6, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Substituicao Tributaria], "###,###,##0.00"), lColunaI + 16.6, lColunaI + 18.4, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![ICMS 17], "###,###,##0.00"), lColunaI + 18.4, lColunaI + 20.2, lLinhaI + 0.15, lLocal
                ImprimeValor Format(![Valor Produto], "###,###,##0.00"), lColunaI + 20.2, lColunaI + 22, lLinhaI + 0.15, lLocal
                xValorCombustivel = ![Valor Combustivel] - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, !Data + 1, !Data + 1, "")
                ImprimeValor Format(xValorCombustivel, "###,###,##0.00"), lColunaI + 22, lColunaI + 23.8, lLinhaI + 0.15, lLocal
                xValor = ![Valor Produto] + xValorCombustivel
                ImprimeValor Format(xValor, "###,###,##0.00"), lColunaI + 23.8, lColunaI + 26, lLinhaI + 0.15, lLocal
                
                lVendaBruta = lVendaBruta + (![Totalizador Geral Final] - ![Totalizador Geral Inicial])
                lCancelamento = lCancelamento + ![Cancelamento de Item]
                lContabil = lContabil + ![Valor Contabil]
                lIsentasNaoTributadas = lIsentasNaoTributadas + ![Isentas Nao Tributadas]
                lSubstituicaoTributaria = lSubstituicaoTributaria + ![Substituicao Tributaria]
                lICMS17 = lICMS17 + ![ICMS 17]
                lTotalProduto = lTotalProduto + ![Valor Produto]
                lTotalCombustivel = lTotalCombustivel + xValorCombustivel
                lTotal = lTotal + xValor
                
                .MoveNext
            Loop
        End If
    
    End With
    rstMovimentoMapaResumo.Close
    Set rstMovimentoMapaResumo = Nothing
    
    
    
    
    
    
    
    'xValor = 0
    'If MovMapaResumo.ICMS17 > 0 Then
    '    xValor = MovMapaResumo.ICMS17 * 17 / 100
    'End If
    'ImprimeValor Format(xValor, "###,###,##0.00"), lColunaI + 22.65, lColunaI + 24.5, lLinhaI, lLocal
    
    'ImprimeValor Format(MovMapaResumo.TotalizadorGeralFinal, "###,###,##0.00"), lColunaI + 3.9, lColunaI + 6.6, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.TotalizadorGeralInicial, "###,###,##0.00"), lColunaI + 6.6, lColunaI + 9.3, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.CancelamentoItem, "###,###,##0.00"), lColunaI + 9.3, lColunaI + 11.3, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.ValorContabil, "###,###,##0.00"), lColunaI + 11.3, lColunaI + 13.5, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.IsentasNaoTributadas, "###,###,##0.00"), lColunaI + 13.5, lColunaI + 15.3, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.SubstituicaoTributaria, "###,###,##0.00"), lColunaI + 15.3, lColunaI + 17.18, lLinhaI + 4.95, lLocal
    'ImprimeValor Format(MovMapaResumo.ICMS17, "###,###,##0.00"), lColunaI + 17.18, lColunaI + 19, lLinhaI + 4.95, lLocal
    
    'ImprimeTexto MovMapaResumo.Observacao1, lColunaI + 0, lColunaI + 13, lLinhaI + 6.5, lLocal
    'ImprimeTexto MovMapaResumo.Observacao2, lColunaI + 0, lColunaI + 13, lLinhaI + 7.2, lLocal
    
End Sub
Private Sub ImpCab()
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    
    'seleciona medidas para centímetros
    Printer.ColorMode = 2
    Printer.ScaleMode = 7
    Printer.PrintQuality = vbPRPQLow
    Printer.PaperSize = vbPRPSLegal
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
    
    Empresa.LocalizarCodigo (g_empresa)
    
    'Bordas Externas
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 26, lLinhaI + 0), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 0, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 26, lLinhaI + 0)-(lColunaI + 26, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 3.2)-(lColunaI + 26, lLinhaI + 3.2), RGB(xR, xG, xB)
    
    'Linhas Horizontais do Cabecalho
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeCentralizado "MAPA RESUMO E.C.F.    ( C O R R E Ç Ã O )                    ", lColunaI + 0, lColunaI + 17, lLinhaI + 0.05, lLocal
    Printer.FontBold = False
    Printer.FontSize = 8
    ImprimeTexto "NÚMERO", lColunaI + 17 + 0.2, lColunaI + 22, lLinhaI + 0.15, lLocal
    ImprimeTexto "DATA", lColunaI + 22 + 0.2, lColunaI + 26, lLinhaI + 0.15, lLocal
    
    ImprimeTexto "NOME", lColunaI + 0.2, lColunaI + 20, lLinhaI + 0.75, lLocal
    ImprimeTexto "INSCRIÇÃO ESTADUAL", lColunaI + 18.5 + 0.2, lColunaI + 26, lLinhaI + 0.75, lLocal
    ImprimeTexto "ENDEREÇO", lColunaI + 0.2, lColunaI + 11, lLinhaI + 1.35, lLocal
    ImprimeTexto "MUNICÍPIO", lColunaI + 11 + 0.2, lColunaI + 19, lLinhaI + 1.35, lLocal
    ImprimeTexto "UF", lColunaI + 18 + 0.2, lColunaI + 20.1, lLinhaI + 1.35, lLocal
    ImprimeTexto "CNPJ", lColunaI + 19.5 + 0.2, lColunaI + 26, lLinhaI + 1.35, lLocal
    Printer.DrawWidth = 1
    Printer.Line (lColunaI + 0, lLinhaI + 0.6)-(lColunaI + 26, lLinhaI + 0.6), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 1.2)-(lColunaI + 26, lLinhaI + 1.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 1.8)-(lColunaI + 26, lLinhaI + 1.8), RGB(xR, xG, xB)
    
    'Linhas Verticais do Cabecalho
    ImprimeCentralizado "DIA", lColunaI, lColunaI + 0.8, lLinhaI + 2.3, lLocal
    ImprimeCentralizado "CONT.", lColunaI + 0.8, lColunaI + 1.8, lLinhaI + 1.9, lLocal
    ImprimeCentralizado "RED.", lColunaI + 0.8, lColunaI + 1.8, lLinhaI + 2.3, lLocal
    ImprimeCentralizado "Z", lColunaI + 0.8, lColunaI + 1.8, lLinhaI + 2.7, lLocal
    ImprimeCentralizado "COO", lColunaI + 1.8, lColunaI + 2.8, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "FINAL", lColunaI + 1.8, lColunaI + 2.8, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "DOCUMENTO", lColunaI + 2.8, lColunaI + 5, lLinhaI + 1.8, lLocal
    ImprimeCentralizado "PRÉ-IMPRESSO", lColunaI + 2.8, lColunaI + 5, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "SER.", lColunaI + 2.8, lColunaI + 3.5, lLinhaI + 2.7, lLocal
    ImprimeCentralizado "N. ORDEM", lColunaI + 3.5, lColunaI + 5, lLinhaI + 2.7, lLocal
    
    ImprimeCentralizado "VENDA", lColunaI + 5, lColunaI + 7, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "BRUTA", lColunaI + 5, lColunaI + 7, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "DESCONTO E", lColunaI + 7, lColunaI + 9, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "CANCELAM.", lColunaI + 7, lColunaI + 9, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "TOTALIZ.", lColunaI + 9, lColunaI + 11, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "DE ISS", lColunaI + 9, lColunaI + 11, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "VALOR", lColunaI + 11, lColunaI + 13, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "CONTÁBIL", lColunaI + 11, lColunaI + 13, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "BASE DE CÁLCULO", lColunaI + 13, lColunaI + 20.2, lLinhaI + 2, lLocal
    ImprimeCentralizado "VENDA CONCILIADA", lColunaI + 20.2, lColunaI + 26, lLinhaI + 2, lLocal
    
    ImprimeCentralizado "ISENTAS", lColunaI + 13, lColunaI + 14.8, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "NÃO TRIB.", lColunaI + 14.8, lColunaI + 16.6, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "S.T.", lColunaI + 16.6, lColunaI + 18.4, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "17%", lColunaI + 18.4, lColunaI + 20.2, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "Produto", lColunaI + 20.2, lColunaI + 22, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "Combustível", lColunaI + 22, lColunaI + 23.8, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "Total", lColunaI + 23.8, lColunaI + 26, lLinhaI + 2.55, lLocal
    
    ImprimeCentralizado "E.C.F.", lColunaI + 20.2, lColunaI + 22, lLinhaI + 2.85, lLocal
    ImprimeCentralizado "L.M.C.", lColunaI + 22, lColunaI + 23.8, lLinhaI + 2.85, lLocal
    ImprimeCentralizado "Conciliado", lColunaI + 23.8, lColunaI + 26, lLinhaI + 2.85, lLocal
    
    
    Printer.Line (lColunaI + 11, lLinhaI + 1.2)-(lColunaI + 11, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 17, lLinhaI + 0)-(lColunaI + 17, lLinhaI + 0.6), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18, lLinhaI + 1.2)-(lColunaI + 18, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.5, lLinhaI + 0.6)-(lColunaI + 18.5, lLinhaI + 1.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 19.5, lLinhaI + 1.2)-(lColunaI + 19.5, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22, lLinhaI + 0)-(lColunaI + 22, lLinhaI + 0.6), RGB(xR, xG, xB)
    
    
    'Linhas do Detalhe
    Printer.Line (lColunaI + 2.8, lLinhaI + 2.5)-(lColunaI + 5, lLinhaI + 2.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI + 2.5)-(lColunaI + 26, lLinhaI + 2.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 3.2)-(lColunaI + 26, lLinhaI + 3.2), RGB(xR, xG, xB)
    
    
    Printer.Line (lColunaI + 0.8, lLinhaI + 1.8)-(lColunaI + 0.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 1.8, lLinhaI + 1.8)-(lColunaI + 1.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 2.8, lLinhaI + 1.8)-(lColunaI + 2.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 3.5, lLinhaI + 2.5)-(lColunaI + 3.5, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 5, lLinhaI + 1.8)-(lColunaI + 5, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 7, lLinhaI + 1.8)-(lColunaI + 7, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 9, lLinhaI + 1.8)-(lColunaI + 9, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 11, lLinhaI + 1.8)-(lColunaI + 11, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI + 1.8)-(lColunaI + 13, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 14.8, lLinhaI + 2.5)-(lColunaI + 14.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 16.6, lLinhaI + 2.5)-(lColunaI + 16.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.4, lLinhaI + 2.5)-(lColunaI + 18.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20.2, lLinhaI + 1.8)-(lColunaI + 20.2, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22, lLinhaI + 2.5)-(lColunaI + 22, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 23.8, lLinhaI + 2.5)-(lColunaI + 23.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    
    'Printer.Line (lColunaI + 13, lLinhaI + 5.4)-(lColunaI + 13, lLinhaI + 8.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 19, lLinhaI + 7.2)-(lColunaI + 19, lLinhaI + 8.5), RGB(xR, xG, xB)
    
    Printer.FontSize = 12
    Printer.FontBold = False
    ImprimeCentralizado Format(Month(msk_data_i.Text), "0000") & "/" & Format(Year(msk_data_i.Text), "0000"), lColunaI + 18.5, lColunaI + 22, lLinhaI + 0.1, lLocal
    ImprimeCentralizado msk_data_f.Text, lColunaI + 23, lColunaI + 26, lLinhaI + 0.1, lLocal
    
    ImprimeTexto Empresa.Nome, lColunaI + 1.5, lColunaI + 20, lLinhaI + 0.7, lLocal
    ImprimeTexto Empresa.InscricaoEstadual, lColunaI + 22, lColunaI + 26, lLinhaI + 0.7, lLocal
    
    ImprimeTexto Trim(Empresa.Endereco) & ", " & Trim(Empresa.Bairro), lColunaI + 2, lColunaI + 11, lLinhaI + 1.3, lLocal
    ImprimeTexto Empresa.Cidade, lColunaI + 13, lColunaI + 19, lLinhaI + 1.3, lLocal
    ImprimeCentralizado Empresa.Estado, lColunaI + 18.5, lColunaI + 19.5, lLinhaI + 1.3, lLocal
    ImprimeTexto fMascaraCNPJ(Empresa.CGC), lColunaI + 21, lColunaI + 26, lLinhaI + 1.3, lLocal
    
    lLinhaI = 3.65
    
    'Printer.DrawWidth = 1
    'Printer.EndDoc
End Sub
Private Sub ImpTotal()
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    
    xR = 0
    xG = 0
    xB = 0
    
    lLinhaI = lLinhaI + 0.4
    
    Printer.FontBold = True
    
    ImprimeCentralizado "** TOTAIS DO DIA", lColunaI, lColunaI + 4, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lVendaBruta, "###,###,##0.00"), lColunaI + 5, lColunaI + 7, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lCancelamento, "###,###,##0.00"), lColunaI + 7, lColunaI + 9, lLinhaI + 0.15, lLocal
    ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 9, lColunaI + 11, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lContabil, "###,###,##0.00"), lColunaI + 11, lColunaI + 13, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lIsentasNaoTributadas, "###,###,##0.00"), lColunaI + 13, lColunaI + 14.8, lLinhaI + 0.15, lLocal
    ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 14.8, lColunaI + 16.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lSubstituicaoTributaria, "###,###,##0.00"), lColunaI + 16.6, lColunaI + 18.4, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lICMS17, "###,###,##0.00"), lColunaI + 18.4, lColunaI + 20.2, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lTotalProduto, "###,###,##0.00"), lColunaI + 20.2, lColunaI + 22, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lTotalCombustivel, "###,###,##0.00"), lColunaI + 22, lColunaI + 23.8, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lTotal, "###,###,##0.00"), lColunaI + 23.8, lColunaI + 26, lLinhaI + 0.15, lLocal
    
    Printer.FontBold = False
    
    'Printer.Line (lColunaI + 0.8, lLinhaI)-(lColunaI + 0.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 1.8, lLinhaI)-(lColunaI + 1.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 2.8, lLinhaI)-(lColunaI + 2.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 3.5, lLinhaI)-(lColunaI + 3.5, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 5, lLinhaI)-(lColunaI + 5, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 7, lLinhaI)-(lColunaI + 7, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 9, lLinhaI)-(lColunaI + 9, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 11, lLinhaI)-(lColunaI + 11, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI)-(lColunaI + 13, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 14.8, lLinhaI)-(lColunaI + 14.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 16.6, lLinhaI)-(lColunaI + 16.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.4, lLinhaI)-(lColunaI + 18.4, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20.2, lLinhaI)-(lColunaI + 20.2, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22, lLinhaI)-(lColunaI + 22, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 23.8, lLinhaI)-(lColunaI + 23.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    
    ImprimeCentralizado "OBSERVAÇÕES", lColunaI, lColunaI + 13, lLinhaI + 0.6, lLocal
    ImprimeCentralizado "RESPONSÁVEL PELO ESTABELECIMENTO", lColunaI + 13, lColunaI + 26, lLinhaI + 0.6, lLocal
    ImprimeTexto "NOME", lColunaI + 13.2, lColunaI + 26, lLinhaI + 1.1, lLocal
    ImprimeTexto "FUNÇÃO", lColunaI + 13.2, lColunaI + 19, lLinhaI + 2.1, lLocal
    ImprimeTexto "ASSINATURA", lColunaI + 19.2, lColunaI + 26, lLinhaI + 2.1, lLocal
    
    Printer.Line (lColunaI + 0, lLinhaI)-(lColunaI + 0, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI)-(lColunaI + 13, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 19, lLinhaI + 2)-(lColunaI + 19, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 26, lLinhaI)-(lColunaI + 26, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 3)-(lColunaI + 26, lLinhaI + 3), RGB(xR, xG, xB)
    
    'Printer.Line (lColunaI + 0, lLinhaI)-(lColunaI + 26, lLinhaI), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 0.5)-(lColunaI + 26, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 1)-(lColunaI + 26, lLinhaI + 1), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI + 2)-(lColunaI + 26, lLinhaI + 2), RGB(xR, xG, xB)
    Printer.EndDoc

End Sub
Private Sub ZeraVariaveis()
    lLocal = 1
    lLinhaI = CCur(ReadINI("MAPA RESUMO", "Margem Superior", gArquivoIni))
    lColunaI = CCur(ReadINI("MAPA RESUMO", "Margem Esquerda", gArquivoIni))
    
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 20
    lLinhaTab = 0
    lVendaBruta = 0
    lCancelamento = 0
    lContabil = 0
    lIsentasNaoTributadas = 0
    lSubstituicaoTributaria = 0
    lICMS17 = 0
    lTotalCombustivel = 0
    lTotalProduto = 0
    lTotal = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    ImpCab
    ImpDet
    ImpTotal
    cmd_sair.SetFocus
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_imprimir.SetFocus
    g_string = " "
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    'Converte Mapa_Resumo para MapaResumoCorrecao
    'ConverteMapaResumo
    If ValidaCampos Then
        'If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        'End If
    End If
End Sub
Private Sub ConverteMapaResumo()
    Dim xContinua As Boolean
    
    Exit Sub
    xContinua = True
    If MovimentoMapaResumo.LocalizarUltimo(g_empresa) Then
        If MovimentoMapaResumo.LocalizarPrimeiro Then
            Do Until xContinua = False
                MovimentoMapaResumoCorr.Empresa = MovimentoMapaResumo.Empresa
                MovimentoMapaResumoCorr.Data = MovimentoMapaResumo.Data
                MovimentoMapaResumoCorr.numero = MovimentoMapaResumo.numero
                MovimentoMapaResumoCorr.ECFNumero = MovimentoMapaResumo.ECFNumero
                MovimentoMapaResumoCorr.ContagemOperacaoInicial = MovimentoMapaResumo.ContagemOperacaoInicial
                MovimentoMapaResumoCorr.ContagemOperacaoFinal = MovimentoMapaResumo.ContagemOperacaoFinal
                MovimentoMapaResumoCorr.TotalizadorGeralFinal = MovimentoMapaResumo.TotalizadorGeralFinal
                MovimentoMapaResumoCorr.TotalizadorGeralInicial = MovimentoMapaResumo.TotalizadorGeralInicial
                MovimentoMapaResumoCorr.CancelamentoItem = MovimentoMapaResumo.CancelamentoItem
                MovimentoMapaResumoCorr.ValorContabil = MovimentoMapaResumo.ValorContabil
                MovimentoMapaResumoCorr.IsentasNaoTributadas = MovimentoMapaResumo.Isentas
                MovimentoMapaResumoCorr.SubstituicaoTributaria = MovimentoMapaResumo.SubstituicaoTributaria
                MovimentoMapaResumoCorr.ICMS17 = MovimentoMapaResumo.ICMS17
                MovimentoMapaResumoCorr.ContadorReducoesZ = MovimentoMapaResumo.ContadorReducoesZ
                MovimentoMapaResumoCorr.Observacao1 = MovimentoMapaResumo.Observacao1
                MovimentoMapaResumoCorr.Observacao2 = MovimentoMapaResumo.Observacao2
                MovimentoMapaResumoCorr.ValorCombustivel = MovimentoBomba.ValorVendaPeriodo(g_empresa, MovimentoMapaResumo.Data, MovimentoMapaResumo.Data, "", 1, 9)
                MovimentoMapaResumoCorr.ValorCombustivel = MovimentoMapaResumoCorr.ValorCombustivel - MovimentoAfericao.TotalPeriodo(g_empresa, MovimentoMapaResumo.Data, 0, False)
                MovimentoMapaResumoCorr.ValorProduto = MovimentoCupomFiscal.ValorProdutoVendaData(g_empresa, 0, MovimentoMapaResumo.Data, MovimentoMapaResumo.Data, 1, 9, 0)
                If Not MovimentoMapaResumoCorr.Incluir Then
                    MsgBox "Erro ao incluir registro MapaResumoCorrecao", vbCritical, "Erro de Integridade"
                End If
                If Not MovimentoMapaResumo.LocalizarProximo Then
                    xContinua = False
                End If
            Loop
        End If
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
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
    'TestaImpressora
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
    MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
    MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
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
        cmd_imprimir.SetFocus
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
