VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_mapa_resumo 
   Caption         =   "Emissão do Mapa Resumo do ECF"
   ClientHeight    =   3300
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   8505
   Icon            =   "emissao_mapa_resumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_mapa_resumo.frx":030A
   ScaleHeight     =   3300
   ScaleWidth      =   8505
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2280
      Picture         =   "emissao_mapa_resumo.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime o Mapa Resumo do ECF (Correção)."
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5880
      Picture         =   "emissao_mapa_resumo.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2340
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtObservacao 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1680
         Width           =   8055
      End
      Begin VB.TextBox txtObservacao 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1320
         Width           =   8055
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   7680
         Picture         =   "emissao_mapa_resumo.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_mapa_resumo.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_mapa_resumo.frx":55A0
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
         Left            =   6600
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
         Caption         =   "&Observação"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   5580
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
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
Attribute VB_Name = "emissao_mapa_resumo"
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
Dim lIsentas As Currency
Dim lNaoIncidencia As Currency
Dim lSubstituicaoTributaria As Currency
Dim lICMS12 As Currency
Dim lICMS17 As Currency
Dim lImpostoDebitado As Currency

Private Empresa As New cEmpresa
Private rstMovimentoMapaResumo As New adodb.Recordset

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
End Sub
Private Sub ImpDet()
    Dim xImpostoDebitado As Currency
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    
    xR = 0
    xG = 0
    xB = 0
    
    With rstMovimentoMapaResumo
        If lLinhaI > 16 Then
            Printer.NewPage
            lLinhaI = CCur(ReadINI("MAPA RESUMO", "Margem Superior", gArquivoIni))
            ImpCab
        End If
        Printer.FontSize = 8
        Printer.FontBold = False
        lLinhaI = lLinhaI + 0.4
        
        
        Printer.Line (lColunaI + 0, lLinhaI)-(lColunaI + 0, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 0.8, lLinhaI)-(lColunaI + 0.8, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 1.6, lLinhaI)-(lColunaI + 1.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 2.4, lLinhaI)-(lColunaI + 2.4, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 3.5, lLinhaI)-(lColunaI + 3.5, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 4.4, lLinhaI)-(lColunaI + 4.4, lLinhaI + 0.5), RGB(xR, xG, xB)
        'Printer.Line (lColunaI + 5.1, lLinhaI)-(lColunaI + 5.1, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 6.7, lLinhaI)-(lColunaI + 6.7, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 8.6, lLinhaI)-(lColunaI + 8.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 10.6, lLinhaI)-(lColunaI + 10.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 12.6, lLinhaI)-(lColunaI + 12.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 14.6, lLinhaI)-(lColunaI + 14.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 16.4, lLinhaI)-(lColunaI + 16.4, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 18.2, lLinhaI)-(lColunaI + 18.2, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 20, lLinhaI)-(lColunaI + 20, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 21.8, lLinhaI)-(lColunaI + 21.8, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 23.6, lLinhaI)-(lColunaI + 23.6, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 25.4, lLinhaI)-(lColunaI + 25.4, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 26, lLinhaI)-(lColunaI + 26, lLinhaI + 0.5), RGB(xR, xG, xB)
        Printer.Line (lColunaI + 0, lLinhaI + 0.5)-(lColunaI + 26, lLinhaI + 0.5), RGB(xR, xG, xB)
        
        ImprimeCentralizado Format(Day(!Data), "00"), lColunaI + 0, lColunaI + 0.8, lLinhaI + 0.15, lLocal
        ImprimeCentralizado Format(![ECF Numero], "00"), lColunaI + 0.8, lColunaI + 1.6, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Contador de Reducoes Z], "#####0"), lColunaI + 1.6, lColunaI + 2.4, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Contagem de Operacao Final], "#####0"), lColunaI + 2.4, lColunaI + 3.5, lLinhaI + 0.15, lLocal
        
        
        
        ImprimeValor Format(![Totalizador Geral Final] - ![Totalizador Geral Inicial], "###,###,##0.00"), lColunaI + 6.7, lColunaI + 8.6, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Cancelamento de Item] + !Desconto - !Acrescimo, "###,###,##0.00"), lColunaI + 8.6, lColunaI + 10.6, lLinhaI + 0.15, lLocal
        ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 10.6, lColunaI + 12.6, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Valor Contabil], "###,###,##0.00"), lColunaI + 12.6, lColunaI + 14.6, lLinhaI + 0.15, lLocal
        ImprimeValor Format(!Isentas, "###,###,##0.00"), lColunaI + 14.6, lColunaI + 16.4, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Nao Incidencia], "###,###,##0.00"), lColunaI + 16.9, lColunaI + 18.2, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![Substituicao Tributaria], "###,###,##0.00"), lColunaI + 18.2, lColunaI + 20, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![ICMS 12], "###,###,##0.00"), lColunaI + 20, lColunaI + 21.8, lLinhaI + 0.15, lLocal
        ImprimeValor Format(![ICMS 17], "###,###,##0.00"), lColunaI + 21.8, lColunaI + 23.6, lLinhaI + 0.15, lLocal
        xImpostoDebitado = 0
        If ![ICMS 17] > 0 Then
            xImpostoDebitado = (![ICMS 12] * 12 / 100) + (![ICMS 17] * 17 / 100)
        End If
        ImprimeValor Format(xImpostoDebitado, "###,###,##0.00"), lColunaI + 23.6, lColunaI + 25.4, lLinhaI + 0.15, lLocal
        
        lVendaBruta = lVendaBruta + (![Totalizador Geral Final] - ![Totalizador Geral Inicial])
        lCancelamento = lCancelamento + ![Cancelamento de Item] + !Desconto - !Acrescimo
        lContabil = lContabil + ![Valor Contabil]
        lIsentas = lIsentas + !Isentas
        lNaoIncidencia = lNaoIncidencia + ![Nao Incidencia]
        lSubstituicaoTributaria = lSubstituicaoTributaria + ![Substituicao Tributaria]
        lICMS12 = lICMS12 + ![ICMS 12]
        lICMS17 = lICMS17 + ![ICMS 17]
        lImpostoDebitado = lImpostoDebitado + xImpostoDebitado
    End With
End Sub
Private Sub ImpCab()
    Dim xR As Integer
    Dim xG As Integer
    Dim xB As Integer
    
    On Error Resume Next
    
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
    ImprimeCentralizado "MAPA RESUMO E.C.F. (EQUIPAMENTO EMISSOR DE CUPOM FISCAL) - MR", lColunaI + 0, lColunaI + 17, lLinhaI + 0.05, lLocal
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
    ImprimeCentralizado "ECF", lColunaI + 0.8, lColunaI + 1.6, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "N.", lColunaI + 0.8, lColunaI + 1.6, lLinhaI + 2.6, lLocal
    ImprimeCentralizado "CONT", lColunaI + 1.6, lColunaI + 2.4, lLinhaI + 1.9, lLocal
    ImprimeCentralizado "RED.", lColunaI + 1.6, lColunaI + 2.4, lLinhaI + 2.3, lLocal
    ImprimeCentralizado "Z", lColunaI + 1.6, lColunaI + 2.4, lLinhaI + 2.7, lLocal
    ImprimeCentralizado "COO", lColunaI + 2.4, lColunaI + 3.5, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "FINAL", lColunaI + 2.4, lColunaI + 3.5, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "DOCUMENTO", lColunaI + 3.5, lColunaI + 6.7, lLinhaI + 1.8, lLocal
    ImprimeCentralizado "PRÉ-IMPRESSO", lColunaI + 3.5, lColunaI + 6.7, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "SER.", lColunaI + 3.5, lColunaI + 4.4, lLinhaI + 2.7, lLocal
    ImprimeCentralizado "N. ORDEM", lColunaI + 4.4, lColunaI + 6.7, lLinhaI + 2.7, lLocal
    
    ImprimeCentralizado "VENDA", lColunaI + 6.7, lColunaI + 8.6, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "BRUTA", lColunaI + 6.7, lColunaI + 8.6, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "DESCONTO E", lColunaI + 8.6, lColunaI + 10.6, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "CANCELAM.", lColunaI + 8.6, lColunaI + 10.6, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "TOTALIZ.", lColunaI + 10.6, lColunaI + 12.6, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "DE ISS", lColunaI + 10.6, lColunaI + 12.6, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "VALOR", lColunaI + 12.6, lColunaI + 14.6, lLinhaI + 2.2, lLocal
    ImprimeCentralizado "CONTÁBIL", lColunaI + 12.6, lColunaI + 14.6, lLinhaI + 2.6, lLocal
    
    ImprimeCentralizado "BASE DE CÁLCULO", lColunaI + 20, lColunaI + 26, lLinhaI + 2, lLocal
    ImprimeCentralizado "ISENTAS", lColunaI + 14.6, lColunaI + 16.4, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "NÃO", lColunaI + 16.4, lColunaI + 18.2, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "INCIDÊNCIA", lColunaI + 16.4, lColunaI + 18.2, lLinhaI + 2.85, lLocal
    ImprimeCentralizado "SUBSTITUIÇ.", lColunaI + 18.2, lColunaI + 20, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "TRIBUTÁRIA", lColunaI + 18.2, lColunaI + 20, lLinhaI + 2.85, lLocal
    
    ImprimeCentralizado "ICMS", lColunaI + 20, lColunaI + 21.8, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "12%", lColunaI + 20, lColunaI + 21.8, lLinhaI + 2.85, lLocal
    ImprimeCentralizado "ICMS", lColunaI + 21.8, lColunaI + 23.6, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "17%", lColunaI + 21.8, lColunaI + 23.6, lLinhaI + 2.85, lLocal
    ImprimeCentralizado "IMPOSTO", lColunaI + 23.6, lColunaI + 25.4, lLinhaI + 2.55, lLocal
    ImprimeCentralizado "DEBITADO", lColunaI + 23.6, lColunaI + 25.4, lLinhaI + 2.85, lLocal
    ImprimeCentralizado " ", lColunaI + 25.4, lColunaI + 26, lLinhaI + 2.55, lLocal
    
    
    Printer.Line (lColunaI + 11, lLinhaI + 1.2)-(lColunaI + 11, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 17, lLinhaI + 0)-(lColunaI + 17, lLinhaI + 0.6), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18, lLinhaI + 1.2)-(lColunaI + 18, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.5, lLinhaI + 0.6)-(lColunaI + 18.5, lLinhaI + 1.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 19.5, lLinhaI + 1.2)-(lColunaI + 19.5, lLinhaI + 1.8), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 22, lLinhaI + 0)-(lColunaI + 22, lLinhaI + 0.6), RGB(xR, xG, xB)
    
    
    'Linhas do Detalhe (Horizontais menores)
    Printer.Line (lColunaI + 3.5, lLinhaI + 2.5)-(lColunaI + 6.7, lLinhaI + 2.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 14.6, lLinhaI + 2.5)-(lColunaI + 26, lLinhaI + 2.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 0, lLinhaI + 3.2)-(lColunaI + 26, lLinhaI + 3.2), RGB(xR, xG, xB)
    
    
    'Linhas Verticais
    Printer.Line (lColunaI + 0.8, lLinhaI + 1.8)-(lColunaI + 0.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 1.6, lLinhaI + 1.8)-(lColunaI + 1.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 2.4, lLinhaI + 1.8)-(lColunaI + 2.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 3.5, lLinhaI + 1.8)-(lColunaI + 3.5, lLinhaI + 3.2), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 4.4, lLinhaI + 1.8)-(lColunaI + 4.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 4.4, lLinhaI + 2.5)-(lColunaI + 4.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 6.7, lLinhaI + 1.8)-(lColunaI + 6.7, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 8.6, lLinhaI + 1.8)-(lColunaI + 8.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 10.6, lLinhaI + 1.8)-(lColunaI + 10.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 12.6, lLinhaI + 1.8)-(lColunaI + 12.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 14.6, lLinhaI + 1.8)-(lColunaI + 14.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 16.4, lLinhaI + 2.5)-(lColunaI + 16.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.2, lLinhaI + 2.5)-(lColunaI + 18.2, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20, lLinhaI + 1.8)-(lColunaI + 20, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 21.8, lLinhaI + 2.5)-(lColunaI + 21.8, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 23.6, lLinhaI + 2.5)-(lColunaI + 23.6, lLinhaI + 3.2), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 25.4, lLinhaI + 2.5)-(lColunaI + 25.4, lLinhaI + 3.2), RGB(xR, xG, xB)
    
    'Printer.Line (lColunaI + 13, lLinhaI + 5.4)-(lColunaI + 13, lLinhaI + 8.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 19, lLinhaI + 7.2)-(lColunaI + 19, lLinhaI + 8.5), RGB(xR, xG, xB)
    
    Printer.FontSize = 12
    Printer.FontBold = False
    If Empresa.Nome Like "*POSTO DO RATINHO*" Then
    Else
        ImprimeCentralizado Format(Month(msk_data_i.Text), "00") & "/" & Format(Year(msk_data_i.Text), "0000"), lColunaI + 18.5, lColunaI + 22, lLinhaI + 0.1, lLocal
    End If
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
    ImprimeValor Format(lVendaBruta, "###,###,##0.00"), lColunaI + 6.7, lColunaI + 8.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lCancelamento, "###,###,##0.00"), lColunaI + 8.6, lColunaI + 10.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(0, "###,###,##0.00"), lColunaI + 10.6, lColunaI + 12.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lContabil, "###,###,##0.00"), lColunaI + 12.6, lColunaI + 14.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lIsentas, "###,###,##0.00"), lColunaI + 14.6, lColunaI + 16.4, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lNaoIncidencia, "###,###,##0.00"), lColunaI + 16.4, lColunaI + 18.2, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lSubstituicaoTributaria, "###,###,##0.00"), lColunaI + 18.2, lColunaI + 20, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lICMS12, "###,###,##0.00"), lColunaI + 20, lColunaI + 21.8, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lICMS17, "###,###,##0.00"), lColunaI + 21.8, lColunaI + 23.6, lLinhaI + 0.15, lLocal
    ImprimeValor Format(lImpostoDebitado, "###,###,##0.00"), lColunaI + 23.6, lColunaI + 25.4, lLinhaI + 0.15, lLocal

    Printer.FontBold = False
    
    'Printer.Line (lColunaI + 0.8, lLinhaI)-(lColunaI + 0.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 1.8, lLinhaI)-(lColunaI + 1.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 2.8, lLinhaI)-(lColunaI + 2.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    'Printer.Line (lColunaI + 3.5, lLinhaI)-(lColunaI + 3.5, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 6.7, lLinhaI)-(lColunaI + 6.7, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 8.6, lLinhaI)-(lColunaI + 8.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 10.6, lLinhaI)-(lColunaI + 10.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 12.6, lLinhaI)-(lColunaI + 12.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 14.6, lLinhaI)-(lColunaI + 14.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 16.4, lLinhaI)-(lColunaI + 16.4, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 18.2, lLinhaI)-(lColunaI + 18.2, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 20, lLinhaI)-(lColunaI + 20, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 21.8, lLinhaI)-(lColunaI + 21.8, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 23.6, lLinhaI)-(lColunaI + 23.6, lLinhaI + 0.5), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 25.4, lLinhaI)-(lColunaI + 25.4, lLinhaI + 0.5), RGB(xR, xG, xB)
    
    ImprimeCentralizado "OBSERVAÇÕES", lColunaI, lColunaI + 13, lLinhaI + 0.6, lLocal
    ImprimeCentralizado "RESPONSÁVEL PELO ESTABELECIMENTO", lColunaI + 13, lColunaI + 26, lLinhaI + 0.6, lLocal
    ImprimeTexto "NOME", lColunaI + 13.2, lColunaI + 26, lLinhaI + 1.1, lLocal
    ImprimeTexto "FUNÇÃO", lColunaI + 13.2, lColunaI + 19, lLinhaI + 2.1, lLocal
    ImprimeTexto "ASSINATURA", lColunaI + 19.2, lColunaI + 26, lLinhaI + 2.1, lLocal
    
    ImprimeTexto txtObservacao(0).Text, lColunaI + 0.2, lColunaI + 13, lLinhaI + 1.1, lLocal
    ImprimeTexto txtObservacao(1).Text, lColunaI + 0.2, lColunaI + 13, lLinhaI + 1.6, lLocal
    
    Printer.Line (lColunaI + 0, lLinhaI)-(lColunaI + 0, lLinhaI + 3), RGB(xR, xG, xB)
    Printer.Line (lColunaI + 13, lLinhaI + 0.5)-(lColunaI + 13, lLinhaI + 3), RGB(xR, xG, xB)
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
    lIsentas = 0
    lNaoIncidencia = 0
    lSubstituicaoTributaria = 0
    lICMS12 = 0
    lICMS17 = 0
    lImpostoDebitado = 0
End Sub
Private Sub LoopMapaResumo()
    lSQL = ""
    lSQL = lSQL & "SELECT Data, [Contador de Reducoes Z], [Contagem de Operacao Final], "
    lSQL = lSQL & "       [Totalizador Geral Inicial], [Totalizador Geral Final], [Cancelamento de Item], "
    lSQL = lSQL & "       Acrescimo, Desconto, [Valor Contabil], Isentas,  [Nao Incidencia], "
    lSQL = lSQL & "       [Substituicao Tributaria], [ICMS 12], [ICMS 17], [ECF Numero]"
    lSQL = lSQL & "  FROM Mapa_Resumo"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " ORDER BY Data, [ECF Numero]"
    Set rstMovimentoMapaResumo = Conectar.RsConexao(lSQL)
    If rstMovimentoMapaResumo.RecordCount > 0 Then
        Do Until rstMovimentoMapaResumo.EOF
            ImpDet
            rstMovimentoMapaResumo.MoveNext
        Loop
    End If
    rstMovimentoMapaResumo.Close
    Set rstMovimentoMapaResumo = Nothing
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    ImpCab
    LoopMapaResumo
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
    If ValidaCampos Then
        'If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        'End If
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
        txtObservacao(0).SetFocus
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
Private Sub txtObservacao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            txtObservacao(1).SetFocus
        Else
            cmd_imprimir.SetFocus
        End If
    End If
End Sub
