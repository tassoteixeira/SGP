VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_impressao_duplicata 
   Caption         =   "Impressão de Duplicatas"
   ClientHeight    =   5745
   ClientLeft      =   3600
   ClientTop       =   3375
   ClientWidth     =   9330
   Icon            =   "lst_impressao_duplicata.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_impressao_duplicata.frx":030A
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1860
      Picture         =   "lst_impressao_duplicata.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Visualiza duplicatas à receber (geral)."
      Top             =   4800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   4320
      Picture         =   "lst_impressao_duplicata.frx":162A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Imprime duplicatas à receber (geral)."
      Top             =   4800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6780
      Picture         =   "lst_impressao_duplicata.frx":2904
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9135
      Begin VB.TextBox txtNumeroDuplicata 
         Height          =   285
         Left            =   6780
         MaxLength       =   11
         TabIndex        =   36
         Top             =   4260
         Width           =   1095
      End
      Begin VB.TextBox txtValorDuplicata 
         Height          =   285
         Left            =   2580
         MaxLength       =   10
         TabIndex        =   34
         Top             =   4260
         Width           =   1095
      End
      Begin VB.TextBox txtNumeroFatura 
         Height          =   285
         Left            =   6780
         MaxLength       =   11
         TabIndex        =   32
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txtValorFatura 
         Height          =   285
         Left            =   2580
         MaxLength       =   10
         TabIndex        =   30
         Top             =   3900
         Width           =   1095
      End
      Begin VB.CommandButton cmdDataVencimento 
         Height          =   315
         Left            =   7920
         Picture         =   "lst_impressao_duplicata.frx":3BDE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox txtPracaPagamento 
         Height          =   285
         Left            =   2580
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1980
         Width           =   2775
      End
      Begin VB.TextBox txtLocalCobranca 
         Height          =   285
         Left            =   2580
         MaxLength       =   40
         TabIndex        =   16
         Top             =   2340
         Width           =   4935
      End
      Begin VB.TextBox txtRazaoSocial 
         Height          =   285
         Left            =   2580
         MaxLength       =   40
         TabIndex        =   2
         Top             =   180
         Width           =   4935
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   2580
         MaxLength       =   60
         TabIndex        =   4
         Top             =   540
         Width           =   6435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   2580
         MaxLength       =   20
         TabIndex        =   6
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   2580
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1620
         Width           =   1455
      End
      Begin VB.TextBox txtInscricaoEstadual 
         Height          =   285
         Left            =   6780
         MaxLength       =   20
         TabIndex        =   22
         Top             =   3060
         Width           =   1995
      End
      Begin VB.CommandButton cmdDataEmissao 
         Height          =   315
         Left            =   3720
         Picture         =   "lst_impressao_duplicata.frx":4EB8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   3540
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskDataEmissao 
         Height          =   300
         Left            =   2580
         TabIndex        =   24
         Top             =   3540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCpf 
         Height          =   300
         Left            =   2580
         TabIndex        =   18
         Top             =   2700
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Format          =   "___.___.___-__"
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCnpj 
         Height          =   300
         Left            =   2580
         TabIndex        =   20
         Top             =   3060
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         Format          =   "__.___.___/____-__"
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   300
         Left            =   6780
         TabIndex        =   10
         Top             =   1260
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "__.___-___"
         Mask            =   "##.###-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataVencimento 
         Height          =   300
         Left            =   6780
         TabIndex        =   27
         Top             =   3540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da Duplicata"
         Height          =   300
         Index           =   3
         Left            =   4920
         TabIndex        =   35
         Top             =   4260
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor da Duplicata"
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   4260
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da Fatura"
         Height          =   300
         Index           =   0
         Left            =   4920
         TabIndex        =   31
         Top             =   3900
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor da Fatura"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   3900
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Data do Vencimento"
         Height          =   300
         Left            =   4920
         TabIndex        =   26
         Top             =   3540
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9120
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Label Label12 
         Caption         =   "Praça de pagamento"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   1980
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Endereço de Cobrança"
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   2340
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "C.E.P."
         Height          =   300
         Left            =   4920
         TabIndex        =   9
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Razão social"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Estado"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Endereço"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Cidade"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Telefone "
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "C.P.F."
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   2700
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "CNPJ"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Inscrição Estadual"
         Height          =   300
         Left            =   4920
         TabIndex        =   21
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   3540
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Imagem 
      Height          =   1920
      Left            =   2940
      Picture         =   "lst_impressao_duplicata.frx":6192
      Top             =   4740
      Visible         =   0   'False
      Width           =   6390
   End
End
Attribute VB_Name = "lst_impressao_duplicata"
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

Dim lLinha2 As Currency
Dim lMargemLin As Currency
Dim lMargemCol As Currency
Dim lCabLin As Currency
Dim lLinMax As Currency
Dim lColMax As Currency
'Fim de variáveis padrão para relatório

Private Cliente As New cCliente
Private Empresa As New cEmpresa
Private Sub AtivaBotoes()
    cmd_visualizar.Enabled = True
    cmd_imprimir.Enabled = True
    cmd_sair.Enabled = True
End Sub
Private Sub DesativaBotoes()
    cmd_visualizar.Enabled = False
    cmd_imprimir.Enabled = False
    cmd_sair.Enabled = False
End Sub
Private Sub Finaliza()
    Set Cliente = Nothing
    Set Empresa = Nothing
End Sub
Private Sub ImpCabEpson()
    Dim x_linha As String
    Dim i As Integer
    'Dim x_numero As Currency
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 256"
        'BioImprime "@@Printer.Width = 2286"
        'BioImprime "@@Printer.Height = 1524"
        'Printer.PaperSize = 256
        'Printer.ScaleMode = 7
        'x_numero = Printer.Height
        'x_numero = Printer.Width
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    Else
        BioImprime "@@Printer.NewPage"
    End If
    lPagina = lPagina + 1
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.FontBold = True"
    
    If Empresa.LocalizarCodigo(g_empresa) Then
        x_linha = "                                                                                "
        BioImprime "@Printer.Print " & x_linha
        
        x_linha = "                                                                                "
        Mid(x_linha, 39, 40) = Empresa.Nome
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                                                                                "
        Mid(x_linha, 39, 40) = Empresa.Endereco
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                                                                                "
        Mid(x_linha, 39, 30) = Empresa.Cidade
        Mid(x_linha, 70, 6) = "UF: " & Empresa.Estado
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                                                                                "
        BioImprime "@Printer.Print " & x_linha
        Mid(x_linha, 39, 50) = "CGC: " & fMascaraCNPJ(Trim(Empresa.CGC)) & "  IE: " & Trim(Empresa.InscricaoEstadual)
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                                                                                "
        Mid(x_linha, 39, 40) = "EMITIDA EM: " & mskDataEmissao.Text
        BioImprime "@Printer.Print " & x_linha
    End If

    
    'BioImprime "@@Printer.FontBold = True"
    'x_linha = "|                                                                  PAGINA: ___ |"
    'Mid(x_linha, 3, 40) = g_nome_empresa
    'Mid(x_linha, 76, 3) = Format(lPagina, "000")
    'BioImprime "@Printer.Print " & x_linha
    'BioImprime "@@Printer.FontBold = False"
    'x_linha = "| RELACAO DAS DUPLICATAS A RECEBER (ANALITICO)              CIDADE, __/__/____ |"
    '
    'i = Len(Trim(g_cidade_empresa))
    'Mid(x_linha, 37 + 30 - i, i) = Trim(g_cidade_empresa)
    'If chk_sintetico Then
    '    Mid(x_linha, 37, 9) = "SINTETICO"
    'End If
    'Mid(x_linha, 69, 10) = mskDataEmissao
    'BioImprime "@Printer.Print " & x_linha
    'x_linha = "| TIPO DE DOCUMENTO.:                                                          |"
    'Mid(x_linha, 23, 25) = cbo_tipo_documento
    'BioImprime "@Printer.Print " & x_linha
    'x_linha = "| REFERENTE A.......: __/__/____ A __/__/____                                  |"
    'Mid(x_linha, 23, 10) = mskDataEmissao_i
    'Mid(x_linha, 36, 10) = mskDataEmissao_f
    'BioImprime "@Printer.Print " & x_linha
    'If chk_sintetico Then
    '    BioImprime "@@Printer.FontName = Draft 17cpi"
    '    x_linha = "+----------+------------------------------------------+------------------------------------------+----+----------------+----------------+"
    '    BioImprime "@Printer.Print " & x_linha
    '    x_linha = "|  CODIGO  | NOME DO CLIENTE                          | ENDERECO                                 | UF |    TELEFONE    |    V A L O R   |"
    '    BioImprime "@Printer.Print " & x_linha
    '    x_linha = "+----------+------------------------------------------+------------------------------------------+----+----------------+----------------+"
    '    BioImprime "@Printer.Print " & x_linha
    'End If
End Sub
Private Sub ImpCabHP()
    Dim x_linha As String
    Dim i As Integer
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        'Printer.PaperSize = 1
        Printer.FontName = "Arial"
        Printer.FontBold = False
    End If
    lPagina = lPagina + 1
    Printer.CurrentY = 0
    
    
    'Seleciona tamanho da fonte
    Printer.FontSize = 34
    'Printer.Line (0, 0)-(0, 26)
    Printer.DrawWidth = 4
    Printer.Line (lMargemCol, lMargemLin)-(lColMax, lLinMax), RGB(0, 0, 0), B
    
    Printer.Line (lMargemCol, lMargemLin + 3.5)-(lColMax, lMargemLin + 3.5), RGB(0, 0, 0)
    Printer.Line (lMargemCol + 11.5, lMargemLin)-(lMargemCol + 11.5, lMargemLin + 3.5), RGB(0, 0, 0)
    
    If Empresa.LocalizarCodigo(g_empresa) Then
        Printer.FontSize = 22
        Printer.DrawWidth = 6
        Printer.FontBold = True
        Printer.ForeColor = RGB(0, 0, 0)
        ImprimeCentralizado "M A S T E R P U M A S", lMargemCol + 3.2, lMargemCol + 10.5, lMargemLin, lLocal
        Printer.FontSize = 10
        Printer.FontBold = False
        ImprimeCentralizado Empresa.Nome, lMargemCol + 3.2, lMargemCol + 10.5, lMargemLin + 0.9, lLocal
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.FontSize = 9
        'ImprimeCentralizado "Metalassê - Colchôes - Espumas - Tule", lMargemCol + 3.2, lMargemCol + 10.5, lMargemLin + 1.4, lLocal
        'ImprimeCentralizado "Tecidos - Plásticos - Filó", lMargemCol + 3.2, lMargemCol + 10.5, lMargemLin + 1.9, lLocal
        Printer.FontSize = 10
        Printer.FontBold = True
        ImprimeCentralizado "FONES: (0xx62) 271-1202, (0xx62) 295-8932, FAX: (0xx62) 271-4420", lMargemCol, lMargemCol + 11.5, lMargemLin + 2.5, lLocal
        Printer.FontBold = False
        'ImprimeCentralizado Trim(Empresa.Endereco) & " - " & Trim(Empresa.Bairro) & " - " & Trim(Empresa.Cidade) & " - " & Trim(Empresa.Estado), lMargemCol, lMargemCol + 11.5, lMargemLin + 3, lLocal
        Imagem.Picture = LoadPicture("\vb5\sgp\icons\Duplicata.bmp")
        Printer.ScaleMode = 1
        Printer.PaintPicture Imagem.Picture, 350, 320, Imagem.Width, Imagem.Height
        Printer.ScaleMode = 7
        
        
        Printer.FontSize = 9
        ''ImprimeString Trim(Empresa.Endereco) & " - " & Trim(Empresa.Bairro), lMargemCol + 11.5 + 1, lMargemLin + 0.2, lLocal
        ''ImprimeString "CEP: " & Format(Mid(Empresa.CEP, 1, 5), "00,000") & "-" & Format(Mid(Empresa.CEP, 6, 3), "000") & " - " & Trim(Empresa.Cidade) & " - " & Trim(Empresa.Estado), lMargemCol + 11.5 + 1, lMargemLin + 0.7, lLocal
        Printer.FontSize = 12
        ImprimeString "CNPJ(MF): " & fMascaraCNPJ(Empresa.CGC), lMargemCol + 11.5 + 1, lMargemLin + 1.5, lLocal
        ImprimeString "Inscrição Estadual: " & Empresa.InscricaoEstadual, lMargemCol + 11.5 + 1, lMargemLin + 2, lLocal
        Printer.FontSize = 10
        ImprimeString "Data de Emissão ", lMargemCol + 11.5 + 1, lMargemLin + 3, lLocal
        Printer.FontSize = 12
        ImprimeString mskDataEmissao, lMargemCol + 11.5 + 4, lMargemLin + 3, lLocal
        
        
        Printer.Line (lMargemCol, lMargemLin + 3.9)-(lMargemCol + 13, lMargemLin + 3.9), RGB(0, 0, 0)
        Printer.Line (lMargemCol, lMargemLin + 4.3)-(lMargemCol + 16.5, lMargemLin + 4.3), RGB(0, 0, 0)
        Printer.Line (lMargemCol, lMargemLin + 5.1)-(lMargemCol + 16.5, lMargemLin + 5.1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 3.6, lMargemLin + 3.9)-(lMargemCol + 3.6, lMargemLin + 5.1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 6.5, lMargemLin + 3.5)-(lMargemCol + 6.5, lMargemLin + 5.1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 10.5, lMargemLin + 3.9)-(lMargemCol + 10.5, lMargemLin + 5.1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 13, lMargemLin + 3.5)-(lMargemCol + 13, lMargemLin + 5.1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 16.5, lMargemLin + 3.5)-(lMargemCol + 16.5, lMargemLin + 6.6), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 16.5, lMargemLin + 6.6)-(lColMax, lMargemLin + 6.6), RGB(0, 0, 0)
        Printer.FontSize = 8
        ImprimeCentralizado "FATURA", lMargemCol, lMargemCol + 6.5, lMargemLin + 3.55, lLocal
        ImprimeCentralizado "DUPLICATA", lMargemCol + 6.5, lMargemCol + 13, lMargemLin + 3.55, lLocal
        Printer.FontSize = 11
        ImprimeCentralizado "VENCIMENTO", lMargemCol + 13, lMargemCol + 16.5, lMargemLin + 3.7, lLocal
        Printer.FontSize = 8
        ImprimeCentralizado "VALOR R$", lMargemCol, lMargemCol + 3.6, lMargemLin + 3.95, lLocal
        ImprimeCentralizado "NÚMERO", lMargemCol + 3.6, lMargemCol + 6.5, lMargemLin + 3.95, lLocal
        ImprimeCentralizado "VALOR R$", lMargemCol + 6.5, lMargemCol + 10.5, lMargemLin + 3.95, lLocal
        ImprimeCentralizado "N. DE ORDEM", lMargemCol + 10.5, lMargemCol + 13, lMargemLin + 3.95, lLocal
        ImprimeCentralizado "Para uso da", lMargemCol + 16.5, lMargemCol + lColMax, lMargemLin + 3.55, lLocal
        ImprimeCentralizado "Instituição Financeira", lMargemCol + 16.5, lMargemCol + lColMax, lMargemLin + 4.05, lLocal
    
        ImprimeString "DESCONTO DE:", lMargemCol + 5.3, lMargemLin + 5.3, lLocal
        ImprimeString "% SOBRE:", lMargemCol + 9, lMargemLin + 5.3, lLocal
        ImprimeString "ATÉ:", lMargemCol + 13.5, lMargemLin + 5.3, lLocal
        ImprimeString "CONDIÇÕES ESPECIAIS:", lMargemCol + 5.3, lMargemLin + 6, lLocal
    
        Printer.Line (lMargemCol + 3.5, lMargemLin + 5.5)-(lMargemCol + 3.5, lMargemLin + lLinMax - 1), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 5, lMargemLin + 5.1)-(lMargemCol + 5, lMargemLin + lLinMax - 0.5), RGB(0, 0, 0)
    
        Printer.Line (lMargemCol + 5, lMargemLin + 6.6)-(lColMax, lMargemLin + 6.6), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 5, lMargemLin + 10)-(lColMax, lMargemLin + 10), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 7, lMargemLin + 10.75)-(lColMax, lMargemLin + 10.75), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 5, lMargemLin + 11.5)-(lColMax, lMargemLin + 11.5), RGB(0, 0, 0)
        Printer.Line (lMargemCol + 7, lMargemLin + 10)-(lMargemCol + 7, lMargemLin + 11.5), RGB(0, 0, 0)
        
        
        ImprimeString "NOME DO SACADO:", lMargemCol + 5.3, lMargemLin + 6.9, lLocal
        ImprimeString "ENDEREÇO:", lMargemCol + 5.3, lMargemLin + 7.4, lLocal
        ImprimeString "MUNICÍPIO:", lMargemCol + 5.3, lMargemLin + 7.9, lLocal
        ImprimeString "CEP:", lMargemCol + 13, lMargemLin + 7.9, lLocal
        ImprimeString "UF:", lMargemCol + 17, lMargemLin + 7.9, lLocal
        ImprimeString "PRAÇA PAGAMENTO:", lMargemCol + 5.3, lMargemLin + 8.4, lLocal
        ImprimeString "FONE:", lMargemCol + 13, lMargemLin + 8.4, lLocal
        ImprimeString "END. COBRANÇA:", lMargemCol + 5.3, lMargemLin + 8.9, lLocal
        ImprimeString "CNPJ(MF)/CPF:", lMargemCol + 5.3, lMargemLin + 9.4, lLocal
        ImprimeString "INSC. EST.:", lMargemCol + 13, lMargemLin + 9.4, lLocal
        
        ImprimeCentralizado "VALOR", lMargemCol + 5, lMargemCol + 7, lMargemLin + 10.05, lLocal
        ImprimeCentralizado "POR", lMargemCol + 5, lMargemCol + 7, lMargemLin + 10.55, lLocal
        ImprimeCentralizado "EXTENSO", lMargemCol + 5, lMargemCol + 7, lMargemLin + 11.05, lLocal
        ImprimeString "Reconheço(emos) a exatidão desta DUPLICATA DE VENDA MERCANTIL, na importância acima que", lMargemCol + 5.3, lMargemLin + 11.55, lLocal
        ImprimeString " pagarei(emos) à " & Empresa.Nome & ", ou à sua ordem na praça e vencimento acima indicados.", lMargemCol + 5.3, lMargemLin + 11.95, lLocal
    
        ImprimeString "Em: ", lMargemCol + 5.3, lMargemLin + 12.8, lLocal
        Printer.FontSize = 12
        ImprimeCentralizado mskDataEmissao, lMargemCol + 5, lMargemCol + 11, lMargemLin + 12.6, lLocal
    
        Printer.FontSize = 8
        ImprimeCentralizado "Data do Aceite", lMargemCol + 5, lMargemCol + 11, lMargemLin + 13.1, lLocal
        ImprimeCentralizado "Assinatura do Sacado", lMargemCol + 13, lColMax, lMargemLin + 13.1, lLocal
        Printer.Line (lMargemCol + 13.5, lMargemLin + 13)-(lColMax - 0.5, lMargemLin + 13), RGB(0, 0, 0)
    End If
End Sub
Private Sub ImpClienteEpson(x_valor As Currency)
    Dim x_linha As String
    Dim i As Integer
    Dim xValorExtenso As String
    x_linha = "                                                                                "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & x_linha
    
    x_linha = "                                                                                "
    'Mid(x_linha, 21, 4) = Format(Cliente.Codigo, "0000")
    Mid(x_linha, 35, 40) = txtRazaoSocial.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    If Val(txtTelefone.Text) > 0 Then
        Mid(x_linha, 58, 20) = "FONE: " & fMascaraTelefone(txtTelefone.Text)
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 35, 40) = txtEndereco.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 35, 28) = txtCidade.Text
    Mid(x_linha, vbInformation, 2) = txtUF.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 35, 40) = txtLocalCobranca.Text
    'End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    If mskCnpj.Text <> "" Then
        Mid(x_linha, 35, 20) = mskCnpj.Text
        'Mid(x_linha, 60, 4) = "IE: "
        Mid(x_linha, vbInformation, 20) = txtInscricaoEstadual.Text
    Else
        Mid(x_linha, 35, 20) = mskCpf.Text
        Mid(x_linha, 60, 4) = "RG: "
        'Mid(x_linha, vbInformation, 20) = Cliente.Identidade
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 59, 15) = "CEP: " & txtCep.Text
    BioImprime "@Printer.Print " & x_linha
   
    x_linha = "                                                                                "
    BioImprime "@Printer.Print " & x_linha
    
    
    xValorExtenso = FazExtenso(x_valor)
    
    For i = Len(xValorExtenso) + 1 To 132
        xValorExtenso = xValorExtenso & "*"
    Next
    
    x_linha = "                                                                                "
    Mid(x_linha, 35, 44) = Mid(xValorExtenso, 1, 44)
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 35, 44) = Mid(xValorExtenso, 45, 44)
    BioImprime "@Printer.Print " & x_linha
    x_linha = "                                                                                "
    Mid(x_linha, 35, 44) = Mid(xValorExtenso, 89, 44)
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpClienteHP(x_valor As Currency)
    Dim x_linha As String
    Dim i As Integer
    Dim xValorExtenso As String
    Printer.FontSize = 10
    Printer.FontBold = True
    'ImprimeString Format(Cliente.Codigo, "0000") & " - " & Cliente.RazaoSocial, lMargemCol + 8.5, lMargemLin + 6.9, lLocal
    ImprimeString txtRazaoSocial.Text, lMargemCol + 8.5, lMargemLin + 6.9, lLocal
    ImprimeString txtEndereco.Text, lMargemCol + 8.5, lMargemLin + 7.4, lLocal
    ImprimeString txtCidade.Text, lMargemCol + 8.5, lMargemLin + 7.9, lLocal
    ImprimeString txtCep.Text, lMargemCol + 14, lMargemLin + 7.9, lLocal
    ImprimeString txtUF.Text, lMargemCol + 18.5, lMargemLin + 7.9, lLocal
    ImprimeString txtPracaPagamento.Text, lMargemCol + 8.5, lMargemLin + 8.4, lLocal
    ImprimeString txtTelefone.Text, lMargemCol + 15.5, lMargemLin + 8.4, lLocal
    ImprimeString txtLocalCobranca.Text, lMargemCol + 8.5, lMargemLin + 8.9, lLocal
    If mskCnpj.Text <> "" Then
        ImprimeString mskCnpj.Text, lMargemCol + 8.5, lMargemLin + 9.4, lLocal
        ImprimeString txtInscricaoEstadual.Text, lMargemCol + 15, lMargemLin + 9.4, lLocal
    Else
        ImprimeString mskCpf.Text, lMargemCol + 8.5, lMargemLin + 9.4, lLocal
    End If
    xValorExtenso = FazExtenso(x_valor)
    For i = Len(xValorExtenso) + 1 To 132
        xValorExtenso = xValorExtenso & "*"
    Next
    ImprimeString Mid(xValorExtenso, 1, 66), lMargemCol + 7.3, lMargemLin + 10.15, lLocal
    ImprimeString Mid(xValorExtenso, 67, 66), lMargemCol + 7.3, lMargemLin + 11#, lLocal
    Printer.EndDoc
End Sub
Private Sub ImpDados()
    Call ImpDuplicata
    If lPagina > 0 Then
        If lLocal = 0 Then
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo
            frm_preview.Show 1
        Else
        End If
    End If
End Sub
Private Sub ImpDuplicata()
    Dim x_linha As String
    Dim i As Integer
    Dim xLocalErro As Integer
    
    On Error GoTo FileError
    
    xLocalErro = 0
    If lLocal = 0 Then
        ImpCabEpson
        xLocalErro = 2
        x_linha = "                                                                                "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        xLocalErro = 3
    Else
        ImpCabHP
        xLocalErro = 4
    End If
    If lLocal = 0 Then
        xLocalErro = 5
        x_linha = "                                                                                "
        Mid(x_linha, 1, 9) = "VEND: " & Format(0, "000")
        Mid(x_linha, 12, 6) = txtNumeroFatura.Text
        i = Len(txtValorFatura.Text)
        Mid(x_linha, 20 + 12 - i, i) = txtValorFatura.Text
        Mid(x_linha, 33, 9) = Mid(txtNumeroDuplicata.Text, 1, 9)
        i = Len(txtValorDuplicata.Text)
        Mid(x_linha, 43 + 12 - i, i) = txtValorDuplicata.Text
        Mid(x_linha, 56, 10) = mskDataVencimento.Text
        BioImprime "@Printer.Print " & x_linha
        xLocalErro = 6
        Call ImpClienteEpson(fValidaValor(txtValorDuplicata.Text))
        xLocalErro = 7
    Else
        
        Printer.FontBold = True
        Printer.FontSize = 10
        ImprimeCentralizado txtValorFatura.Text, lMargemCol, lMargemCol + 3.6, lMargemLin + 4.55, lLocal
        xLocalErro = 8
        ImprimeCentralizado txtNumeroFatura.Text, lMargemCol + 3.6, lMargemCol + 6.5, lMargemLin + 4.55, lLocal
        xLocalErro = 9
        ImprimeCentralizado txtValorDuplicata.Text, lMargemCol + 6.5, lMargemCol + 10.5, lMargemLin + 4.55, lLocal
        xLocalErro = 10
        ImprimeCentralizado txtNumeroDuplicata.Text, lMargemCol + 10.5, lMargemCol + 13, lMargemLin + 4.55, lLocal
        xLocalErro = 11
        Printer.FontBold = True
        Printer.FontSize = 12
        ImprimeCentralizado mskDataVencimento.Text, lMargemCol + 13, lMargemCol + 16.5, lMargemLin + 4.4, lLocal
        xLocalErro = 12
        Call ImpClienteHP(fValidaValor(txtValorDuplicata.Text))
        xLocalErro = 13
    End If
    Exit Sub

FileError:
    MsgBox "Erro na rotina ImpDuplicata:" & vbCrLf & "Erro no local:" & xLocalErro & vbCrLf & "Erro:" & Error, vbInformation, "Erro desconhecido!"
End Sub
Private Sub Relatorio()
    DesativaBotoes
    ZeraVariaveis
    ImpDados
    AtivaBotoes
    cmd_sair.SetFocus
End Sub
Private Sub cmdDataEmissao_Click()
    g_string = mskDataEmissao.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        mskDataEmissao.Text = RetiraGString(1)
        mskDataVencimento.Text = RetiraGString(2)
        txtValorFatura.SetFocus
    Else
        mskDataEmissao.Text = RetiraGString(1)
        mskDataVencimento.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Relatorio
        End If
    End If
    cmd_sair.SetFocus
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    
    If Trim(txtRazaoSocial.Text) = "" Then
        MsgBox "Informe a razão social.", vbInformation, "Dados incompleto!"
        txtRazaoSocial.SetFocus
    ElseIf Trim(txtEndereco.Text) = "" Then
        MsgBox "Informe o endereço.", vbInformation, "Dados incompleto!"
        txtEndereco.SetFocus
    ElseIf Trim(txtCidade.Text) = "" Then
        MsgBox "Informe a cidade.", vbInformation, "Dados incompleto!"
        txtCidade.SetFocus
    ElseIf Trim(txtUF.Text) = "" Then
        MsgBox "Informe a sígla do estado.", vbInformation, "Dados incompleto!"
        txtUF.SetFocus
    ElseIf Val(txtCep.Text) < 1 Then
        MsgBox "Informe o cep.", vbInformation, "Dados incompleto!"
        txtCep.SetFocus
    ElseIf Trim(txtTelefone.Text) = "" Then
        MsgBox "Informe o telefone.", vbInformation, "Dados incompleto!"
        txtTelefone.SetFocus
    ElseIf Trim(txtPracaPagamento.Text) = "" Then
        MsgBox "Informe a praça de pagamento.", vbInformation, "Dados incompleto!"
        txtPracaPagamento.SetFocus
    ElseIf Trim(txtLocalCobranca.Text) = "" Then
        MsgBox "Informe o local de cobrança.", vbInformation, "Dados incompleto!"
        txtLocalCobranca.SetFocus
    ElseIf Not IsDate(mskDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Dados incompleto!"
        mskDataEmissao.SetFocus
    ElseIf Not IsDate(mskDataVencimento.Text) Then
        MsgBox "Informe a data do vencimento.", vbInformation, "Dados incompleto!"
        mskDataVencimento.SetFocus
    ElseIf Not fValidaValor2(txtValorFatura.Text) > 0 Then
        MsgBox "Informe o valor da fatura.", vbInformation, "Dados incompleto!"
        txtValorFatura.SetFocus
    ElseIf Trim(txtNumeroFatura.Text) = "" Then
        MsgBox "Informe o número da fatura.", vbInformation, "Dados incompleto!"
        txtNumeroFatura.SetFocus
    ElseIf Not fValidaValor2(txtValorDuplicata.Text) > 0 Then
        MsgBox "Informe o valor da duplicata.", vbInformation, "Dados incompleto!"
        txtValorDuplicata.SetFocus
    ElseIf Trim(txtNumeroDuplicata.Text) = "" Then
        MsgBox "Informe o número da duplicata.", vbInformation, "Dados incompleto!"
        txtNumeroDuplicata.SetFocus
    ElseIf txtNumeroDuplicata.Text = "" Then
        MsgBox "Informe o número da duplicata.", vbInformation, "Dados incompleto!"
        txtNumeroDuplicata.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    
    lLinha2 = 0
    lMargemCol = 0.5
    lMargemLin = 0.5
    lCabLin = 0
    lColMax = 20.5
    lLinMax = 14
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Relatorio
        End If
    End If
End Sub

Private Sub cmdDataVencimento_Click()
    g_string = mskDataVencimento.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        mskDataEmissao.Text = RetiraGString(1)
        mskDataVencimento.Text = RetiraGString(2)
        txtValorFatura.SetFocus
    Else
        mskDataVencimento.Text = RetiraGString(1)
        txtValorFatura.SetFocus
    End If
    g_string = ""
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 1
    If Not IsDate(mskDataEmissao.Text) Then
        mskDataEmissao.Text = Format(Date, "dd/mm/yyyy")
    End If
    txtPracaPagamento.Text = "A MESMA"
    txtLocalCobranca.Text = "O MESMO"
    txtRazaoSocial.SetFocus
End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub mskCnpj_GotFocus()
    mskCnpj.SelStart = 0
    mskCnpj.SelLength = Len(mskCnpj.Text)
End Sub
Private Sub mskCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtInscricaoEstadual.SetFocus
    End If
End Sub
Private Sub mskCpf_GotFocus()
    mskCpf.SelStart = 0
    mskCpf.SelLength = Len(mskCpf.Text)
End Sub
Private Sub mskCpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskCnpj.SetFocus
    End If
End Sub
Private Sub mskDataEmissao_GotFocus()
    mskDataEmissao.SelStart = 0
    mskDataEmissao.SelLength = 5
End Sub
Private Sub mskDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskDataVencimento.SetFocus
    End If
End Sub
Private Sub txtNumerodDplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub mskDataVencimento_GotFocus()
    mskDataVencimento.SelStart = 0
    mskDataVencimento.SelLength = 5
End Sub
Private Sub mskDataVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtValorFatura.SetFocus
    End If
End Sub
Private Sub txtCep_GotFocus()
    txtCep.SelStart = 0
    txtCep.SelLength = Len(txtCep.Text)
End Sub
Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtTelefone.SetFocus
    End If
End Sub
Private Sub txtCidade_GotFocus()
    txtCidade.SelStart = 0
    txtCidade.SelLength = Len(txtCidade.Text)
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtUF.SetFocus
    End If
End Sub
Private Sub txtEndereco_GotFocus()
    txtEndereco.SelStart = 0
    txtEndereco.SelLength = Len(txtEndereco.Text)
End Sub
Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCidade.SetFocus
    End If
End Sub
Private Sub txtInscricaoEstadual_GotFocus()
    txtInscricaoEstadual.SelStart = 0
    txtInscricaoEstadual.SelLength = Len(txtInscricaoEstadual.Text)
End Sub
Private Sub txtInscricaoEstadual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskDataEmissao.SetFocus
    End If
End Sub
Private Sub txtLocalCobranca_GotFocus()
    txtLocalCobranca.SelStart = 0
    txtLocalCobranca.SelLength = Len(txtLocalCobranca.Text)
End Sub
Private Sub txtLocalCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskCpf.SetFocus
    End If
End Sub
Private Sub txtNumeroDuplicata_GotFocus()
    txtNumeroDuplicata.SelStart = 0
    txtNumeroDuplicata.SelLength = Len(txtNumeroDuplicata.Text)
End Sub
Private Sub txtNumeroDuplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub txtNumeroFatura_GotFocus()
    txtNumeroFatura.SelStart = 0
    txtNumeroFatura.SelLength = Len(txtNumeroFatura.Text)
End Sub
Private Sub txtNumeroFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtValorDuplicata.SetFocus
    End If
End Sub
Private Sub txtNumeroFatura_LostFocus()
    txtNumeroDuplicata.Text = txtNumeroFatura.Text
End Sub
Private Sub txtPracaPagamento_GotFocus()
    txtPracaPagamento.SelStart = 0
    txtPracaPagamento.SelLength = Len(txtPracaPagamento.Text)
End Sub
Private Sub txtPracaPagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtLocalCobranca.SetFocus
    End If
End Sub
Private Sub txtRazaoSocial_GotFocus()
    txtRazaoSocial.SelStart = 0
    txtRazaoSocial.SelLength = Len(txtRazaoSocial.Text)
End Sub
Private Sub txtRazaoSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtEndereco.SetFocus
    End If
End Sub
Private Sub txtTelefone_GotFocus()
    txtTelefone.Text = fDesmascaraTelefone(txtTelefone.Text)
    txtTelefone.SelStart = 0
    txtTelefone.SelLength = Len(txtTelefone.Text)
End Sub
Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPracaPagamento.SetFocus
    End If
End Sub
Private Sub txtTelefone_LostFocus()
    txtTelefone.Text = fMascaraTelefone(txtTelefone.Text)
End Sub
Private Sub txtUF_GotFocus()
    txtUF.SelStart = 0
    txtUF.SelLength = Len(txtUF.Text)
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCep.SetFocus
    End If
End Sub
Private Sub txtValorDuplicata_GotFocus()
    txtValorDuplicata.SelStart = 0
    txtValorDuplicata.SelLength = Len(txtValorDuplicata.Text)
End Sub
Private Sub txtValorDuplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtNumeroDuplicata.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtValorDuplicata_LostFocus()
    If Val(txtValorDuplicata.Text) > 0 Then
        txtValorDuplicata.Text = Format(txtValorDuplicata.Text, "###,##0.00")
    End If
End Sub
Private Sub txtValorFatura_GotFocus()
    txtValorFatura.SelStart = 0
    txtValorFatura.SelLength = Len(txtValorFatura.Text)
End Sub
Private Sub txtValorFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtNumeroFatura.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtValorFatura_LostFocus()
    If Val(txtValorFatura.Text) > 0 Then
        txtValorFatura.Text = Format(txtValorFatura.Text, "###,##0.00")
        txtValorDuplicata.Text = txtValorFatura.Text
    End If
End Sub
