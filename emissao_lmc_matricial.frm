VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_lmc_matricial 
   Caption         =   "Emissão do L.M.C. (Impressora Matricial)"
   ClientHeight    =   2745
   ClientLeft      =   2145
   ClientTop       =   2100
   ClientWidth     =   6915
   Icon            =   "emissao_lmc_matricial.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_lmc_matricial.frx":030A
   ScaleHeight     =   2745
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "emissao_lmc_matricial.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Visualiza o L.M.C."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "emissao_lmc_matricial.frx":162A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "emissao_lmc_matricial.frx":2904
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime o L.M.C."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2640
         Picture         =   "emissao_lmc_matricial.frx":3BDE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_pagina 
         Height          =   285
         Left            =   1620
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1080
         Width           =   675
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1620
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
         Caption         =   "&Página"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_lmc_matricial"
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
Dim l_margem_lin As Currency
Dim l_margem_col As Currency
Dim l_lin_max As Currency
Dim l_col_max As Currency
Dim lTipoCombustivel As String
Dim lData As Date
Dim l_nome_produto As String
Dim lEstoqueAberturaTanque(0 To 10) As Currency
Dim lEstoqueFechamentoTanque(0 To 10) As Currency
Dim lNumeroTanqueAbertura(0 To 10) As Integer
Dim lNumeroTanqueFechamento(0 To 10) As Integer

Dim lAberturaTanque As Currency
Dim lFechamentoTanque As Currency
Dim lQuantidadeAfericao As Currency
Dim l_observacao_1 As String
Dim l_observacao_2 As String
Dim l_observacao_3 As String
Dim l_nota_entrada(0 To 10) As String
Dim lDataEntrada(0 To 10) As Date
Dim l_quantidade_entrada(0 To 10) As Currency
Dim l_tanque_entrada(0 To 10) As String
Dim l_total_entrada As Currency
Dim l_volume_disponivel As Currency
Dim l_bomba(1 To 30) As Integer
Dim l_fechamento(1 To 30) As Currency
Dim l_abertura(1 To 30) As Currency
Dim l_litros_afericao(1 To 30) As Currency
Dim l_litros_vendidos(1 To 30) As Currency
Dim l_tanque(1 To 30) As String
Dim l_estoque_escritural As Currency
Dim l_perdas_sobras As Currency
Dim l_litros_vendidos_dia As Currency
Dim l_valor_vendas_dia As Currency
Dim l_valor_vendas_mes As Currency
Dim lSQL As String

Private Empresa As New cEmpresa
Private LivroLMC As New cLivroLMC
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoBomba As New cMovimentoBomba

Dim rstAfericao As New adodb.Recordset
Dim rstEntradaCombustivel As New adodb.Recordset
Dim rstMovimentoBomba As New adodb.Recordset
Private Sub AtualizaRstAfericao(xMaxRecords As Integer)
    On Error GoTo FileError
    rstAfericao.CursorLocation = adUseClient
    rstAfericao.MaxRecords = xMaxRecords
    rstAfericao.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
FileError:
    rstAfericao.Close
    rstAfericao.CursorLocation = adUseClient
    rstAfericao.MaxRecords = xMaxRecords
    rstAfericao.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Private Sub BuscaPagina()
    'Dim x_existe As Boolean
    
    Dim xData As String
    Dim xDataInicial As Date
    Dim xPaginaInicial As Integer
    Dim i As Integer
    
    xData = msk_data.Text
    Mid(xData, 1, 2) = "01"
    xDataInicial = CDate(xData)
    If LivroLMC.LocalizarCodigo(g_empresa, lTipoCombustivel, "L", xData) Then
        xPaginaInicial = LivroLMC.PaginaInicial + Day(CDate(msk_data.Text)) - 1
    Else
        If LivroLMC.LocalizarCodigo(g_empresa, "TC", "L", xData) Then
            i = fUltimoDiaMes(CDate(msk_data.Text))
            If cbo_combustivel.ListIndex = 0 Then
            End If
            xPaginaInicial = LivroLMC.PaginaInicial + (fUltimoDiaMes(CDate(msk_data.Text)) * (cbo_combustivel.ListIndex + 1) - fUltimoDiaMes(CDate(msk_data.Text))) + Day(CDate(msk_data.Text)) - 1
        Else
            If LivroLMC.LocalizarCombustivelDataAproximada(g_empresa, lTipoCombustivel, "L", CDate(msk_data.Text)) Then
                xPaginaInicial = LivroLMC.PaginaInicial - 1
                For i = Month(LivroLMC.DataInicial) To (Month(CDate(msk_data.Text)) - 1)
                    xData = msk_data.Text
                    Mid(xData, 1, 2) = "01"
                    Mid(xData, 4, 2) = Format(i, "00")
                    xPaginaInicial = xPaginaInicial + fUltimoDiaMes(CDate(xData))
                Next
                xPaginaInicial = xPaginaInicial + Day(CDate(msk_data.Text))
            Else
                MsgBox "Não existe Livro do Lmc Cadastrado!", vbInformation, "Dados Inexistente"
                txt_pagina.Text = ""
                Exit Sub
            End If
        End If
    End If
    txt_pagina.Text = Format(xPaginaInicial, "000")
End Sub
Private Sub ImpCabLMC()
    Dim x_linha As String
    Dim i As Integer
    Dim x_cgc As String
    Dim x_posicao_linha As Currency
    If Empresa.LocalizarCodigo(g_empresa) Then
        x_cgc = fMascaraCNPJ(Empresa.CGC)
    End If
    lNomeArquivo = BioCriaImprime
    'seleciona medidas para centímetros
    BioImprime "@@Printer.ScaleMode = 7"
    'BioImprime "@@Printer.PaperSize = 256"
    BioImprime "@@Printer.PaperSize = 1"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    'teste para imprimir letra correta
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    If g_impressora_matricial = True Then
        BioImprime "@@Printer.FontName = Draft 5cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.CurrentY = 0"
        BioImprime "@@Printer.FontBold = True"
        x_linha = "                    LIVRO DE MOVINTACAO DE COMBUSTIVEIS (L.M.C.)"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@Printer.FontName = Draft 17cpi"
    Else
        BioImprime "@@Printer.FontName = Sans Serif 8cpi"
        x_linha = "                                                                "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                          LIVRO DE MOVINTACAO DE COMBUSTIVEIS (L.M.C.)"
        BioImprime "@Printer.Print " & x_linha
        x_linha = "                                                                "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    End If
    x_linha = "       +--------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | RAZAO SOCIAL.:                                                                                         FOLHA DE NUMERO:   ___  |"
    Mid(x_linha, 25, 40) = g_nome_empresa
    Mid(x_linha, 130, 5) = Format(Val(txt_pagina.Text), "00000")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | C.N.P.J......:                                                 INSCRICAO ESTADUAL.:                                            |"
    Mid(x_linha, 25, 18) = x_cgc
    Mid(x_linha, 94, 20) = Empresa.InscricaoEstadual
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------------------------------------------------------------------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 1) PRODUTO:                                                                                     | 2) DATA OPERACAO: __/__/____ |"
    Mid(x_linha, 22, 20) = l_nome_produto
    Mid(x_linha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------------------------------------------------------------------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 3) ESTOQUE DE ABERTURA                                                                                                         |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--+-------------+--+-------------+--+------------+--+------------+--+------------+--+------------+----+-------------------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpLMC()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim xBomba(1 To 30) As Integer
    Dim xFechamento(1 To 30) As Currency
    Dim xAbertura(1 To 30) As Currency
    Dim xLitrosAfericao(1 To 30) As Currency
    Dim xLitrosVendidos(1 To 30) As Currency
    Dim xTanque(1 To 30) As String
    For i2 = 1 To 30
        xBomba(i2) = 0
        xFechamento(i2) = 0
        xAbertura(i2) = 0
        xLitrosAfericao(i2) = 0
        xLitrosVendidos(i2) = 0
        xTanque(i2) = 0
    Next
    x_linha = "       |TQ|             |TQ|             |TQ|            |TQ|            |TQ|            |TQ|            | 3.1|  ESTOQUE  DE  ABERTURA  |"
    'Número do Tanque da Abertura do Dia
    For i2 = 0 To 10
        If lNumeroTanqueAbertura(i2) > 0 Then
            i = 16 * (i2 + 1) + 3
            If i2 = 0 Then
                i = i - 1
            End If
            Mid(x_linha, i, 2) = Format(lNumeroTanqueAbertura(i2), "00")
        End If
    Next
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--+-------------+--+-------------+--+------------+--+------------+--+------------+--+------------+----+-------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                |                |               |               |               |               |                              |"
    'Medição do Tanque No Início do Dia
    For i2 = 0 To 10
        If lNumeroTanqueAbertura(i2) > 0 Then
            i3 = 16 * (i2 + 1) - 1
            If i2 = 0 Then
                i3 = i3 - 1
            End If
            i = Len(Format(lEstoqueAberturaTanque(i2), "###,##0.00"))
            Mid(x_linha, i3 + 10 - i, i) = Format(lEstoqueAberturaTanque(i2), "###,##0.00")
        End If
    Next
    i = Len(Format(lAberturaTanque, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(lAberturaTanque, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +----------------+----------------+---------------+---------------+---------------+---------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 4) VOLUME RECEBIDO NO DIA (EM LITROS)                           | 4.1) NR. DO TANQUE DESCARGA   | 4.2) VOLUME RECEBIDO         |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | NOTA FISCAL NR.                            DE:                  |                               |                              |"
    If Trim(l_nota_entrada(0)) <> "" Then
        i = Len(Format(l_nota_entrada(0), "##,###,##0"))
        Mid(x_linha, 26 + 10 - i, i) = Format(l_nota_entrada(0), "##,###,##0")
        Mid(x_linha, 57, 10) = Format(lDataEntrada(0), "dd/mm/yyyy")
        Mid(x_linha, 90, 2) = l_tanque_entrada(0)
        i = Len(Format(l_quantidade_entrada(0), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(l_quantidade_entrada(0), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | NOTA FISCAL NR.                            DE:                  |                               |                              |"
    If Trim(l_nota_entrada(1)) <> "" Then
        i = Len(Format(l_nota_entrada(1), "##,###,##0"))
        Mid(x_linha, 26 + 10 - i, i) = Format(l_nota_entrada(1), "##,###,##0")
        Mid(x_linha, 57, 10) = Format(lDataEntrada(1), "dd/mm/yyyy")
        Mid(x_linha, 90, 2) = l_tanque_entrada(1)
        i = Len(Format(l_quantidade_entrada(1), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(l_quantidade_entrada(1), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | NOTA FISCAL NR.                            DE:                  |                               |                              |"
    If Trim(l_nota_entrada(2)) <> "" Then
        i = Len(Format(l_nota_entrada(2), "##,###,##0"))
        Mid(x_linha, 26 + 10 - i, i) = Format(l_nota_entrada(2), "##,###,##0")
        Mid(x_linha, 57, 10) = Format(lDataEntrada(2), "dd/mm/yyyy")
        Mid(x_linha, 90, 2) = l_tanque_entrada(2)
        i = Len(Format(l_quantidade_entrada(2), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(l_quantidade_entrada(2), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 | 4.3) TOTAL RECEBIDO           |                              |"
    l_total_entrada = l_quantidade_entrada(0) + l_quantidade_entrada(1) + l_quantidade_entrada(2) + l_quantidade_entrada(3) + l_quantidade_entrada(4) + l_quantidade_entrada(5)
    i = Len(Format(l_total_entrada, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(l_total_entrada, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 5) VOLUME VENDIDO NO DIA (EM LITROS)                            | 4.4) VOLUME DISPONIVEL        |                              |"
    l_volume_disponivel = l_total_entrada + lAberturaTanque
    i = Len(Format(l_volume_disponivel, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(l_volume_disponivel, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 5.1) TANQUE | 5.2) BICO | 5.3) (+)FECHAMENTO | 5.4) (-)ABERTURA | 5.5) (-) AFERICOES            | 5.6) (=) VENDAS NO BICO      |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    i = 0
    For i2 = 1 To 30
        If l_bomba(i2) > 0 Then
            i = i + 1
            xBomba(i) = l_bomba(i2)
            xFechamento(i) = l_fechamento(i2)
            xAbertura(i) = l_abertura(i2)
            xLitrosAfericao(i) = l_litros_afericao(i2)
            xLitrosVendidos(i) = l_litros_vendidos(i2)
            xTanque(i) = l_tanque(i2)
        End If
    Next
    If xAbertura(1) > 0 Or xFechamento(1) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(1)
        Mid(x_linha, 28, 2) = Format(xBomba(1), "00")
        i = Len(Format(xFechamento(1), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(1), "###,##0.00")
        i = Len(Format(xAbertura(1), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(1), "###,##0.00")
        i = Len(Format(xLitrosAfericao(1), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(1), "###,##0.00")
        i = Len(Format(xLitrosVendidos(1), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(1), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(2) > 0 Or xFechamento(2) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(2)
        Mid(x_linha, 28, 2) = Format(xBomba(2), "00")
        i = Len(Format(xFechamento(2), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(2), "###,##0.00")
        i = Len(Format(xAbertura(2), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(2), "###,##0.00")
        i = Len(Format(xLitrosAfericao(2), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(2), "###,##0.00")
        i = Len(Format(xLitrosVendidos(2), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(2), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(3) > 0 Or xFechamento(3) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(3)
        Mid(x_linha, 28, 2) = Format(xBomba(3), "00")
        i = Len(Format(xFechamento(3), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(3), "###,##0.00")
        i = Len(Format(xAbertura(3), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(3), "###,##0.00")
        i = Len(Format(xLitrosAfericao(3), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(3), "###,##0.00")
        i = Len(Format(xLitrosVendidos(3), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(3), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(4) > 0 Or xFechamento(4) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(4)
        Mid(x_linha, 28, 2) = Format(xBomba(4), "00")
        i = Len(Format(xFechamento(4), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(4), "###,##0.00")
        i = Len(Format(xAbertura(4), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(4), "###,##0.00")
        i = Len(Format(xLitrosAfericao(4), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(4), "###,##0.00")
        i = Len(Format(xLitrosVendidos(4), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(4), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(5) > 0 Or xFechamento(5) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(5)
        Mid(x_linha, 28, 2) = Format(xBomba(5), "00")
        i = Len(Format(xFechamento(5), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(5), "###,##0.00")
        i = Len(Format(xAbertura(5), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(5), "###,##0.00")
        i = Len(Format(xLitrosAfericao(5), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(5), "###,##0.00")
        i = Len(Format(xLitrosVendidos(5), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(5), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(6) > 0 Or xFechamento(6) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(6)
        Mid(x_linha, 28, 2) = Format(xBomba(6), "00")
        i = Len(Format(xFechamento(6), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(6), "###,##0.00")
        i = Len(Format(xAbertura(6), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(6), "###,##0.00")
        i = Len(Format(xLitrosAfericao(6), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(6), "###,##0.00")
        i = Len(Format(xLitrosVendidos(6), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(6), "###,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(7) > 0 Or xFechamento(7) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(7)
        Mid(x_linha, 28, 2) = Format(xBomba(7), "00")
        i = Len(Format(xFechamento(7), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(7), "###,##0.00")
        i = Len(Format(xAbertura(7), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(7), "###,##0.00")
        i = Len(Format(xLitrosAfericao(7), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(7), "###,##0.00")
        i = Len(Format(xLitrosVendidos(7), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(7), "###,##0.00")
    End If
    
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(8) > 0 Or xFechamento(8) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(8)
        Mid(x_linha, 28, 2) = Format(xBomba(8), "00")
        i = Len(Format(xFechamento(8), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(8), "###,##0.00")
        i = Len(Format(xAbertura(8), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(8), "###,##0.00")
        i = Len(Format(xLitrosAfericao(8), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(8), "###,##0.00")
        i = Len(Format(xLitrosVendidos(8), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(8), "###,##0.00")
    End If
    
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(9) > 0 Or xFechamento(9) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(9)
        Mid(x_linha, 28, 2) = Format(xBomba(9), "00")
        i = Len(Format(xFechamento(9), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(9), "###,##0.00")
        i = Len(Format(xAbertura(9), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(9), "###,##0.00")
        i = Len(Format(xLitrosAfericao(9), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(9), "###,##0.00")
        i = Len(Format(xLitrosVendidos(9), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(9), "###,##0.00")
    End If
    
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(10) > 0 Or xFechamento(10) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(10)
        Mid(x_linha, 28, 2) = Format(xBomba(10), "00")
        i = Len(Format(xFechamento(10), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(10), "###,##0.00")
        i = Len(Format(xAbertura(10), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(10), "###,##0.00")
        i = Len(Format(xLitrosAfericao(10), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(10), "###,##0.00")
        i = Len(Format(xLitrosVendidos(10), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(10), "###,##0.00")
    End If
    
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(11) > 0 Or xFechamento(11) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(11)
        Mid(x_linha, 28, 2) = Format(xBomba(11), "00")
        i = Len(Format(xFechamento(11), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(11), "###,##0.00")
        i = Len(Format(xAbertura(11), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(11), "###,##0.00")
        i = Len(Format(xLitrosAfericao(11), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(11), "###,##0.00")
        i = Len(Format(xLitrosVendidos(11), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(11), "###,##0.00")
    End If
    
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |             |           |                    |                  |                               |                              |"
    If xAbertura(12) > 0 Or xFechamento(12) > 0 Then
        Mid(x_linha, 15, 2) = xTanque(12)
        Mid(x_linha, 28, 2) = Format(xBomba(12), "00")
        i = Len(Format(xFechamento(12), "###,##0.00"))
        Mid(x_linha, 40 + 10 - i, i) = Format(xFechamento(12), "###,##0.00")
        i = Len(Format(xAbertura(12), "###,##0.00"))
        Mid(x_linha, 60 + 10 - i, i) = Format(xAbertura(12), "###,##0.00")
        i = Len(Format(xLitrosAfericao(12), "###,##0.00"))
        Mid(x_linha, 85 + 10 - i, i) = Format(xLitrosAfericao(12), "###,##0.00")
        i = Len(Format(xLitrosVendidos(12), "###,##0.00"))
        Mid(x_linha, 125 + 10 - i, i) = Format(xLitrosVendidos(12), "###,##0.00")
    End If
    
    
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------+-----------+--------------------+------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |10)   VALOR DAS VENDAS R$|                                       | 5.7) VENDAS NO DIA            |                              |"
    i = Len(Format(l_litros_vendidos_dia, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(l_litros_vendidos_dia, "###,##0.00")
'    ImprimeValor Format(l_litros_vendidos_dia, "###,##0.0"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 14.05, lLocal
'    ImprimeValor Format(l_valor_vendas_dia, "#,###,###,##0.00"), l_margem_col + 6.3, l_margem_col + 9.6, l_margem_lin + 15.25, lLocal
'    l_estoque_escritural = l_volume_disponivel - l_litros_vendidos_dia
'    ImprimeValor Format(l_estoque_escritural, "###,###.0"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 15.25, lLocal
'    ImprimeValor Format(l_valor_vendas_mes, "#,###,###,##0.00"), l_margem_col + 6.3, l_margem_col + 9.6, l_margem_lin + 16.05, lLocal
'    ImprimeValor Format(lFechamentoTanque, "###,###.0"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 16.05, lLocal
'    l_perdas_sobras = lFechamentoTanque - l_estoque_escritural
'    ImprimeValor Format(l_perdas_sobras, "###,##0.0;(###,##0.0)"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 16.85, lLocal
'    ImprimeTexto l_observacao_1, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 18.9, lLocal
'    ImprimeTexto l_observacao_2, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 19.7, lLocal
'    ImprimeTexto l_observacao_3, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 20.5, lLocal
    
    
    
    
    
    
   'x_linha = "         1         2         3         4         5         6         7         8         9        10        11        12        13     13"
   'x_linha = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-------------------------+                                       +-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |10.1) VALOR DAS VENDAS NO DIA                                    | 6)   ESTOQUE ESCRITURAL       |                              |"
    i = Len(Format(l_valor_vendas_dia, "##,###,##0.00"))
    Mid(x_linha, 42 + 13 - i, i) = Format(l_valor_vendas_dia, "##,###,##0.00")
    l_estoque_escritural = l_volume_disponivel - l_litros_vendidos_dia
    i = Len(Format(l_estoque_escritural, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(l_estoque_escritural, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |10.2) VALOR DAS VENDAS DO MES                                    | 7)   ESTOQUE DE FECHAMENTO    |                              |"
    i = Len(Format(l_valor_vendas_mes, "##,###,##0.00"))
    Mid(x_linha, 42 + 13 - i, i) = Format(l_valor_vendas_mes, "##,###,##0.00")
    i = Len(Format(lFechamentoTanque, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(lFechamentoTanque, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |11)   PARA USO DO REVENDEDOR                                     | 8)   (-)PERDAS / (+)SOBRAS (*)|                              |"
    l_perdas_sobras = lFechamentoTanque - l_estoque_escritural
    i = Len(Format(l_perdas_sobras, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(l_perdas_sobras, "###,##0.00;-###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 +-------------------------------+------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 | 12)  DESTINADO A FISCALIZACAO  A.N.P.                        |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |13)   OBSERVACOES                                                |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    Mid(x_linha, 28, 40) = l_observacao_1
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    Mid(x_linha, 28, 40) = l_observacao_2
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 +--------------------------------------------------------------+"
    Mid(x_linha, 28, 40) = l_observacao_3
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 | OUTROS ORGAOS FISCAIS                                        |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    'x_linha = "       |                                                                 |                                                              |"
    'BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                                 |                                                              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +-----------------------------------------------------------------+--------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |                                                  CONCILIACOES DOS ESTOQUE                                                      |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--------------+--+------------+--+------------+--+------------+--+------------+--+------------+--+------------+-----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       |              |TQ|            |TQ|            |TQ|            |TQ|            |TQ|            |TQ|            |    T O T A L    |"
    'Número do Tanque da Fechamento do Dia
    For i2 = 0 To 10
        If lNumeroTanqueFechamento(i2) > 0 Then
            i = 16 * (i2 + 1) + 16
            Mid(x_linha, i, 2) = Format(lNumeroTanqueFechamento(i2), "00")
        End If
    Next
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--------------+--+------------+--+------------+--+------------+--+------------+--+------------+--+------------+-----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | 9)FECH.FISICO|               |               |               |               |               |               |                 |"
    'Medição do Tanque No Final do Dia
    For i2 = 0 To 10
        If lNumeroTanqueFechamento(i2) > 0 Then
            i3 = 16 * (i2 + 1) + 12
            'If i2 = 0 Then
            '    i3 = i3 - 1
            'End If
            i = Len(Format(lEstoqueFechamentoTanque(i2), "###,##0.00"))
            Mid(x_linha, i3 + 10 - i, i) = Format(lEstoqueFechamentoTanque(i2), "###,##0.00")
        End If
    Next
    i = Len(Format(lFechamentoTanque, "###,##0.00"))
    Mid(x_linha, 125 + 10 - i, i) = Format(lFechamentoTanque, "###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--------------+---------------+---------------+---------------+---------------+---------------+---------------+-----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       | (*) ATENCAO: se o resultado for negativo, pode estar havendo vazamento do produto para o meio ambiente.                        |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "       +--------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
    'BioImprime "@@Printer.FontBold = True"
    'x_linha = "|                                                                  PAGINA: ___ |"
    'Mid(x_linha, 3, 40) = g_nome_empresa
    'Mid(x_linha, 76, 3) = Format(lPagina, "000")
    'BioImprime "@Printer.Print " & x_linha
    'BioImprime "@@Printer.FontBold = False"
    'x_linha = "| LIVRO DE PRECOS                                           CIDADE, __/__/____ |"
    'i = Len(Trim(g_cidade_empresa))
    'Mid(x_linha, 37 + 30 - i, i) = Trim(g_cidade_empresa)
    'Mid(x_linha, 69, 10) = msk_data.Text
    'BioImprime "@Printer.Print " & x_linha
    'BioImprime "@@Printer.FontName = Draft 17cpi"
    'BioImprime "@@Printer.FontBold = False"
    'x_linha = "+-------+------------------------------------------+-----+------------------------------------+-------+----------------+----------------+"
    'BioImprime "@Printer.Print " & x_linha
    'x_linha = "| CODIGO| DISCRIMINACAO DOS PRODUTOS               | UN. | TIPO DE TRIBUTACAO                 | ALIQ. | PRECO DE CUSTO | PRECO DE VENDA |"
    'BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|L.M.C. (Impressora Matricial)|@|"
    frm_preview.Show 1
End Sub
Private Sub IncrementaData()
    msk_data.Text = Format(CDate(msk_data.Text) + 1, "dd/mm/yyyy")
    lData = CDate(msk_data.Text)
    BuscaPagina
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    
    l_margem_col = 2.2
    l_margem_lin = 2.5
    l_col_max = 20
    l_lin_max = 28
    l_valor_vendas_dia = 0
    l_valor_vendas_mes = 0
    l_total_entrada = 0
    l_litros_vendidos_dia = 0
    lQuantidadeAfericao = 0
    For i = 1 To 30
        l_bomba(i) = 0
        l_fechamento(i) = 0
        l_abertura(i) = 0
        l_litros_afericao(i) = 0
        l_litros_vendidos(i) = 0
        l_tanque(i) = ""
    Next
    For i = 0 To 10
        l_nota_entrada(i) = ""
        lDataEntrada(i) = 0
        l_quantidade_entrada(i) = 0
        l_tanque_entrada(i) = ""
    Next
    For i = 0 To 10
        lEstoqueAberturaTanque(i) = 0
        lEstoqueFechamentoTanque(i) = 0
        lNumeroTanqueAbertura(i) = 0
        lNumeroTanqueFechamento(i) = 0
    Next
    lAberturaTanque = 0
    l_observacao_1 = ""
    l_observacao_2 = ""
    l_observacao_3 = ""
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set LivroLMC = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoBomba = Nothing
End Sub
Private Sub Relatorio()
    Dim i As Integer
    Dim x_data_teste As String * 10
    ZeraVariaveis
    
    'Localiza Medição de Combustível de Abertura do Dia
    lAberturaTanque = 0
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, lData, lTipoCombustivel) Then
        i = 0
        lEstoqueAberturaTanque(i) = MedicaoCombustivel.Quantidade
        lNumeroTanqueAbertura(i) = MedicaoCombustivel.NumeroTanque
        lAberturaTanque = MedicaoCombustivel.Quantidade
        l_observacao_1 = MedicaoCombustivel.Observacao1
        l_observacao_2 = MedicaoCombustivel.Observacao2
        l_observacao_3 = MedicaoCombustivel.Observacao3
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, lData, lTipoCombustivel) = False
            i = i + 1
            lEstoqueAberturaTanque(i) = MedicaoCombustivel.Quantidade
            lNumeroTanqueAbertura(i) = MedicaoCombustivel.NumeroTanque
            lAberturaTanque = lAberturaTanque + MedicaoCombustivel.Quantidade
        Loop
    Else
        MsgBox "Não existe medição de combustíveis de abertura nesta data!", vbInformation, "Atenção!"
    End If
    
    'Localiza Medição de Combustível de Fechamento do Dia
    lFechamentoTanque = 0
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, lData + 1, lTipoCombustivel) Then
        i = 0
        lEstoqueFechamentoTanque(i) = MedicaoCombustivel.Quantidade
        lNumeroTanqueFechamento(i) = MedicaoCombustivel.NumeroTanque
        lFechamentoTanque = MedicaoCombustivel.Quantidade
        l_valor_vendas_dia = l_valor_vendas_dia - MedicaoCombustivel.DescontoDiaAnterior
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, lData + 1, lTipoCombustivel) = False
            i = i + 1
            lEstoqueFechamentoTanque(i) = MedicaoCombustivel.Quantidade
            lNumeroTanqueFechamento(i) = MedicaoCombustivel.NumeroTanque
            lFechamentoTanque = lFechamentoTanque + MedicaoCombustivel.Quantidade
        Loop
    Else
        MsgBox "Não existe medição de combustíveis de fechamento para esta data!", vbInformation, "Atenção!"
    End If
    
    
'    tbl_entrada_combustivel.Index = "id_data"
'    tbl_entrada_combustivel.Seek ">", g_empresa, lData, lTipoCombustivel, "          "
'    For i = 1 To 3
'        If tbl_entrada_combustivel.NoMatch Then
'            Exit For
'        End If
'        If tbl_entrada_combustivel!Empresa <> g_empresa Then
'            Exit For
'        End If
'        If tbl_entrada_combustivel!Data <> lData Then
'            Exit For
'        End If
'        If tbl_entrada_combustivel![Tipo de Combustivel] <> lTipoCombustivel Then
'            Exit For
'        End If
'        l_nota_entrada(i) = tbl_entrada_combustivel![Numero da Nota]
'        lDataEntrada(i) = tbl_entrada_combustivel!Data
'        l_quantidade_entrada(i) = tbl_entrada_combustivel!Quantidade
'        l_tanque_entrada(i) = lNumeroTanqueAbertura(0)
'        tbl_entrada_combustivel.MoveNext
'        If tbl_entrada_combustivel.EOF Then
'            Exit For
'        End If
'    Next
    
    
    'lê entradas de combustíveis
    lSQL = ""
    lSQL = lSQL & "SELECT [Numero da Nota], Quantidade, [Numero do Tanque]"
    lSQL = lSQL & "  FROM Entrada_Combustivel_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND DATA = " & preparaData(lData)
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel)
    lSQL = lSQL & " ORDER BY [Numero da Nota]"
    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
    i = -1
    If rstEntradaCombustivel.RecordCount > 0 Then
        Do Until rstEntradaCombustivel.EOF
            i = i + 1
            l_nota_entrada(i) = rstEntradaCombustivel![Numero da Nota]
            lDataEntrada(i) = lData
            l_quantidade_entrada(i) = rstEntradaCombustivel!Quantidade
            'l_tanque_entrada(i) = lNumeroTanqueAbertura(0)
            l_tanque_entrada(i) = rstEntradaCombustivel![Numero do Tanque]
            rstEntradaCombustivel.MoveNext
        Loop
    End If
    rstEntradaCombustivel.Close
    
    
    
    'Lê movimentação das Aferições
    lSQL = "SELECT [Codigo da Bomba], Quantidade, [Valor Total] FROM Movimento_Afericao_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data = " & preparaData(lData)
    lSQL = lSQL & " AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel)
    Set rstAfericao = Conectar.RsConexao(lSQL)
    With rstAfericao
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                l_litros_afericao(![Codigo da Bomba]) = !Quantidade
                lQuantidadeAfericao = lQuantidadeAfericao + !Quantidade
                l_valor_vendas_dia = l_valor_vendas_dia - ![Valor Total]
                .MoveNext
            Loop
        End If
        .Close
    End With

    
    
    'lê movimentação das bombas
'    tbl_movimento_bomba.Index = "id_data_tipo_combustivel"
'    tbl_movimento_bomba.Seek ">", g_empresa, lData, lTipoCombustivel, "0", "0"
'    Do Until tbl_movimento_bomba.EOF
'        If tbl_movimento_bomba.NoMatch Then
'            Exit Do
'        End If
'        If tbl_movimento_bomba!Empresa <> g_empresa Then
'            Exit Do
'        End If
'        If tbl_movimento_bomba!Data <> lData Then
'            Exit Do
'        End If
'        If Trim(tbl_movimento_bomba![Tipo de Combustivel]) <> Trim(lTipoCombustivel) Then
'            Exit Do
'        End If
'        i = tbl_movimento_bomba![Codigo da Bomba]
'        l_bomba(i) = tbl_movimento_bomba![Codigo da Bomba]
'        If l_abertura(i) = 0 Then
'            l_abertura(i) = tbl_movimento_bomba!Abertura
'        End If
'        l_fechamento(i) = tbl_movimento_bomba!Encerrante
'        l_litros_vendidos(i) = l_litros_vendidos(i) + tbl_movimento_bomba![Quantidade da Saida]
'        l_tanque(i) = tbl_movimento_bomba![Numero do Tanque]
'        l_valor_vendas_dia = l_valor_vendas_dia + tbl_movimento_bomba![Quantidade da Saida] * tbl_movimento_bomba![Preco de Venda]
'        tbl_movimento_bomba.MoveNext
'    Loop
    
    'lê movimentação das bombas
    lSQL = ""
    lSQL = lSQL & "SELECT [Codigo da Bomba], Abertura, Encerrante, [Quantidade da Saida], [Preco de Venda], [Numero do Tanque]"
    lSQL = lSQL & "  FROM Movimento_Bomba_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data = " & preparaData(lData)
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel)
    lSQL = lSQL & " ORDER BY Empresa, Data, [Tipo de Combustivel], [Codigo da Bomba], Periodo, SubCaixa"
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    If rstMovimentoBomba.RecordCount > 0 Then
        Do Until rstMovimentoBomba.EOF
            i = rstMovimentoBomba![Codigo da Bomba]
            l_bomba(i) = rstMovimentoBomba![Codigo da Bomba]
            If l_abertura(i) = 0 Then
                l_abertura(i) = rstMovimentoBomba!Abertura
            End If
            l_fechamento(i) = rstMovimentoBomba!Encerrante
            l_litros_vendidos(i) = l_litros_vendidos(i) + rstMovimentoBomba![Quantidade da Saida]
            l_tanque(i) = rstMovimentoBomba![Numero do Tanque]
            l_valor_vendas_dia = l_valor_vendas_dia + Format(rstMovimentoBomba![Quantidade da Saida] * rstMovimentoBomba![Preco de Venda], "00000000.00")
            rstMovimentoBomba.MoveNext
        Loop
    End If
    rstMovimentoBomba.Close
    
    'calcula vendas do mes
'    x_data_teste = lData
'    Mid(x_data_teste, 1, 2) = "01"
'    tbl_movimento_bomba.Index = "id_data_tipo_combustivel"
'    tbl_movimento_bomba.Seek ">", g_empresa, x_data_teste, lTipoCombustivel, "0", "0"
'    Do Until tbl_movimento_bomba.EOF
'        If tbl_movimento_bomba.NoMatch Then
'            Exit Do
'        End If
'        If tbl_movimento_bomba!Empresa <> g_empresa Then
'            Exit Do
'        End If
'        If tbl_movimento_bomba!Data > lData Then
'            Exit Do
'        End If
'        If Trim(tbl_movimento_bomba![Tipo de Combustivel]) = Trim(lTipoCombustivel) Then
'            l_valor_vendas_mes = l_valor_vendas_mes + tbl_movimento_bomba![Quantidade da Saida] * tbl_movimento_bomba![Preco de Venda]
'        End If
'        tbl_movimento_bomba.MoveNext
'    Loop
    'calcula vendas do mes
    x_data_teste = lData
    Mid(x_data_teste, 1, 2) = "01"
    l_valor_vendas_mes = MovimentoBomba.ValorVendaPeriodo(g_empresa, CDate(x_data_teste), lData, lTipoCombustivel, 1, 9)
    
    
    'diminui nas vendas do mês os descontos do mês
    l_valor_vendas_mes = l_valor_vendas_mes - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, CDate(CDate(x_data_teste) + 1), CDate(lData + 1), lTipoCombustivel)
    
    
    'Calcula Aferições do Mês
    lSQL = "SELECT SUM([Valor Total]) AS Total FROM Movimento_Afericao_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(x_data_teste))
    lSQL = lSQL & " AND Data <= " & preparaData(lData)
    lSQL = lSQL & " AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel)
    Set rstAfericao = Conectar.RsConexao(lSQL)
    With rstAfericao
        If Not .EOF Then
            If Not IsNull(!total) Then
                l_valor_vendas_mes = l_valor_vendas_mes - !total
            End If
        End If
        .Close
    End With
    
    
    
    l_litros_vendidos_dia = l_litros_vendidos_dia - lQuantidadeAfericao
    For i = 1 To 24
        l_litros_vendidos_dia = l_litros_vendidos_dia + l_litros_vendidos(i)
    Next
    'lbl_estoque_abertura.Caption = Format(lAberturaTanque, "######,#")
    'lbl_total_recebido.Caption = Format((l_quantidade_entrada(1) + l_quantidade_entrada(2) + l_quantidade_entrada(3)), "###,###")
    'lbl_vendas_dia.Caption = Format(l_litros_vendidos_dia, "###,###.0")
    'lbl_afericao.Caption = Format(lQuantidadeAfericao, "###,###.0")
    
    'lbl_estoque_escritural.Caption = Format((lAberturaTanque + l_quantidade_entrada(1) + l_quantidade_entrada(2) + l_quantidade_entrada(3) - l_litros_vendidos_dia), "###,###.0")
    'lbl_estoque_fechamento.Caption = Format(lFechamentoTanque, "###,###.0")
    'lbl_perdas_sobras.Caption = Format((lFechamentoTanque - (lAberturaTanque + l_quantidade_entrada(1) + l_quantidade_entrada(2) + l_quantidade_entrada(3) - l_litros_vendidos_dia)), "###,###.0;(###,###.0)")
    
    'If (MsgBox("Deseja realmente Imprimir Esta Página?", 4 + 32 + 0, "Imprime L.M.C.!")) = 6 Then
        'seleciona medidas para centímetros
         'Printer.ScaleMode = 7
        'Seleciona Formulário de cheque
        'Printer.PaperSize = 256
        'Seleciona largura do formulário
        'Printer.ScaleWidth = 20
        'l_lin_max = Printer.ScaleWidth
        'Seleciona altura do formulário
        'Printer.ScaleHeight = 26
        'l_lin_max = Printer.ScaleHeight
        'Seleciona nome da fonte
        'Printer.FontName = "Arial"
        'Printer.FontName = "Arial"
        'If lLocal = 0 Then
        '    Load frm_preview
        'End If
        
        
        Call ImpCabLMC
        Call ImpLMC
        'If lLocal = 0 Then
        '    frm_preview.Show
        'End If
        'Printer.EndDoc
    'End If
    IncrementaData
End Sub
Private Sub cbo_combustivel_GotFocus()
    SendMessageLong cbo_combustivel.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_pagina.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex <> -1 Then
        l_nome_produto = Mid(cbo_combustivel, 6, Len(cbo_combustivel))
        lTipoCombustivel = Mid(cbo_combustivel, 1, 2)
    Else
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cbo_combustivel.SetFocus
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
Private Sub cmd_sair_Click()
    Finaliza
    If lLocal = 0 Then
        Unload frm_preview
    End If
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
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione um combustível.", 64, "Atenção!"
        cbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
    'If g_nome_usuario = "L.M.C." Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    'Else
    '    MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
    'End If
    
    lLocal = 1
    lData = g_data_def
    PreencheCboCombustivel
End Sub
Private Sub PreencheCboCombustivel()
    Dim rstCombustivel As New adodb.Recordset
        
    cbo_combustivel.Clear
    lSQL = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
    Set rstCombustivel = Conectar.RsConexao(lSQL)
    If rstCombustivel.RecordCount > 0 Then
        Do Until rstCombustivel.EOF
            cbo_combustivel.AddItem rstCombustivel!Codigo & " - " & rstCombustivel!Nome
            rstCombustivel.MoveNext
        Loop
    End If
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(lData, "dd/mm/yyyy")
    End If
    msk_data.SelStart = 0
    msk_data.SelLength = 2
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If IsDate(msk_data.Text) Then
        lData = CDate(msk_data.Text)
    End If
End Sub
Private Sub txt_pagina_GotFocus()
    BuscaPagina
    txt_pagina.SelStart = 0
    txt_pagina.SelLength = Len(txt_pagina.Text)
End Sub
Private Sub txt_pagina_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_pagina_LostFocus()
    txt_pagina.Text = Format(Val(txt_pagina.Text), "##000")
End Sub

