VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_emissao_lmc 
   Caption         =   "Emissão do L.M.C."
   ClientHeight    =   6195
   ClientLeft      =   2145
   ClientTop       =   2100
   ClientWidth     =   6975
   Icon            =   "Rel_lmc.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Rel_lmc.frx":030A
   ScaleHeight     =   6195
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4500
      Picture         =   "Rel_lmc.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5160
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1680
      Picture         =   "Rel_lmc.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprime o L.M.C."
      Top             =   5160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "Rel_lmc.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "Rel_lmc.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_pagina 
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1560
         Width           =   675
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1140
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_data_f 
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
      Begin MSMask.MaskEdBox msk_data_i 
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Página"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Combustível"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Caption         =   "Total das Aferições"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lbl_afericao 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   18
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label lbl_perdas_sobras 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   24
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lbl_estoque_fechamento 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   22
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lbl_estoque_escritural 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lbl_vendas_dia 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   16
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lbl_total_recebido 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lbl_estoque_abertura 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "- Perdas + Sobras"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   4740
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Estoque Fechamento"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   4380
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Estoque Escritural"
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   4020
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Vendas no Dia"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   3180
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Total Recebido"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   2820
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Estoque de Abertura"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2460
      Width           =   2295
   End
End
Attribute VB_Name = "frm_emissao_lmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_margem_lin As Currency
Dim l_margem_col As Currency
Dim l_lin_max As Currency
Dim l_col_max As Currency
Dim l_local As Integer
Dim lTipoCombustivel As String
Dim lData As Date
Dim lDataI As Date
Dim lDataF As Date
Dim lNumeroPaginaLmc As Integer
Dim lNomeProduto As String
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
Dim lNotaEntrada(0 To 20) As String
Dim lDataEntrada(0 To 20) As Date
Dim lQuantidadeEntrada(0 To 20) As Currency
Dim lTanqueEntrada(0 To 20) As String
Dim lQuantidadeNotasEntradas As Integer
Dim l_total_entrada As Currency
Dim l_volume_disponivel As Currency
Dim l_bomba(1 To 40) As Integer       'mudado estas variaveis de 30 para 40 pois ao imprimir
Dim l_fechamento(1 To 40) As Currency ' o lmc de do posto paineiras que contem mais de 30
Dim l_abertura(1 To 40) As Currency   ' bicos dava erro estorando o tamanho da variavel
Dim l_litros_afericao(1 To 40) As Currency 'variaveis mudadas l_bomba, l_fechamento,
Dim l_litros_vendidos(1 To 40) As Currency 'l_abertura, l_litros_afericao, l_litros_vendidos,
Dim l_tanque(1 To 40) As String            'l_tanque.
Dim l_estoque_escritural As Currency
Dim l_perdas_sobras As Currency
Dim l_litros_vendidos_dia As Currency
Dim l_valor_vendas_dia As Currency
Dim l_valor_vendas_mes As Currency
Dim lSQL As String

Private Empresa As New cEmpresa
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoBomba As New cMovimentoBomba
Private LivroLMC As New cLivroLMC

Dim rstAfericao As New adodb.Recordset
Dim rstEntradaCombustivel As New adodb.Recordset
Dim rstMovimentoBomba As New adodb.Recordset

Private Sub BuscaPagina()
    'Dim x_existe As Boolean
    
    Dim xData As String
    Dim xPaginaInicial As Integer
    Dim i As Integer
    
    xData = Format(lData, "dd/mm/yyyy")
    Mid(xData, 1, 2) = "01"
    If LivroLMC.LocalizarCodigo(g_empresa, lTipoCombustivel, "L", xData) Then
        xPaginaInicial = LivroLMC.PaginaInicial + Day(lData) - 1
    Else
        If LivroLMC.LocalizarCodigo(g_empresa, "TC", "L", xData) Then
            i = fUltimoDiaMes(lData)
            If cbo_combustivel.ListIndex = 0 Then
            End If
            xPaginaInicial = LivroLMC.PaginaInicial + (fUltimoDiaMes(lData) * (cbo_combustivel.ListIndex + 1) - fUltimoDiaMes(lData)) + Day(lData) - 1
        Else
            If LivroLMC.LocalizarCombustivelDataAproximada(g_empresa, lTipoCombustivel, "L", lData) Then
                xPaginaInicial = LivroLMC.PaginaInicial - 1
                For i = Month(LivroLMC.DataInicial) To (Month(lData) - 1)
                    xData = Format(lData, "dd/mm/yyyy")
                    Mid(xData, 1, 2) = "01"
                    Mid(xData, 4, 2) = Format(i, "00")
                    xPaginaInicial = xPaginaInicial + fUltimoDiaMes(CDate(xData))
                Next
                xPaginaInicial = xPaginaInicial + Day(lData)
            Else
                MsgBox "Não existe Livro do Lmc Cadastrado!", vbInformation, "Dados Inexistente"
                txt_pagina.Text = ""
                lNumeroPaginaLmc = 0
                Exit Sub
            End If
        End If
    End If
    lNumeroPaginaLmc = xPaginaInicial ' Format(xPaginaInicial, "000")
    txt_pagina.Text = Format(lNumeroPaginaLmc, "000")
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    'If tbl_setup![tipo de formulario] = 1 Then
        l_margem_col = 2.2
        l_margem_lin = 0.5
        l_col_max = 20
        l_lin_max = 26.7
    'Else
    '    l_margem_col = 2.2
    '    l_margem_lin = 2.5
    '    l_col_max = 20
    '    l_lin_max = 28
    'End If
    l_valor_vendas_dia = 0
    l_valor_vendas_mes = 0
    l_total_entrada = 0
    l_litros_vendidos_dia = 0
    lQuantidadeAfericao = 0
    For i = 1 To 40
        l_bomba(i) = 0
        l_fechamento(i) = 0
        l_abertura(i) = 0
        l_litros_afericao(i) = 0
        l_litros_vendidos(i) = 0
        l_tanque(i) = ""
    Next
    For i = 0 To 20
        lNotaEntrada(i) = ""
        lDataEntrada(i) = 0
        lQuantidadeEntrada(i) = 0
        lTanqueEntrada(i) = ""
    Next
    lQuantidadeNotasEntradas = 0
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
Private Sub ImprimeDados()
    Dim x_tanque(1 To 40) As String
    Dim x_string As String
    Dim x_bomba(1 To 15) As String
    Dim x_fechamento(1 To 15) As String
    Dim x_abertura(1 To 15) As String
    Dim x_litros_vendidos(1 To 15) As String
    Dim x_litros_afericao(1 To 15) As String
    Dim i As Integer
    Dim i2 As Integer
    
    Printer.FontSize = 10
    
    Printer.FontSize = 16
    ImprimeCentralizado Format(lNumeroPaginaLmc, "000"), l_margem_col + 15.3, l_margem_col + 18, l_margem_lin + 1.6, l_local
    ImprimeTexto lNomeProduto, l_margem_col + 1.4, l_margem_col + 14.8, l_margem_lin + 2.4, l_local
    Printer.FontSize = 10
    ImprimeCentralizado lData, l_margem_col + 15.85, l_margem_col + 17.8, l_margem_lin + 2.5, l_local
    'ImprimeString "1", l_margem_col + 1.6, l_margem_lin + 4.15, l_local
    'ImprimeString "2", l_margem_col + 3.6, l_margem_lin + 4.15, l_local
    'ImprimeString "3", l_margem_col + 5.6, l_margem_lin + 4.15, l_local
    'ImprimeString "4", l_margem_col + 7.6, l_margem_lin + 4.15, l_local
    'ImprimeString "5", l_margem_col + 9.6, l_margem_lin + 4.15, l_local
    'ImprimeString "6", l_margem_col + 11.6, l_margem_lin + 4.15, l_local
    
    'medição física no início do dia
    For i = 0 To 10
        If lNumeroTanqueAbertura(i) > 0 Then
            ImprimeCentralizadoB Format(lEstoqueAberturaTanque(i), "###,##0.00"), l_margem_col + ((i + 1) * 2) - 2, l_margem_col + ((i + 1) * 2), l_margem_lin + 4.95, l_local
            ImprimeString lNumeroTanqueAbertura(i), l_margem_col + ((i + 1) * 2) - 2 + 1.6, l_margem_lin + 4.15, l_local
        End If
    Next
    'ImprimeCentralizadoB Format(lAberturaTanque_1, "###,##0.00"), l_margem_col, l_margem_col + 2, l_margem_lin + 4.95, l_local
    'ImprimeCentralizadoB Format(lAberturaTanque_2, "###,##0.00"), l_margem_col + 2, l_margem_col + 4, l_margem_lin + 4.95, l_local
    'ImprimeCentralizadoB Format(lAberturaTanque_3, "###,##0.00"), l_margem_col + 4, l_margem_col + 6, l_margem_lin + 4.95, l_local
    'ImprimeCentralizadoB Format(lAberturaTanque_4, "###,##0.00"), l_margem_col + 6, l_margem_col + 8, l_margem_lin + 4.95, l_local
    'ImprimeCentralizadoB Format(lAberturaTanque_5, "###,##0.00"), l_margem_col + 8, l_margem_col + 10, l_margem_lin + 4.95, l_local
    'ImprimeCentralizadoB Format(lAberturaTanque_6, "###,##0.00"), l_margem_col + 10, l_margem_col + 12, l_margem_lin + 4.95, l_local
    ImprimeValor Format(lAberturaTanque, "###,##0.00"), l_margem_col + 12, l_margem_col + 16.8, l_margem_lin + 4.95, l_local
    
    'entradas de combustiveis
    If lDataEntrada(7) <> "0:00:00" Then
        'reduz o tamanho da fonte
        Printer.FontSize = 7
        ImprimeTexto lNotaEntrada(0), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 6.1, l_local
        x_string = ""
        If lDataEntrada(0) <> "0:00:00" Then
            x_string = lDataEntrada(0)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 6.1, l_local
        ImprimeCentralizado lTanqueEntrada(0), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 6.1, l_local
        ImprimeValorB Format(lQuantidadeEntrada(0), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 6.1, l_local
        
        ImprimeTexto lNotaEntrada(1), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 6.4, l_local
        x_string = ""
        If lDataEntrada(1) <> "0:00:00" Then
            x_string = lDataEntrada(1)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 6.4, l_local
        ImprimeCentralizado lTanqueEntrada(1), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 6.4, l_local
        ImprimeValorB Format(lQuantidadeEntrada(1), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 6.4, l_local
        
        ImprimeTexto lNotaEntrada(2), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 6.7, l_local
        x_string = ""
        If lDataEntrada(2) <> "0:00:00" Then
            x_string = lDataEntrada(2)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 6.7, l_local
        ImprimeCentralizado lTanqueEntrada(2), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 6.7, l_local
        ImprimeValorB Format(lQuantidadeEntrada(2), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 6.7, l_local
        
        ImprimeTexto lNotaEntrada(3), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7, l_local
        x_string = ""
        If lDataEntrada(3) <> "0:00:00" Then
            x_string = lDataEntrada(3)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7, l_local
        ImprimeCentralizado lTanqueEntrada(3), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7, l_local
        ImprimeValorB Format(lQuantidadeEntrada(3), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7, l_local
        
        ImprimeTexto lNotaEntrada(4), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7.3, l_local
        x_string = ""
        If lDataEntrada(4) <> "0:00:00" Then
            x_string = lDataEntrada(4)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7.3, l_local
        ImprimeCentralizado lTanqueEntrada(4), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7.3, l_local
        ImprimeValorB Format(lQuantidadeEntrada(4), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7.3, l_local
        
        ImprimeTexto lNotaEntrada(5), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7.6, l_local
        x_string = ""
        If lDataEntrada(5) <> "0:00:00" Then
            x_string = lDataEntrada(5)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7.6, l_local
        ImprimeCentralizado lTanqueEntrada(5), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7.6, l_local
        ImprimeValorB Format(lQuantidadeEntrada(5), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7.6, l_local
        
        ImprimeTexto lNotaEntrada(6), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7.9, l_local
        x_string = ""
        If lDataEntrada(6) <> "0:00:00" Then
            x_string = lDataEntrada(6)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7.9, l_local
        ImprimeCentralizado lTanqueEntrada(6), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7.9, l_local
        ImprimeValorB Format(lQuantidadeEntrada(6), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7.9, l_local
    
        ImprimeTexto lNotaEntrada(7), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 8.2, l_local
        x_string = ""
        If lDataEntrada(7) <> "0:00:00" Then
            x_string = lDataEntrada(7)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 8.2, l_local
        ImprimeCentralizado lTanqueEntrada(7), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 8.2, l_local
        ImprimeValorB Format(lQuantidadeEntrada(7), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.2, l_local
    
'        ImprimeTexto lNotaEntrada(8), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 8.5, l_local
'        x_string = ""
'        If lDataEntrada(8) <> "0:00:00" Then
'            x_string = lDataEntrada(8)
'        End If
'        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 8.5, l_local
'        ImprimeCentralizado lTanqueEntrada(8), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 8.5, l_local
'        ImprimeValorB Format(lQuantidadeEntrada(8), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.5, l_local
'
'        ImprimeTexto lNotaEntrada(9), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 8.8, l_local
'        x_string = ""
'        If lDataEntrada(9) <> "0:00:00" Then
'            x_string = lDataEntrada(9)
'        End If
'        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 8.8, l_local
'        ImprimeCentralizado lTanqueEntrada(9), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 8.8, l_local
'        ImprimeValorB Format(lQuantidadeEntrada(9), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.8, l_local
'
'        ImprimeTexto lNotaEntrada(10), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 10, l_local
'        x_string = ""
'        If lDataEntrada(10) <> "0:00:00" Then
'            x_string = lDataEntrada(10)
'        End If
'        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 10, l_local
'        ImprimeCentralizado lTanqueEntrada(10), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 10, l_local
'        ImprimeValorB Format(lQuantidadeEntrada(10), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 10, l_local
'        'volta a fonte ao normal
'        Printer.FontSize = 10
    Else
        ImprimeTexto lNotaEntrada(0), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 6.25, l_local
        x_string = ""
        If lDataEntrada(0) <> "0:00:00" Then
            x_string = lDataEntrada(0)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 6.25, l_local
        ImprimeCentralizado lTanqueEntrada(0), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 6.25, l_local
        ImprimeValorB Format(lQuantidadeEntrada(0), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 6.25, l_local
        ImprimeTexto lNotaEntrada(1), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 6.75, l_local
        x_string = ""
        If lDataEntrada(1) <> "0:00:00" Then
            x_string = lDataEntrada(1)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 6.75, l_local
        ImprimeCentralizado lTanqueEntrada(1), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 6.75, l_local
        ImprimeValorB Format(lQuantidadeEntrada(1), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 6.75, l_local
        ImprimeTexto lNotaEntrada(2), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7.25, l_local
        x_string = ""
        If lDataEntrada(2) <> "0:00:00" Then
            x_string = lDataEntrada(2)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7.25, l_local
        ImprimeCentralizado lTanqueEntrada(2), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7.25, l_local
        ImprimeValorB Format(lQuantidadeEntrada(2), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7.25, l_local
        ImprimeTexto lNotaEntrada(3), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 7.75, l_local
        x_string = ""
        If lDataEntrada(3) <> "0:00:00" Then
            x_string = lDataEntrada(3)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 7.75, l_local
        ImprimeCentralizado lTanqueEntrada(3), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 7.75, l_local
        ImprimeValorB Format(lQuantidadeEntrada(3), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 7.75, l_local
        ImprimeTexto lNotaEntrada(4), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 8.25, l_local
        x_string = ""
        If lDataEntrada(4) <> "0:00:00" Then
            x_string = lDataEntrada(4)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 8.25, l_local
        ImprimeCentralizado lTanqueEntrada(4), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 8.25, l_local
        ImprimeValorB Format(lQuantidadeEntrada(4), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.25, l_local
        ImprimeTexto lNotaEntrada(5), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 8.75, l_local
        x_string = ""
        If lDataEntrada(5) <> "0:00:00" Then
            x_string = lDataEntrada(5)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 8.75, l_local
        ImprimeCentralizado lTanqueEntrada(5), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 8.75, l_local
        ImprimeValorB Format(lQuantidadeEntrada(5), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.75, l_local
        ImprimeTexto lNotaEntrada(6), l_margem_col + 1.9, l_margem_col + 6.5, l_margem_lin + 9.25, l_local
        x_string = ""
        If lDataEntrada(6) <> "0:00:00" Then
            x_string = lDataEntrada(6)
        End If
        ImprimeCentralizado x_string, l_margem_col + 6.9, l_margem_col + 10, l_margem_lin + 9.25, l_local
        ImprimeCentralizado lTanqueEntrada(6), l_margem_col + 10, l_margem_col + 13.6, l_margem_lin + 9.25, l_local
        ImprimeValorB Format(lQuantidadeEntrada(6), "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 9.25, l_local
    End If
    

    If lDataEntrada(8) <> "0:00:00" Then
        Printer.FontSize = 16
        ImprimeTexto "...", l_margem_col + 16.15, l_margem_col + 16.8, l_margem_lin + 8.1, l_local
        Printer.FontSize = 10
    End If
    
    
    
    For i = 0 To 20
        l_total_entrada = l_total_entrada + lQuantidadeEntrada(i)
    Next
    ImprimeValor Format(l_total_entrada, "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 8.85, l_local
    l_volume_disponivel = l_total_entrada + lAberturaTanque
    ImprimeValor Format(l_volume_disponivel, "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 9.7, l_local
    
    'Dados da movimentação das bombas
    i2 = 0
    For i = 1 To 40
        If l_abertura(i) > 0 Or l_fechamento(i) > 0 Then
            i2 = i2 + 1
            x_tanque(i2) = l_tanque(i)
            x_bomba(i2) = l_bomba(i)
            x_fechamento(i2) = Format(l_fechamento(i), "###,###.00")
            x_abertura(i2) = Format(l_abertura(i), "###,###.00")
            x_litros_afericao(i2) = Format(l_litros_afericao(i), "###,##0.00")
            x_litros_vendidos(i2) = Format(l_litros_vendidos(i) - l_litros_afericao(i), "###,##0.00")
            'l_litros_vendidos_dia = l_litros_vendidos_dia + l_litros_vendidos(i)
        End If
    Next
    Printer.FontSize = 9
    ImprimeCentralizado x_tanque(1), l_margem_col, l_margem_col + 2, l_margem_lin + 10.8, l_local
    ImprimeCentralizado x_bomba(1), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 10.8, l_local
    ImprimeValor x_fechamento(1), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 10.8, l_local
    ImprimeValor x_abertura(1), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 10.8, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 10.8, l_local
    ImprimeValor x_litros_afericao(1), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 10.8, l_local
    ImprimeValor x_litros_vendidos(1), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 10.8, l_local
    
    ImprimeCentralizado x_tanque(2), l_margem_col, l_margem_col + 2, l_margem_lin + 11.1, l_local
    ImprimeCentralizado x_bomba(2), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 11.1, l_local
    ImprimeValor x_fechamento(2), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 11.1, l_local
    ImprimeValor x_abertura(2), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 11.1, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 11.1, l_local
    ImprimeValor x_litros_afericao(2), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 11.1, l_local
    ImprimeValor x_litros_vendidos(2), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 11.1, l_local
    
    ImprimeCentralizado x_tanque(3), l_margem_col, l_margem_col + 2, l_margem_lin + 11.4, l_local
    ImprimeCentralizado x_bomba(3), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 11.4, l_local
    ImprimeValor x_fechamento(3), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 11.4, l_local
    ImprimeValor x_abertura(3), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 11.4, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 11.4, l_local
    ImprimeValor x_litros_afericao(3), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 11.4, l_local
    ImprimeValor x_litros_vendidos(3), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 11.4, l_local
    
    ImprimeCentralizado x_tanque(4), l_margem_col, l_margem_col + 2, l_margem_lin + 11.7, l_local
    ImprimeCentralizado x_bomba(4), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 11.7, l_local
    ImprimeValor x_fechamento(4), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 11.7, l_local
    ImprimeValor x_abertura(4), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 11.7, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 11.7, l_local
    ImprimeValor x_litros_afericao(4), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 11.7, l_local
    ImprimeValor x_litros_vendidos(4), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 11.7, l_local
    
    ImprimeCentralizado x_tanque(5), l_margem_col, l_margem_col + 2, l_margem_lin + 12, l_local
    ImprimeCentralizado x_bomba(5), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 12, l_local
    ImprimeValor x_fechamento(5), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 12, l_local
    ImprimeValor x_abertura(5), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 12, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 12, l_local
    ImprimeValor x_litros_afericao(5), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 12, l_local
    ImprimeValor x_litros_vendidos(5), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 12, l_local
    
    ImprimeCentralizado x_tanque(6), l_margem_col, l_margem_col + 2, l_margem_lin + 12.3, l_local
    ImprimeCentralizado x_bomba(6), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 12.3, l_local
    ImprimeValor x_fechamento(6), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 12.3, l_local
    ImprimeValor x_abertura(6), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 12.3, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 12.3, l_local
    ImprimeValor x_litros_afericao(6), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 12.3, l_local
    ImprimeValor x_litros_vendidos(6), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 12.3, l_local
    
    ImprimeCentralizado x_tanque(7), l_margem_col, l_margem_col + 2, l_margem_lin + 12.6, l_local
    ImprimeCentralizado x_bomba(7), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 12.6, l_local
    ImprimeValor x_fechamento(7), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 12.6, l_local
    ImprimeValor x_abertura(7), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 12.6, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 12.6, l_local
    ImprimeValor x_litros_afericao(7), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 12.6, l_local
    ImprimeValor x_litros_vendidos(7), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 12.6, l_local
    
    ImprimeCentralizado x_tanque(8), l_margem_col, l_margem_col + 2, l_margem_lin + 12.9, l_local
    ImprimeCentralizado x_bomba(8), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 12.9, l_local
    ImprimeValor x_fechamento(8), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 12.9, l_local
    ImprimeValor x_abertura(8), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 12.9, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 12.9, l_local
    ImprimeValor x_litros_afericao(8), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 12.9, l_local
    ImprimeValor x_litros_vendidos(8), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 12.9, l_local
    
    ImprimeCentralizado x_tanque(9), l_margem_col, l_margem_col + 2, l_margem_lin + 13.2, l_local
    ImprimeCentralizado x_bomba(9), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 13.2, l_local
    ImprimeValor x_fechamento(9), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 13.2, l_local
    ImprimeValor x_abertura(9), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 13.2, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 13.2, l_local
    ImprimeValor x_litros_afericao(9), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 13.2, l_local
    ImprimeValor x_litros_vendidos(9), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 13.2, l_local
    
    ImprimeCentralizado x_tanque(10), l_margem_col, l_margem_col + 2, l_margem_lin + 13.5, l_local
    ImprimeCentralizado x_bomba(10), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 13.5, l_local
    ImprimeValor x_fechamento(10), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 13.5, l_local
    ImprimeValor x_abertura(10), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 13.5, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 13.5, l_local
    ImprimeValor x_litros_afericao(10), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 13.5, l_local
    ImprimeValor x_litros_vendidos(10), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 13.5, l_local
    
    ImprimeCentralizado x_tanque(11), l_margem_col, l_margem_col + 2, l_margem_lin + 13.8, l_local
    ImprimeCentralizado x_bomba(11), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 13.8, l_local
    ImprimeValor x_fechamento(11), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 13.8, l_local
    ImprimeValor x_abertura(11), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 13.8, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 13.8, l_local
    ImprimeValor x_litros_afericao(11), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 13.8, l_local
    ImprimeValor x_litros_vendidos(11), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 13.8, l_local
    
    ImprimeCentralizado x_tanque(12), l_margem_col, l_margem_col + 2, l_margem_lin + 14.1, l_local
    ImprimeCentralizado x_bomba(12), l_margem_col + 2, l_margem_col + 4.5, l_margem_lin + 14.1, l_local
    ImprimeValor x_fechamento(12), l_margem_col + 4.5, l_margem_col + 6.8, l_margem_lin + 14.1, l_local
    ImprimeValor x_abertura(12), l_margem_col + 7.2, l_margem_col + 9.6, l_margem_lin + 14.1, l_local
    ImprimeValor "    ", l_margem_col + 10, l_margem_col + 13.2, l_margem_lin + 14.1, l_local
    ImprimeValor x_litros_afericao(12), l_margem_col + 10, l_margem_col + 13, l_margem_lin + 14.1, l_local
    ImprimeValor x_litros_vendidos(12), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 14.1, l_local
    
   
    
    Printer.FontSize = 10
    ImprimeValor Format(l_litros_vendidos_dia, "###,##0.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 15.05, l_local
    
    ImprimeValor Format(l_valor_vendas_dia, "#,###,###,##0.00"), l_margem_col + 6.3, l_margem_col + 9.6, l_margem_lin + 16.25, l_local
    l_estoque_escritural = l_volume_disponivel - l_litros_vendidos_dia
    ImprimeValor Format(l_estoque_escritural, "###,###.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 16.25, l_local
    
    ImprimeValor Format(l_valor_vendas_mes, "#,###,###,##0.00"), l_margem_col + 6.3, l_margem_col + 9.6, l_margem_lin + 17.05, l_local
    ImprimeValor Format(lFechamentoTanque, "###,###.00"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 17.05, l_local
    
    l_perdas_sobras = lFechamentoTanque - l_estoque_escritural
    ImprimeValor Format(l_perdas_sobras, "###,##0.00;(###,##0.00)"), l_margem_col + 13.6, l_margem_col + 16.8, l_margem_lin + 17.85, l_local
    
    ImprimeTexto l_observacao_1, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 19.9, l_local
    ImprimeTexto l_observacao_2, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 20.7, l_local
    ImprimeTexto l_observacao_3, l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 21.5, l_local
    
    If lQuantidadeNotasEntradas >= 8 Then
        MsgBox "Neste dia, teve mais de 8 NF de Entrada de Combustível." & vbCrLf & "Como no papel do LMC impresso cabe apenas 8 NF." & vbCrLf & "Então DEVE ser impresso o Relatório de Entrada de Combustiveis da Data: " & lData & "." & vbCrLf & "E anexa-lo para encadernamento." & vbCrLf & "Em caso de dúvida, entre em contato com o suporte técnico.", vbInformation, "Estouro de capacidade de QTD de NF"
        ImprimeTexto "** NOTAS DE ENTRADAS NO RELATÓRIO EM ANEXO", l_margem_col + 0.1, l_margem_col + 10, l_margem_lin + 22.3, l_local
    End If
    
    
    'ImprimeString "1", l_margem_col + 5, l_margem_lin + 24.4, l_local
    'ImprimeString "2", l_margem_col + 7, l_margem_lin + 24.4, l_local
    'ImprimeString "3", l_margem_col + 9, l_margem_lin + 24.4, l_local
    'ImprimeString "4", l_margem_col + 11, l_margem_lin + 24.4, l_local
    'ImprimeString "5", l_margem_col + 13, l_margem_lin + 24.4, l_local
    'ImprimeString "6", l_margem_col + 15, l_margem_lin + 24.4, l_local
    
    
    'medição física no fim do dia
    For i = 0 To 10
        If lNumeroTanqueFechamento(i) > 0 Then
            ImprimeCentralizadoB Format(lEstoqueFechamentoTanque(i), "###,##0.00"), l_margem_col + ((i + 1) * 2) - 2 + 3.4, l_margem_col + ((i + 1) * 2) + 3.4, l_margem_lin + 25.1, l_local
            ImprimeString lNumeroTanqueFechamento(i), l_margem_col + ((i + 1) * 2) - 2 + 5, l_margem_lin + 24.2, l_local
        End If
    Next
    'ImprimeCentralizadoB Format(lFechamentoTanque_1, "###,##0.00"), l_margem_col + 3.4, l_margem_col + 5.4, l_margem_lin + 25.2, l_local
    'ImprimeCentralizadoB Format(lFechamentoTanque_2, "###,##0.00"), l_margem_col + 5.4, l_margem_col + 7.4, l_margem_lin + 25.2, l_local
    'ImprimeCentralizadoB Format(lFechamentoTanque_3, "###,##0.00"), l_margem_col + 7.4, l_margem_col + 9.4, l_margem_lin + 25.2, l_local
    'ImprimeCentralizadoB Format(lFechamentoTanque_4, "###,##0.00"), l_margem_col + 9.4, l_margem_col + 11.4, l_margem_lin + 25.2, l_local
    'ImprimeCentralizadoB Format(lFechamentoTanque_5, "###,##0.00"), l_margem_col + 11.4, l_margem_col + 13.4, l_margem_lin + 25.2, l_local
    'ImprimeCentralizadoB Format(lFechamentoTanque_6, "###,##0.00"), l_margem_col + 13.4, l_margem_col + 15.4, l_margem_lin + 25.2, l_local
    ImprimeCentralizado Format(lFechamentoTanque, "###,##0.00"), l_margem_col + 15.4, l_margem_col + 17.8, l_margem_lin + 25.2, l_local
End Sub
Private Sub ImprimeGrade()
    Dim x_cgc As String
    'Seleciona tamanho da fonte
    Printer.FontSize = 34
    'Printer.Line (0, 0)-(0, 26)
    Printer.DrawWidth = 8
    'Printer.ForeColor = RGB(256, 0, 0) 'dados em vermelho
    Printer.Line (l_margem_col, l_margem_lin)-(l_col_max, l_lin_max), RGB(0, 0, 0), B
    
    If Empresa.LocalizarCodigo(g_empresa) Then
        x_cgc = fMascaraCNPJ(Empresa.CGC)
        Printer.FontSize = 12
        Printer.DrawWidth = 6
        ImprimeString UCase(Empresa.Nome), l_margem_col + 0.8, l_margem_lin + 0.2, l_local
        ImprimeString "CNPJ: " & x_cgc, l_margem_col + 0.8, l_margem_lin + 0.8, l_local
        ImprimeString "Inscrição Estadual.: " & Empresa.InscricaoEstadual, l_margem_col + 8.8, l_margem_lin + 0.8, l_local
        Printer.Line (l_margem_col, l_margem_lin + 1.4)-(l_col_max, l_margem_lin + 1.4), RGB(0, 0, 0)
    End If
    
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    ImprimeString "LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEIS (L.M.C.)", l_margem_col + 0.8, l_margem_lin + 1.6, l_local
    Printer.FontSize = 8
    ImprimeString "Folha:", l_margem_col + 14.6, l_margem_lin + 1.75, l_local
    Printer.DrawWidth = 3
    Printer.Line (l_margem_col, l_margem_lin + 2.3)-(l_col_max, l_margem_lin + 2.3), RGB(0, 0, 0)
    
    Printer.Line (l_margem_col, l_margem_lin + 3.1)-(l_col_max, l_margem_lin + 3.1), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 14.8, l_margem_lin + 2.3)-(17, l_margem_lin + 3.1), RGB(0, 0, 0)
    ImprimeString "1 Produto:", l_margem_col + 0.1, l_margem_lin + 2.55, l_local
    ImprimeString "2 Data:", l_margem_col + 15, l_margem_lin + 2.55, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 3.9)-(l_col_max, l_margem_lin + 3.9), RGB(0, 0, 0)
    ImprimeString "3 Estoque de abertura (medição física no início do dia)", l_margem_col + 0.1, l_margem_lin + 3.35, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 4.7)-(l_col_max, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 1.4, l_margem_lin + 3.9)-(l_margem_col + 1.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 2, l_margem_lin + 3.9)-(l_margem_col + 2, l_margem_lin + 5.5), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 3.4, l_margem_lin + 3.9)-(l_margem_col + 3.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 4, l_margem_lin + 3.9)-(l_margem_col + 4, l_margem_lin + 5.5), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 5.4, l_margem_lin + 3.9)-(l_margem_col + 5.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 6, l_margem_lin + 3.9)-(l_margem_col + 6, l_margem_lin + 5.5), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 7.4, l_margem_lin + 3.9)-(l_margem_col + 7.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 8, l_margem_lin + 3.9)-(l_margem_col + 8, l_margem_lin + 5.5), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 9.4, l_margem_lin + 3.9)-(l_margem_col + 9.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 10, l_margem_lin + 3.9)-(l_margem_col + 10, l_margem_lin + 23.4), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 11.4, l_margem_lin + 3.9)-(l_margem_col + 11.4, l_margem_lin + 4.7), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 12, l_margem_lin + 3.9)-(l_margem_col + 12, l_margem_lin + 5.5), RGB(0, 0, 0)
    ImprimeString "Tanque", l_margem_col + 0.2, l_margem_lin + 4.15, l_local
    ImprimeString "Tanque", l_margem_col + 2.2, l_margem_lin + 4.15, l_local
    ImprimeString "Tanque", l_margem_col + 4.2, l_margem_lin + 4.15, l_local
    ImprimeString "Tanque", l_margem_col + 6.2, l_margem_lin + 4.15, l_local
    ImprimeString "Tanque", l_margem_col + 8.2, l_margem_lin + 4.15, l_local
    ImprimeString "Tanque", l_margem_col + 10.2, l_margem_lin + 4.15, l_local
    ImprimeString "3.1 Estoque de abertura", l_margem_col + 12.2, l_margem_lin + 4.15, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 5.5)-(l_col_max, l_margem_lin + 5.5), RGB(0, 0, 0)
    
    Printer.Line (l_margem_col, l_margem_lin + 6.1)-(l_col_max, l_margem_lin + 6.1), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 13.6, l_margem_lin + 5.5)-(l_margem_col + 13.6, l_margem_lin + 17.6), RGB(0, 0, 0)
    ImprimeString "4 Volume recebido no dia (em litros)", l_margem_col + 0.1, l_margem_lin + 5.65, l_local
    ImprimeString "4.1 Nr. tanque descarga", l_margem_col + 10.1, l_margem_lin + 5.65, l_local
    ImprimeString "4.2 Volume recebido", l_margem_col + 13.7, l_margem_lin + 5.65, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 8.7)-(l_col_max, l_margem_lin + 8.7), RGB(0, 0, 0)
    ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 6.25, l_local
    ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 6.25, l_local
    ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 6.75, l_local
    ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 6.75, l_local
    ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 7.25, l_local
    ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 7.25, l_local
    ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 7.75, l_local
    ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 7.75, l_local
    ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 8.25, l_local
    ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 8.25, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 9.3)-(l_col_max, l_margem_lin + 9.3), RGB(0, 0, 0)
    ImprimeString "4.3 Total recebido", l_margem_col + 10.1, l_margem_lin + 8.85, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 10.2)-(l_col_max, l_margem_lin + 10.2), RGB(0, 0, 0)
    ImprimeString "5 Volume vendido no dia (em litros)", l_margem_col + 0.1, l_margem_lin + 9.6, l_local
    ImprimeString "4.4 Volume disponível", l_margem_col + 10.1, l_margem_lin + 9.4, l_local
    ImprimeString "(3.1 + 4.3)", l_margem_col + 11, l_margem_lin + 9.8, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 10.8)-(l_col_max, l_margem_lin + 10.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 2, l_margem_lin + 10.2)-(l_margem_col + 2, l_margem_lin + 14.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 4.5, l_margem_lin + 10.2)-(l_margem_col + 4.5, l_margem_lin + 14.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 7.2, l_margem_lin + 10.2)-(l_margem_col + 7.2, l_margem_lin + 14.8), RGB(0, 0, 0)
    ImprimeString "5.1 Tanque", l_margem_col + 0.1, l_margem_lin + 10.35, l_local
    ImprimeString "5.2 Bico", l_margem_col + 2.1, l_margem_lin + 10.35, l_local
    ImprimeString "5.3 + Fechamento", l_margem_col + 4.6, l_margem_lin + 10.35, l_local
    ImprimeString "5.4 - Abertura", l_margem_col + 7.3, l_margem_lin + 10.35, l_local
    ImprimeString "5.5 - Aferição", l_margem_col + 10.1, l_margem_lin + 10.35, l_local
    ImprimeString "5.6 = vendas no bico", l_margem_col + 13.7, l_margem_lin + 10.35, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 14.8)-(l_col_max, l_margem_lin + 14.8), RGB(0, 0, 0)
    
    Printer.Line (l_margem_col, l_margem_lin + 15.6)-(l_col_max, l_margem_lin + 15.6), RGB(0, 0, 0)
    ImprimeString "10 Valor das vendas", l_margem_col + 0.1, l_margem_lin + 15.05, l_local
    ImprimeString "5.7 Vendas no dia", l_margem_col + 10.1, l_margem_lin + 15.05, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 16.8)-(l_col_max, l_margem_lin + 16.8), RGB(0, 0, 0)
    ImprimeString "10.1 Valor das vendas do dia", l_margem_col + 0.1, l_margem_lin + 15.75, l_local
    ImprimeString "(5.7 x Preço bomba)", l_margem_col + 1.1, l_margem_lin + 16.25, l_local
    ImprimeString "6 Estoque escritural", l_margem_col + 10.1, l_margem_lin + 15.75, l_local
    ImprimeString "(4.4 - 5.7)", l_margem_col + 11, l_margem_lin + 16.25, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 17.6)-(l_col_max, l_margem_lin + 17.6), RGB(0, 0, 0)
    ImprimeString "10.2 Valor acumulado no mês", l_margem_col + 0.1, l_margem_lin + 17.05, l_local
    ImprimeString "7 Estoque fechamento (9.1)", l_margem_col + 10.1, l_margem_lin + 17.05, l_local
    
    Printer.Line (l_margem_col + 10, l_margem_lin + 18.4)-(l_col_max, l_margem_lin + 18.4), RGB(0, 0, 0)
    Printer.Line (l_margem_col, l_margem_lin + 18.9)-(l_margem_col + 10, l_margem_lin + 18.9), RGB(0, 0, 0)
    ImprimeString "11 Para uso do revendedor", l_margem_col + 0.1, l_margem_lin + 17.85, l_local
    ImprimeString "8 - Perdas + sobras (*)", l_margem_col + 10.1, l_margem_lin + 17.85, l_local
    ImprimeString "12 Destinado a fiscalização DNC", l_margem_col + 10.1, l_margem_lin + 18.6, l_local
    
    Printer.Line (l_margem_col + 10, l_margem_lin + 21.3)-(l_col_max, l_margem_lin + 21.3), RGB(0, 0, 0)
    Printer.Line (l_margem_col, l_margem_lin + 23.4)-(l_col_max, l_margem_lin + 23.4), RGB(0, 0, 0)
    ImprimeString "13 Observações", l_margem_col + 0.1, l_margem_lin + 19.1, l_local
    ImprimeString "Outros orgãos fiscais", l_margem_col + 10.1, l_margem_lin + 21.5, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 24)-(l_col_max, l_margem_lin + 24), RGB(0, 0, 0)
    ImprimeString "Conciliação dos Estoques", l_margem_col + 7.3, l_margem_lin + 23.5, l_local
   
    Printer.Line (l_margem_col + 3.4, l_margem_lin + 24.8)-(l_col_max, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col, l_margem_lin + 25.6)-(l_col_max, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 3.4, l_margem_lin + 24)-(l_margem_col + 3.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 4.8, l_margem_lin + 24)-(l_margem_col + 4.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 5.4, l_margem_lin + 24)-(l_margem_col + 5.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 6.8, l_margem_lin + 24)-(l_margem_col + 6.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 7.4, l_margem_lin + 24)-(l_margem_col + 7.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 8.8, l_margem_lin + 24)-(l_margem_col + 8.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 9.4, l_margem_lin + 24)-(l_margem_col + 9.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 10.8, l_margem_lin + 24)-(l_margem_col + 10.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 11.4, l_margem_lin + 24)-(l_margem_col + 11.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 12.8, l_margem_lin + 24)-(l_margem_col + 12.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 13.4, l_margem_lin + 24)-(l_margem_col + 13.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 14.8, l_margem_lin + 24)-(l_margem_col + 14.8, l_margem_lin + 24.8), RGB(0, 0, 0)
    Printer.Line (l_margem_col + 15.4, l_margem_lin + 24)-(l_margem_col + 15.4, l_margem_lin + 25.6), RGB(0, 0, 0)
    ImprimeString "9 Fechamento físico", l_margem_col + 0.1, l_margem_lin + 24.5, l_local
    ImprimeString "Tanque", l_margem_col + 3.6, l_margem_lin + 24.2, l_local
    ImprimeString "Tanque", l_margem_col + 5.6, l_margem_lin + 24.2, l_local
    ImprimeString "Tanque", l_margem_col + 7.6, l_margem_lin + 24.2, l_local
    ImprimeString "Tanque", l_margem_col + 9.6, l_margem_lin + 24.2, l_local
    ImprimeString "Tanque", l_margem_col + 11.6, l_margem_lin + 24.2, l_local
    ImprimeString "Tanque", l_margem_col + 13.6, l_margem_lin + 24.2, l_local
    ImprimeString "9.1  Total", l_margem_col + 15.6, l_margem_lin + 24.2, l_local
    ImprimeString "(*) Atenção se o resultado for negativo, pode estar havendo vazamento do produto para o meio ambiente", l_margem_col + 0.5, l_margem_lin + 25.8, l_local
End Sub
Private Sub Relatorio()
    Dim i As Integer
    Dim x_data_teste As String
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
        'l_valor_vendas_dia = l_valor_vendas_dia - MedicaoCombustivel.DescontoDiaAnterior
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, lData + 1, lTipoCombustivel) = False
            i = i + 1
            lEstoqueFechamentoTanque(i) = MedicaoCombustivel.Quantidade
            lNumeroTanqueFechamento(i) = MedicaoCombustivel.NumeroTanque
            lFechamentoTanque = lFechamentoTanque + MedicaoCombustivel.Quantidade
        Loop
    Else
        MsgBox "Não existe medição de combustíveis de fechamento para esta data!", vbInformation, "Atenção!"
    End If
    
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
            lQuantidadeNotasEntradas = lQuantidadeNotasEntradas + 1
            i = i + 1
            'If i = 5 Then
            '    MsgBox "Quantidade de NF acima do permitido!", vbCritical + vbOKOnly, "Relatório não será Impresso!"
            '    Exit Do
            '
            'End If
            lNotaEntrada(i) = rstEntradaCombustivel![Numero da Nota]
            lDataEntrada(i) = lData
            lQuantidadeEntrada(i) = rstEntradaCombustivel!Quantidade
            'lTanqueEntrada(i) = lNumeroTanqueAbertura(0)
            lTanqueEntrada(i) = rstEntradaCombustivel![Numero do Tanque]
            rstEntradaCombustivel.MoveNext
        Loop
    End If
    rstEntradaCombustivel.Close
    
    
    'lê NFe de Devoluçao de Combustivel 5661, 6661
    lSQL = ""
    lSQL = lSQL & "SELECT Numero, Quantidade"
    lSQL = lSQL & "  FROM MovimentoNotaFiscalSaidaItem"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data = " & preparaData(lData)
    lSQL = lSQL & "   AND Cancelado = " & preparaBooleano(False)
    lSQL = lSQL & "   AND (CFOP = " & preparaTexto("5661")
    lSQL = lSQL & "        OR CFOP = " & preparaTexto("6661")
    lSQL = lSQL & "        )"
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel) 'ALEX - ESTAVA REPLICANDO DEVOLUÇÕES PARA TODOS OS COMBUSTIVEIS
    lSQL = lSQL & " ORDER BY Numero"
    Set rstEntradaCombustivel = Conectar.RsConexao(lSQL)
    If rstEntradaCombustivel.RecordCount > 0 Then
        Do Until rstEntradaCombustivel.EOF
            lQuantidadeNotasEntradas = lQuantidadeNotasEntradas + 1
            i = i + 1
            lNotaEntrada(i) = rstEntradaCombustivel!numero
            lDataEntrada(i) = lData
            lQuantidadeEntrada(i) = -rstEntradaCombustivel!Quantidade
            lTanqueEntrada(i) = 1
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
    lSQL = ""
    lSQL = lSQL & "SELECT [Codigo da Bomba], Abertura, Encerrante, [Quantidade da Saida], [Preco de Venda], [Numero do Tanque], [Total Desconto], [Total Acrescimo]"
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
            
            If Not IsNull(rstMovimentoBomba![Total Desconto]) Then
                l_valor_vendas_dia = l_valor_vendas_dia - rstMovimentoBomba![Total Desconto]
            End If
            If Not IsNull(rstMovimentoBomba![Total Acrescimo]) Then
                l_valor_vendas_dia = l_valor_vendas_dia + rstMovimentoBomba![Total Acrescimo]
            End If
            
            rstMovimentoBomba.MoveNext
        Loop
    End If
    rstMovimentoBomba.Close
    'diminui nas vendas do dia os descontos do dia
    l_valor_vendas_dia = l_valor_vendas_dia - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, CDate(lData + 1), CDate(lData + 1), lTipoCombustivel)
    
    
    'calcula vendas do mes
    x_data_teste = lData
    Mid(x_data_teste, 1, 2) = "01"
    l_valor_vendas_mes = MovimentoBomba.ValorVendaPeriodo(g_empresa, CDate(x_data_teste), lData, lTipoCombustivel, 1, 9)
    
    
    'diminui nas vendas do mês os descontos do mês
    l_valor_vendas_mes = l_valor_vendas_mes - MedicaoCombustivel.TotalDescontoCombustivel(g_empresa, (CDate(x_data_teste) + 1), CDate(lData + 1), lTipoCombustivel)
    
    
    'Calcula Aferições do Mês
    x_data_teste = lData
    Mid(x_data_teste, 1, 2) = "01"
    lSQL = "SELECT SUM([Valor Total]) as Total FROM Movimento_Afericao_LMC"
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
    For i = 1 To 40
        l_litros_vendidos_dia = l_litros_vendidos_dia + l_litros_vendidos(i)
    Next
    Dim xVolumeRecebido As Currency
    xVolumeRecebido = 0
    For i = 0 To 20
        xVolumeRecebido = xVolumeRecebido + lQuantidadeEntrada(i)
    Next
    
    lbl_estoque_abertura.Caption = Format(lAberturaTanque, "######,#")
    lbl_total_recebido.Caption = Format(xVolumeRecebido, "###,###")
    lbl_vendas_dia.Caption = Format(l_litros_vendidos_dia, "###,###.0")
    lbl_afericao.Caption = Format(lQuantidadeAfericao, "###,###.0")
    
    lbl_estoque_escritural.Caption = Format((lAberturaTanque + xVolumeRecebido - l_litros_vendidos_dia), "###,###.0")
    lbl_estoque_fechamento.Caption = Format(lFechamentoTanque, "###,###.0")
    lbl_perdas_sobras.Caption = Format((lFechamentoTanque - (lAberturaTanque + xVolumeRecebido - l_litros_vendidos_dia)), "###,###.0;(###,###.0)")
    
    'If (MsgBox("Deseja realmente Imprimir Esta Página?", 4 + 32 + 0, "Imprime L.M.C.!")) = 6 Then
        'seleciona medidas para centímetros
         Printer.ScaleMode = 7
        'Seleciona Formulário de cheque
        'Printer.PaperSize = 256
        'Seleciona largura do formulário
        'Printer.ScaleWidth = 20
        'l_lin_max = Printer.ScaleWidth
        'Seleciona altura do formulário
        'Printer.ScaleHeight = 26
        'l_lin_max = Printer.ScaleHeight
        'Seleciona nome da fonte
        Printer.FontName = "Arial"
        Printer.FontName = "Arial"
        If l_local = 0 Then
            Load frm_preview
        End If
        ImprimeGrade
        ImprimeDados
        If l_local = 0 Then
            frm_preview.Show
        End If
        Printer.EndDoc
    'End If
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_pagina.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex <> -1 Then
        lNomeProduto = Mid(cbo_combustivel, 6, Len(cbo_combustivel))
        lTipoCombustivel = Mid(cbo_combustivel, 1, 2)
        lData = CDate(msk_data_i.Text)
        BuscaPagina
    Else
        cbo_combustivel.SetFocus
    End If
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
    cbo_combustivel.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_combustivel.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        lDataI = CDate(msk_data_i.Text)
        lDataF = CDate(msk_data_f.Text)
        lData = CDate(msk_data_i.Text)
        If txt_pagina.Text = "" Then
            BuscaPagina
        End If
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "Data I:" & msk_data_i.Text & " Data F:" & msk_data_f.Text & " Comb:" & cbo_combustivel.Text & " Pagina:" & txt_pagina.Text)
            Do Until lData > lDataF
                'If lDataI <> lDataF Then
                '    BuscaPagina
                'End If
                Relatorio
                lData = lData + 1
                lNumeroPaginaLmc = lNumeroPaginaLmc + 1
            Loop
        End If
        txt_pagina.Text = ""
    End If
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    If l_local = 0 Then
        Unload frm_preview
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    PreencheCboCombustivel
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione um combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
    MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
    MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    If Not IsDate(msk_data_i.Text) Then
        msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
        msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
    End If
    l_local = 1
End Sub
Private Sub PreencheCboCombustivel()
    Dim rstCombustivel As New adodb.Recordset
        
    cbo_combustivel.Clear
    lSQL = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Ordem, Nome"
    Set rstCombustivel = Conectar.RsConexao(lSQL)
    If rstCombustivel.RecordCount > 0 Then
        Do Until rstCombustivel.EOF
            cbo_combustivel.AddItem rstCombustivel!Codigo & " - " & rstCombustivel!Nome
            rstCombustivel.MoveNext
        Loop
    End If
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_combustivel.SetFocus
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
Private Sub txt_pagina_GotFocus()
'    If lDataI <> lDataF Then
'        cmd_imprimir.SetFocus
'    End If
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
    lNumeroPaginaLmc = CLng(txt_pagina.Text)
End Sub
