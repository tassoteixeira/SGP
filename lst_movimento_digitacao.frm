VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_movimento_digitacao 
   Caption         =   "Emissão do Movimento da Digitação"
   ClientHeight    =   2775
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   4815
   Icon            =   "lst_movimento_digitacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   4815
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1080
      Picture         =   "lst_movimento_digitacao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime as movimentações digitadas."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2940
      Picture         =   "lst_movimento_digitacao.frx":1914
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_digitacao.frx":2FA6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_digitacao.frx":4280
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_movimento_digitacao.frx":555A
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
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
Attribute VB_Name = "emissao_movimento_digitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
'Fim de variáveis padrão para relatório
Dim lSQl As String
'Início variáveis MovimentoBomba
Dim l_periodo_i As Integer
Dim l_periodo_f As Integer
Dim l_abertura_bomba(1 To 12) As Currency
Dim l_encerrante_bomba(1 To 12) As Currency
Dim l_litros_saida(1 To 12) As Currency
Dim l_valor_saida(1 To 12) As Currency
Dim lBombaAvista As Currency
Dim l_valor_afericao As Currency
Dim l_litros_a As Currency
Dim l_litros_aa As Currency
Dim l_litros_d As Currency
Dim l_litros_da As Currency
Dim l_litros_g As Currency
Dim l_litros_ga As Currency
Dim l_litros_l As Currency
Dim l_litros_b As Currency
Dim l_litros_t As Currency
Dim l_valor_a As Currency
Dim l_valor_aa As Currency
Dim l_valor_d As Currency
Dim l_valor_da As Currency
Dim l_valor_g As Currency
Dim l_valor_ga As Currency
Dim l_valor_l As Currency
Dim l_valor_b As Currency
Dim l_valor_t As Currency
Dim l_hist_ch_predatado(1 To 3) As Currency
Dim l_hist_ch_vista(1 To 3) As Currency
Dim l_hist_dinheiro(1 To 3) As Currency
Dim l_hist_nota_firma(1 To 3) As Currency
Dim l_hist_amex(1 To 3) As Currency
Dim l_hist_dinners(1 To 3) As Currency
Dim l_hist_hipercheque(1 To 3) As Currency
Dim l_hist_visa(1 To 3) As Currency
Dim l_hist_assalto(1 To 3) As Currency
Dim l_hist_afericao(1 To 3) As Currency
Dim l_hist_transferencia(1 To 3) As Currency
Dim l_hist_total(1 To 3) As Currency
Dim l_dif_caixa(1 To 3) As Currency
Dim l_nome_funcionario As String
'Fim variáveis MovimentoBomba
'Início variáveis Cheque
Dim l_data As Date
Dim lSubTotal As Currency
Dim lTotal As Currency
Dim lSubQtd As Currency
Dim lTotalQtd As Currency
Dim lSubDias As Currency
Dim lTotalDias As Currency
'Fim variáveis Cheque
'Início variáveis Nota
Dim l_cliente As Long
Dim l_conveniado As Long
Dim l_numero_nota As Long
'Fim variáveis Nota
'Início variáveis Oleos_Filtros
Dim l_quantidade As Currency
Dim l_valor As Currency
Dim l_his_ch_predatado As Currency
Dim l_his_ch_vista As Currency
Dim l_his_dinheiro As Currency
Dim l_his_nota_firma As Currency
Dim l_his_total As Currency
Dim l_dife_caixa As Currency
'Fim variáveis Oleos_Filtros
Dim tbl_bomba As Table
Dim tbl_cartao_credito As Table
Dim tbl_cliente As Table
Dim tbl_cliente_conveniado As Table
Dim tbl_combustivel As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_afericao As Table
Dim tbl_movimento_bomba As Table
Dim tbl_movimento_cartao_credito As Table
Dim tbl_movimento_historico As Table
Dim tbl_movimento_lubrificante As Table
Dim tbl_movimento_nota As Table
Dim tbl_produto As Table
Dim tbl_tabela_premiacao As Table

Private MovCheque As New cMovimentoCheque
Private rsCheque As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_bomba.Close
    tbl_cartao_credito.Close
    tbl_cliente.Close
    tbl_cliente_conveniado.Close
    tbl_combustivel.Close
    tbl_funcionario.Close
    tbl_movimento_afericao.Close
    tbl_movimento_bomba.Close
    tbl_movimento_cartao_credito.Close
    tbl_movimento_historico.Close
    tbl_movimento_lubrificante.Close
    tbl_movimento_nota.Close
    tbl_produto.Close
    tbl_tabela_premiacao.Close
    
    Set MovCheque = Nothing
End Sub
Private Sub ZeraVariaveisBomba()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 1 To 12
        l_abertura_bomba(i) = 0
        l_encerrante_bomba(i) = 0
        l_litros_saida(i) = 0
        l_valor_saida(i) = 0
    Next
    lBombaAvista = 0
    l_valor_afericao = 0
    l_litros_a = 0
    l_litros_aa = 0
    l_litros_d = 0
    l_litros_da = 0
    l_litros_g = 0
    l_litros_ga = 0
    l_litros_l = 0
    l_litros_b = 0
    l_litros_t = 0
    l_valor_a = 0
    l_valor_aa = 0
    l_valor_d = 0
    l_valor_da = 0
    l_valor_g = 0
    l_valor_ga = 0
    l_valor_l = 0
    l_valor_b = 0
    l_valor_t = 0
    For i = 1 To 3
        l_hist_ch_predatado(i) = 0
        l_hist_ch_vista(i) = 0
        l_hist_dinheiro(i) = 0
        l_hist_nota_firma(i) = 0
        l_hist_amex(i) = 0
        l_hist_dinners(i) = 0
        l_hist_visa(i) = 0
        l_hist_hipercheque(i) = 0
        l_hist_assalto(i) = 0
        l_hist_afericao(i) = 0
        l_hist_transferencia(i) = 0
        l_hist_total(i) = 0
        l_dif_caixa(i) = 0
    Next
End Sub
Private Sub ZeraVariaveisCartao()
    lLinha = 0
    lPagina = 0
    lTotal = 0
End Sub
Private Sub ZeraVariaveisCheque()
    lLinha = 0
    lPagina = 0
    l_data = 0
    lSubTotal = 0
    lTotal = 0
    lSubQtd = 0
    lTotalQtd = 0
    lSubDias = 0
    lTotalDias = 0
End Sub
Private Sub ZeraVariaveisNota()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lTotal = 0
    l_cliente = 0
    l_conveniado = 0
    l_numero_nota = 0
End Sub
Private Sub ZeraVariaveisLubrificante()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    l_quantidade = 0
    l_valor = 0
    l_his_ch_predatado = 0
    l_his_ch_vista = 0
    l_his_dinheiro = 0
    l_his_nota_firma = 0
    l_his_total = 0
    l_dife_caixa = 0
End Sub
Private Sub RelatorioNota(x_tipo_movimento As Integer)
    ZeraVariaveisNota
    'Verifica movimento
    tbl_movimento_nota.Seek ">", g_empresa, CDate(msk_data_i), l_periodo_i, 0, 0, 0
    If Not tbl_movimento_nota.NoMatch Then
        If tbl_movimento_nota!Empresa = g_empresa And tbl_movimento_nota![Data do Abastecimento] <= CDate(msk_data_f) Then
            Call ImpDadosNota(x_tipo_movimento)
        End If
    End If
End Sub
Private Sub RelatorioLubrificante(x_tipo_movimento As Integer)
    Dim flag_imprime As Boolean
    flag_imprime = False
    ZeraVariaveisLubrificante
    'Verifica movimento
    With tbl_movimento_historico
        .Index = "id_data"
        If l_periodo_i = 4 Then
            x_tipo_movimento = 2
        End If
        .Seek "=", g_empresa, CDate(msk_data_i), l_periodo_i, 1, x_tipo_movimento
        If l_periodo_i = 4 Then
            x_tipo_movimento = 1
        End If
        If .NoMatch Then
            If (MsgBox("Histórico não cadastrado!" & Chr(10) & "Deseja continuar?", vbYesNo + vbDefaultButton2, "Erro de integridade!")) = 7 Then
                Exit Sub
            End If
        End If
    End With
    Call CalculaHistoricoLubrificante(2)
    With tbl_movimento_lubrificante
        .Index = "id_data"
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_f, x_tipo_movimento, 0, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Data <= CDate(msk_data_f) And !Periodo <= l_periodo_f Then
                    flag_imprime = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
    If flag_imprime Then
        Call ImpDadosLubrificante(x_tipo_movimento)
    End If
End Sub
Private Sub Relatorios()
    Dim i As Integer
    For i = 1 To 5
        l_periodo_i = i
        l_periodo_f = i
        If i = 5 Then
            l_periodo_i = 1
            l_periodo_f = 4
        End If
        If i < 5 Then
            Call RelatorioBomba
            Call RelatorioCheque(1)
            Call RelatorioCheque(2)
            Call RelatorioNota(1)
            Call RelatorioNota(2)
            Call RelatorioCartao(1, 2)
            Call RelatorioCartao(2, 2)
            Call RelatorioCartao(1, 3)
            Call RelatorioCartao(2, 3)
            Call RelatorioCartao(1, 1)
            Call RelatorioCartao(2, 1)
            Call RelatorioCartao(1, 4)
            Call RelatorioCartao(2, 4)
            Call RelatorioLubrificante(1)
        Else
            i = 5
'            Call RelatorioLubrificante(1)
'            Call RelatorioCheque(2)
'            Call RelatorioNota(2)
'            Call RelatorioCartao(2, 2)
'            Call RelatorioCartao(2, 3)
'            Call RelatorioCartao(2, 1)
'            Call RelatorioCartao(2, 4)
            Call RelatorioBomba
'            If g_empresa <> 2 Then
'                msk_data_i = CDate(msk_data_i) + 1
'                msk_data_f = CDate(msk_data_f) + 1
'                Call RelatorioCheque(0)
'                msk_data_i = CDate(msk_data_i) - 1
'                msk_data_f = CDate(msk_data_f) - 1
'            Else
'                Call RelatorioCheque(0)
'            End If
        End If
    Next
    cmd_sair.SetFocus
End Sub
Private Sub ImpTotalLubrificante()
    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 12, 1) = "+"
    Mid(x_linha, 55, 1) = "+"
    Mid(x_linha, 61, 1) = "+"
    Mid(x_linha, 80, 1) = "+"
    Mid(x_linha, 96, 1) = "+"
    Mid(x_linha, 115, 1) = "+"
    Mid(x_linha, 126, 1) = "+"
    Mid(x_linha, 130, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 61, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    Mid(x_linha, 96, 1) = "|"
    Mid(x_linha, 115, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Mid(x_linha, 63, 16) = "TOTAL DAS VENDAS"
    i = Len(Format(l_quantidade, "####,##0.00"))
    Mid(x_linha, 84 + 11 - i, i) = Format(l_quantidade, "####,##0.00")
    i = Len(Format(l_valor, "##,###,##0.00"))
    Mid(x_linha, 101 + 13 - i, i) = Format(l_valor, "##,###,##0.00")
    Printer.FontName = "Sans Serif 17cpi"
    y_local = Printer.CurrentY
    ImprimeTexto "  ", 1, 2, 2, 1
    Printer.CurrentY = y_local
    Printer.FontBold = True
    Printer.Print x_linha
    Printer.CurrentY = y_local
    Printer.Print " "
    Printer.FontBold = False
End Sub
Private Sub RelatorioBomba()
    ZeraVariaveisBomba
    'Verifica movimento
    Call CalculaHistoricoBomba
    Call CalculaLubrificanteBomba
    With tbl_movimento_bomba
        .Index = "id_data"
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 0
        If Not .NoMatch Then
            If !Empresa = g_empresa Then
                If !Data <= CDate(msk_data_f) And !Periodo <= l_periodo_f Then
                    ImpDadosBomba
                End If
            End If
        End If
    End With
End Sub
Private Sub RelatorioCartao(x_tipo_movimento As Integer, x_cartao_credito As Integer)
    Dim x_ok As Boolean
    x_ok = False
    ZeraVariaveisCartao
    'Verifica movimento_cartao_credito
    With tbl_movimento_cartao_credito
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                    If x_tipo_movimento = 0 Or x_tipo_movimento = ![Tipo do Movimento] Then
                        If x_cartao_credito = 0 Or x_cartao_credito = ![Codigo do Cartao] Then
                            x_ok = True
                            Exit Do
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End If
        If x_ok Then
            .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 0
            Call ImpDadosCartao(x_tipo_movimento, x_cartao_credito)
        End If
    End With
End Sub
Private Sub RelatorioCheque(x_tipo_movimento As Integer)
    ZeraVariaveisCheque
    'Verifica Movimento_Cheque
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "SELECT [Data do Vencimento], [Data de Emissao], Valor"
    lSQl = lSQl & "  FROM Movimento_Cheque"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND [Data do Vencimento] >= " & preparaData(msk_data_i.Text)
    lSQl = lSQl & "   AND [Data do Vencimento] <= " & preparaData(msk_data_f.Text)
    lSQl = lSQl & " ORDER BY [Data do Vencimento], [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao], [Numero da Conta], [Numero do Cheque]"
    'Abre RecordSet
    Set rsCheque = New adodb.Recordset
    Set rsCheque = Conectar.RsConexao(lSQl)
    If rsCheque.RecordCount > 0 Then
        Call ImpDadosCheque(x_tipo_movimento)
    End If
    If rsCheque.State = 1 Then
        rsCheque.Close
    End If
End Sub
Private Sub ImpCodificacaoContabil()
    Dim x_linha As String
    Dim i As Integer
    Printer.FontName = "Sans Serif 17cpi"
'                       1         2         3         4         5         6         7         8         9        10        11        12        13     13
'              12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "|        VENDA A VISTA       | VENDAS C/CHEQUE POS DATADO |                            |                            |                   |"
    Printer.Print x_linha
    x_linha = "| DEBITAR...:   1-9          | DEBITAR...:                |                            |                            |                   |"
    If l_hist_ch_predatado(3) > 0 Then
        Mid(x_linha, 44, 5) = "166-0"
    End If
    Printer.Print x_linha
    x_linha = "| CREDITAR..: 185-6          | CREDITAR..:                |                            |                            |                   |"
    If l_hist_ch_predatado(3) > 0 Then
        Mid(x_linha, 44, 5) = "137-6"
    End If
    Printer.Print x_linha
    x_linha = "| HISTORICO.:   2-7          | HISTORICO.:                |                            |                            |                   |"
    If l_hist_ch_predatado(3) > 0 Then
        Mid(x_linha, 44, 5) = "  2-7"
    End If
    Printer.Print x_linha
    x_linha = "| VALOR.....:                | VALOR.....:                |                            |                            |                   |"
    If l_hist_ch_predatado(3) > 0 Then
        i = Len(Format(l_hist_ch_predatado(3), "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(l_hist_ch_predatado(3), "###,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "+----------------------------+----------------------------+----------------------------+----------------------------+-------------------+"
    Printer.Print x_linha
    x_linha = "|    VENDAS C/ CARTAO VISA   | VENDAS C/ CARTAO CREDICARD |    VENDAS C/ CARTAO SOLO   | VENDAS C/ CARTAO HIPERCHEQ |                   |"
    Printer.Print x_linha
    x_linha = "| DEBITAR...:                | DEBITAR...:                | DEBITAR...:                | DEBITAR...:                |                   |"
    If l_hist_visa(1) > 0 Then
        Mid(x_linha, 15, 5) = "164-3"
    End If
    If l_hist_dinners(1) > 0 Then
        Mid(x_linha, 44, 5) = "  5-1"
    End If
    If l_hist_amex(1) > 0 Then
        Mid(x_linha, 73, 5) = "163-5"
    End If
    If l_hist_hipercheque(1) > 0 Then
        Mid(x_linha, 102, 5) = "165-1"
    End If
    Printer.Print x_linha
    x_linha = "| CREDITAR..:                | CREDITAR..:                | CREDITAR..:                | CREDITAR..:                |                   |"
    If l_hist_visa(1) > 0 Then
        Mid(x_linha, 15, 5) = "136-8"
    End If
    If l_hist_dinners(1) > 0 Then
        Mid(x_linha, 44, 5) = "136-8"
    End If
    If l_hist_amex(1) > 0 Then
        Mid(x_linha, 73, 5) = "136-8"
    End If
    If l_hist_hipercheque(1) > 0 Then
        Mid(x_linha, 102, 5) = "136-8"
    End If
    Printer.Print x_linha
    x_linha = "| HISTORICO.:                | HISTORICO.:                | HISTORICO.:                | HISTORICO.:                |                   |"
    If l_hist_visa(1) > 0 Then
        Mid(x_linha, 15, 5) = "  2-7"
    End If
    If l_hist_dinners(1) > 0 Then
        Mid(x_linha, 44, 5) = "  2-7"
    End If
    If l_hist_amex(1) > 0 Then
        Mid(x_linha, 73, 5) = "  2-7"
    End If
    If l_hist_hipercheque(1) > 0 Then
        Mid(x_linha, 102, 5) = "  2-7"
    End If
    Printer.Print x_linha
    x_linha = "| VALOR.....:                | VALOR.....:                | VALOR.....:                | VALOR.....:                |                   |"
    If l_hist_visa(1) > 0 Then
        i = Len(Format(l_hist_visa(1), "###,###,##0.00"))
        Mid(x_linha, 15 + 14 - i, i) = Format(l_hist_visa(1), "###,###,##0.00")
    End If
    If l_hist_dinners(1) > 0 Then
        i = Len(Format(l_hist_dinners(1), "###,###,##0.00"))
        Mid(x_linha, 44 + 14 - i, i) = Format(l_hist_dinners(1), "###,###,##0.00")
    End If
    If l_hist_amex(1) > 0 Then
        i = Len(Format(l_hist_amex(1), "###,###,##0.00"))
        Mid(x_linha, 73 + 14 - i, i) = Format(l_hist_amex(1), "###,###,##0.00")
    End If
    If l_hist_hipercheque(1) > 0 Then
        i = Len(Format(l_hist_hipercheque(1), "###,###,##0.00"))
        Mid(x_linha, 102 + 14 - i, i) = Format(l_hist_hipercheque(1), "###,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "+--- Cerrado Informática. ---+----------------------------+----------------------------+----------------------------+-------------------+"
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
End Sub
Private Sub ImpDadosBomba()
    LoopMovimentoBomba
    CalculaAfericaoBomba
    If l_litros_t > 0 Then
        ImpCabBomba
        Call ImpDetBomba(1, l_abertura_bomba(1), l_encerrante_bomba(1), l_litros_saida(1), l_valor_saida(1), "Ch.Predatado:", "@N@" & l_hist_ch_predatado(1))
        Call ImpDetBomba(2, l_abertura_bomba(2), l_encerrante_bomba(2), l_litros_saida(2), l_valor_saida(2), "Ch.Vista....:", "@N@" & l_hist_ch_vista(1))
        Call ImpDetBomba(3, l_abertura_bomba(3), l_encerrante_bomba(3), l_litros_saida(3), l_valor_saida(3), "Dinheiro....:", "@N@" & l_hist_dinheiro(1))
        Call ImpDetBomba(4, l_abertura_bomba(4), l_encerrante_bomba(4), l_litros_saida(4), l_valor_saida(4), "Nota Firma..:", "@N@" & l_hist_nota_firma(1))
        Call ImpDetBomba(5, l_abertura_bomba(5), l_encerrante_bomba(5), l_litros_saida(5), l_valor_saida(5), "Cred/Dinners:", "@N@" & l_hist_dinners(1))
        Call ImpDetBomba(6, l_abertura_bomba(6), l_encerrante_bomba(6), l_litros_saida(6), l_valor_saida(6), "Sollo/Amex..:", "@N@" & l_hist_amex(1))
        Call ImpDetBomba(7, l_abertura_bomba(7), l_encerrante_bomba(7), l_litros_saida(7), l_valor_saida(7), "Visa........:", "@N@" & l_hist_visa(1))
        Call ImpDetBomba(8, l_abertura_bomba(8), l_encerrante_bomba(8), l_litros_saida(8), l_valor_saida(8), "Hipercheque.:", "@N@" & l_hist_hipercheque(1))
        Call ImpDetBomba(9, l_abertura_bomba(9), l_encerrante_bomba(9), l_litros_saida(9), l_valor_saida(9), "Aferição/Dev:", "@N@" & l_hist_afericao(1))
        Call ImpDetBomba(10, l_abertura_bomba(10), l_encerrante_bomba(10), l_litros_saida(10), l_valor_saida(10), "Transf......:", "@N@" & l_hist_transferencia(1))
        Call ImpDetBomba(11, l_abertura_bomba(11), l_encerrante_bomba(11), l_litros_saida(11), l_valor_saida(11), "Assalto.....:", "@N@" & l_hist_assalto(1))
        Call ImpDetBomba(12, l_abertura_bomba(12), l_encerrante_bomba(12), l_litros_saida(12), l_valor_saida(12), "Responsavel: ", "@A@" & l_nome_funcionario)
        Call ImpDetBomba(0, 0, 0, 0, 0, "Total Geral: ", "@N@" & l_hist_total(1))
        Call ImpResumoHistoricosBomba
        Call ImpResumoCombustiveisBomba
        Call ImpResumoChequePreDatadoBomba
        Call ImpResumoDiferencaPrecoBomba
        If Not g_caixa_unificado Then
            If g_usuario = 8 Then
                Call ImpCodificacaoContabil
            End If
        End If
    End If
    Printer.EndDoc
End Sub
Private Sub ImpDadosCartao(x_tipo_movimento As Integer, x_cartao_credito As Integer)
    Call LoopMovimentoCartao(x_tipo_movimento, x_cartao_credito)
    If lTotal > 0 Then
        ImpTotalCartao
        Printer.EndDoc
    End If
End Sub
Private Sub ImpDadosCheque(x_tipo_movimento As Integer)
    Dim x_linha As String
    'loop movimento de cheques
    Do Until rsCheque.EOF
        If x_tipo_movimento = 0 And g_empresa <> 2 Then
            If rsCheque("Data do Vencimento").Value > CDate(msk_data_f.Text) Then
                Exit Do
            End If
        Else
            If rsCheque("Data de Emissao").Value > CDate(msk_data_f) Then
                Exit Do
            End If
        End If
        If x_tipo_movimento = 0 Or rsCheque("Tipo do Movimento").Value = x_tipo_movimento Then
            If lPagina = 0 Then
                Call ImpCabCheque(x_tipo_movimento)
            End If
            If lLinha >= 57 Then
                x_linha = "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
                Mid(x_linha, 84, 22) = " Cerrado Informática. "
                Printer.Print x_linha
                Printer.NewPage
                Call ImpCabCheque(x_tipo_movimento)
            End If
            ImpDetCheque
            lSubTotal = lSubTotal + rsCheque("valor").Value
            lTotal = lTotal + rsCheque("valor").Value
            lSubQtd = lSubQtd + 1
            lTotalQtd = lTotalQtd + 1
            lSubDias = lSubDias + DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value)
            lTotalDias = lTotalDias + DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value)
        End If
        rsCheque.MoveNext
    Loop
    If x_tipo_movimento = 0 And g_empresa <> 2 Then
        ImpSubTotalCheque
    End If
    If lTotal > 0 Then
        ImpTotalCheque
        Printer.EndDoc
    End If
End Sub
Private Sub ImpDadosLubrificante(x_tipo_movimento As Integer)
    Dim x_linha As String
    tbl_movimento_lubrificante.Index = "id_produto"
    Call LoopMovimentoLubrificante(x_tipo_movimento)
    If l_valor > 0 Then
        ImpTotalLubrificante
        ImpHistoricoLubrificante
        Printer.EndDoc
    End If
End Sub
Private Sub ImpDadosNota(x_tipo_movimento As Integer)
    Dim x_linha As String
    'loop movimento de notas de abastecimento
    With tbl_movimento_nota
        Do Until .EOF
            If !Empresa <> g_empresa Or ![Data do Abastecimento] > CDate(msk_data_f) Then
                Exit Do
            End If
            If x_tipo_movimento = 0 Or ![Tipo do Movimento] = x_tipo_movimento Then
                If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                    If lPagina = 0 Then
                        Call ImpCabNota(x_tipo_movimento)
                        ImpClienteNota
                    End If
                    If lLinha >= 57 Then
                        x_linha = String(137, "-")
                        Mid(x_linha, 1, 1) = "+"
                        Mid(x_linha, 12, 1) = "+"
                        Mid(x_linha, 21, 1) = "+"
                        Mid(x_linha, 64, 1) = "+"
                        Mid(x_linha, 77, 1) = "+"
                        Mid(x_linha, 87, 1) = "+"
                        Mid(x_linha, 106, 1) = "+"
                        Mid(x_linha, 137, 1) = "+"
                        Printer.Print x_linha
                        Printer.NewPage
                        Call ImpCabNota(x_tipo_movimento)
                    End If
                    If ![Codigo do Cliente] <> l_cliente Or ![Codigo do Conveniado] <> l_conveniado Or ![Numero da Nota] <> l_numero_nota Then
                        ImpClienteNota
                    End If
                    ImpProdutoNota
                    lSubTotal = lSubTotal + ![Valor Total]
                    lTotal = lTotal + ![Valor Total]
                End If
            End If
            .MoveNext
        Loop
    End With
    If lSubTotal > 0 Then
        ImpSubTotalNota
    End If
    If lTotal > 0 Then
        ImpTotalNota
        Printer.EndDoc
    End If
End Sub
Private Sub ImpSubTotalCheque()
    Dim x_linha As String
    Dim i As Integer
    If lSubTotal > 0 Then
        x_linha = "|            |       |           |          |     |               |            |                                          |             |"
        Mid(x_linha, 35, 10) = "*** TOTAL "
        i = Len(Format(lSubDias / lSubQtd, "#0.00"))
        Mid(x_linha, 46 + 5 - i, i) = Format(lSubDias / lSubQtd, "#0.00")
        i = Len(Format(lSubTotal, "###,###,##0.00"))
        Mid(x_linha, 52 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
        Mid(x_linha, 82, 17) = "NÚMERO DE CHEQUES"
        i = Len(Format(lSubQtd, "####"))
        Mid(x_linha, 104 + 4 - i, i) = Format(lSubQtd, "####")
        Mid(x_linha, 108, 4) = " EM "
        Mid(x_linha, 112, 10) = Format(l_data, "dd/mm/yyyy")
        Printer.FontName = "Sans Serif 17cpi"
        Printer.Print x_linha
        lLinha = lLinha + 1
        lSubTotal = 0
        lSubQtd = 0
        lSubDias = 0
    End If
End Sub
Private Sub ImpSubTotalNota()
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 69, 19) = "TOTAL DA NOTA.....:"
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 91 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    Mid(x_linha, 106, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 12, 1) = "+"
    Mid(x_linha, 21, 1) = "+"
    Mid(x_linha, 64, 1) = "+"
    Mid(x_linha, 77, 1) = "+"
    Mid(x_linha, 87, 1) = "+"
    Mid(x_linha, 106, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    lLinha = lLinha + 2
    lSubTotal = 0
End Sub
Private Sub ImpTotalCartao()
    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 12, 1) = "+"
    Mid(x_linha, 20, 1) = "+"
    Mid(x_linha, 27, 1) = "+"
    Mid(x_linha, 41, 1) = "+"
    Mid(x_linha, 52, 1) = "+"
    Mid(x_linha, 70, 1) = "+"
    Mid(x_linha, 93, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 40, 11) = "*** TOTAL.:"
    Mid(x_linha, 52, 1) = "|"
    i = Len(Format(lTotal, "#,###,##0.00"))
    Mid(x_linha, 57 + 12 - i, i) = Format(lTotal, "#,###,##0.00")
    Mid(x_linha, 70, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Printer.FontName = "Sans Serif 17cpi"
    y_local = Printer.CurrentY
    ImprimeTexto "  ", 1, 2, 2, 1
    Printer.CurrentY = y_local
    Printer.FontBold = True
    Printer.Print x_linha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print x_linha
    Printer.CurrentY = y_local
    Printer.Print " "
    Printer.FontBold = False
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 52, 1) = "+"
    Mid(x_linha, 70, 1) = "+"
    Mid(x_linha, 95, 22) = " Cerrado Informática. "
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
    Printer.Print " "
End Sub
Private Sub ImpTotalCheque()
    Dim x_linha As String
    Dim i As Integer
    Printer.Print "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    x_linha = "|                                           |     |               |            |                                          |             |"
    Mid(x_linha, 35, 10) = "*** TOTAL "
    i = Len(Format(lTotalDias / lTotalQtd, "#0.00"))
    Mid(x_linha, 46 + 5 - i, i) = Format(lTotalDias / lTotalQtd, "#0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 52 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    Mid(x_linha, 82, 17) = "NÚMERO DE CHEQUES"
    i = Len(Format(lTotalQtd, "####"))
    Mid(x_linha, 104 + 4 - i, i) = Format(lTotalQtd, "####")
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    x_linha = "+-------------------------------------------+-----+---------------+------------+------------------------------------------+-------------+"
    Mid(x_linha, 84, 22) = " Cerrado Informática. "
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
    Printer.Print " "
End Sub
Private Sub ImpTotalNota()
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 69, 19) = "TOTAL DO PERIODO..:"
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 91 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    Mid(x_linha, 106, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    Mid(x_linha, 87, 1) = "+"
    Mid(x_linha, 106, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
    Printer.Print " "
End Sub
Private Sub LoopMovimentoBomba()
    Dim i As Integer
    'loop movimento das bombas
    With tbl_movimento_bomba
        Do Until .EOF
            If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                Exit Do
            End If
            If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                i = ![Codigo da Bomba]
                If l_abertura_bomba(i) = 0 Then
                    l_abertura_bomba(i) = !Abertura
                End If
                l_encerrante_bomba(i) = !Encerrante
                l_litros_saida(i) = l_litros_saida(i) + ![Quantidade da Saida]
                l_valor_saida(i) = l_valor_saida(i) + (![Quantidade da Saida] * ![Preco de Venda])
                Select Case Trim(![Tipo de Combustivel])
                    Case "A"
                        l_litros_a = l_litros_a + ![Quantidade da Saida]
                        l_valor_a = l_valor_a + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Case "AA"
                        l_litros_aa = l_litros_aa + ![Quantidade da Saida]
                        l_valor_aa = l_valor_aa + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Case "D"
                        l_litros_d = l_litros_d + ![Quantidade da Saida]
                        l_valor_d = l_valor_d + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Case "DA"
                        l_litros_da = l_litros_da + ![Quantidade da Saida]
                        l_valor_da = l_valor_da + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Case "G"
                        l_litros_g = l_litros_g + ![Quantidade da Saida]
                        l_valor_g = l_valor_g + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Case "GA"
                        l_litros_ga = l_litros_ga + ![Quantidade da Saida]
                        l_valor_ga = l_valor_ga + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                End Select
                If ![Numero do Tanque] = 2 Then
                    lBombaAvista = lBombaAvista + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                End If
                l_litros_t = l_litros_t + ![Quantidade da Saida]
                l_valor_t = l_valor_t + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopMovimentoCartao(x_tipo_movimento As Integer, x_cartao_credito As Integer)
    'loop movimento do cartao de credito
    Dim x_linha As String
    Dim x_nome_cartao As String * 40
    With tbl_movimento_cartao_credito
        Do Until .EOF
            If !Empresa <> g_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                Exit Do
            End If
            If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                If x_tipo_movimento = 0 Or x_tipo_movimento = ![Tipo do Movimento] Then
                    If x_cartao_credito = 0 Or x_cartao_credito = ![Codigo do Cartao] Then
                        If lPagina = 0 Then
                            ImpCabCartao
                        End If
                        If lLinha >= 60 Then
                            x_linha = String(137, "-")
                            Mid(x_linha, 1, 1) = "+"
                            Mid(x_linha, 12, 1) = "+"
                            Mid(x_linha, 20, 1) = "+"
                            Mid(x_linha, 27, 1) = "+"
                            Mid(x_linha, 41, 1) = "+"
                            Mid(x_linha, 52, 1) = "+"
                            Mid(x_linha, 70, 1) = "+"
                            Mid(x_linha, 93, 1) = "+"
                            Mid(x_linha, 95, 22) = " Cerrado Informática. "
                            Mid(x_linha, 137, 1) = "+"
                            Printer.Print x_linha
                            Printer.NewPage
                            ImpCabCartao
                        End If
                        'Le tabela auxiliar
                        tbl_cartao_credito.Seek "=", ![Codigo do Cartao]
                        If Not tbl_cartao_credito.NoMatch Then
                            x_nome_cartao = tbl_cartao_credito!Nome
                        Else
                            x_nome_cartao = "** Não Cadastrado **"
                        End If
                        Call ImpDetCartao(![Data de Emissao], !Periodo, ![Numero do Lancamento], x_nome_cartao, ![Data do Vencimento], !valor, ![Numero do Cartao])
                        lTotal = lTotal + !valor
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopMovimentoLubrificante(x_tipo_movimento As Integer)
    'loop movimento dos lubrificantes
    Dim i As Integer
    Dim x_linha As String
    tbl_produto.Index = "id_nome"
    tbl_produto.MoveFirst
    With tbl_movimento_lubrificante
        Do Until tbl_produto.EOF
            .Seek ">=", g_empresa, tbl_produto!Codigo, CDate(msk_data_i), l_periodo_i, x_tipo_movimento, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Produto2] <> tbl_produto!Codigo Or !Data > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= l_periodo_f And !Periodo <= l_periodo_f Then
                        If ![Tipo do Movimento] = x_tipo_movimento Or x_tipo_movimento = 0 Then
                            If lPagina = 0 Then
                                Call ImpCabLubrificante(x_tipo_movimento)
                            End If
                            If lLinha >= 60 Then
                                x_linha = String(137, "-")
                                Mid(x_linha, 1, 1) = "+"
                                Mid(x_linha, 12, 1) = "+"
                                Mid(x_linha, 55, 1) = "+"
                                Mid(x_linha, 61, 1) = "+"
                                Mid(x_linha, 80, 1) = "+"
                                Mid(x_linha, 96, 1) = "+"
                                Mid(x_linha, 115, 1) = "+"
                                Mid(x_linha, 126, 1) = "+"
                                Mid(x_linha, 130, 1) = "+"
                                Mid(x_linha, 137, 1) = "+"
                                Mid(x_linha, 14, 22) = " Cerrado Informática. "
                                Printer.Print x_linha
                                Printer.NewPage
                                Call ImpCabLubrificante(x_tipo_movimento)
                            End If
                            'Le tabela auxiliar
                            Call ImpDetLubrificante(![Codigo do Produto2], tbl_produto!Nome, tbl_produto!unidade, ![Valor Venda], !Quantidade, ![Valor Total], !Data, !Periodo, ![Codigo do Funcionario])
                            l_valor = l_valor + ![Valor Total]
                            l_quantidade = l_quantidade + !Quantidade
                        End If
                    End If
                    .MoveNext
                Loop
            End If
            tbl_produto.MoveNext
        Loop
    End With
    tbl_produto.Index = "id_codigo"
End Sub
Private Sub CalculaAfericaoBomba()
    Dim i As Integer
    'loop movimento de Afericao
    l_valor_afericao = 0
    With tbl_movimento_afericao
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 0, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                    Select Case Trim(![Tipo de Combustivel])
                        Case "A"
                            l_litros_a = l_litros_a - !Quantidade
                            l_valor_a = l_valor_a - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                        Case "AA"
                            l_litros_aa = l_litros_aa - !Quantidade
                            l_valor_aa = l_valor_aa - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                        Case "D"
                            l_litros_d = l_litros_d - !Quantidade
                            l_valor_d = l_valor_d - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                        Case "DA"
                            l_litros_da = l_litros_da - !Quantidade
                            l_valor_da = l_valor_da - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                        Case "G"
                            l_litros_g = l_litros_g - !Quantidade
                            l_valor_g = l_valor_g - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                        Case "GA"
                            l_litros_ga = l_litros_ga - !Quantidade
                            l_valor_ga = l_valor_ga - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                    End Select
                    l_litros_t = l_litros_t - !Quantidade
                    l_valor_t = l_valor_t - Format(!Quantidade * ![Preco de Venda], "#########0.00")
                    l_valor_afericao = l_valor_afericao + ![Valor Total]
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub CalculaHistoricoBomba()
    Dim i As Integer
    'loop movimento do historico
    With tbl_movimento_historico
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 1, 1
        If Not .NoMatch Then
            If msk_data_i = msk_data_f And l_periodo_i = l_periodo_f Then
                tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
                If tbl_funcionario.NoMatch Then
                    l_nome_funcionario = ""
                Else
                    l_nome_funcionario = tbl_funcionario!Nome
                End If
            End If
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                    i = ![Tipo do Movimento]
                    l_hist_ch_predatado(i) = l_hist_ch_predatado(i) + ![Cheque Pre-Datado]
                    l_hist_ch_vista(i) = l_hist_ch_vista(i) + ![Cheque A Vista]
                    l_hist_dinheiro(i) = l_hist_dinheiro(i) + !Dinheiro
                    l_hist_nota_firma(i) = l_hist_nota_firma(i) + !Nota
                    l_hist_amex(i) = l_hist_amex(i) + !Amex
                    l_hist_dinners(i) = l_hist_dinners(i) + !Dinners
                    l_hist_visa(i) = l_hist_visa(i) + !Visa
                    l_hist_hipercheque(i) = l_hist_hipercheque(i) + !Hipercheque
                    l_hist_assalto(i) = l_hist_assalto(i) + !Assalto
                    l_hist_afericao(i) = l_hist_afericao(i) + !Afericao
                    l_hist_transferencia(i) = l_hist_transferencia(i) + !Transferencia
                    l_hist_total(i) = l_hist_total(i) + !total
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub CalculaHistoricoLubrificante(x_tipo_movimento As Integer)
    Dim i As Integer
    'loop movimento do historico
    With tbl_movimento_historico
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, 1, x_tipo_movimento
        If Not .NoMatch Then
            tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
            If tbl_funcionario.NoMatch Then
                l_nome_funcionario = ""
            Else
                l_nome_funcionario = tbl_funcionario!Nome
            End If
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Periodo = l_periodo_i Then
                    If ![Tipo do Movimento] = x_tipo_movimento Then
                        l_his_ch_predatado = l_his_ch_predatado + ![Cheque Pre-Datado]
                        l_his_ch_vista = l_his_ch_vista + ![Cheque A Vista]
                        l_his_dinheiro = l_his_dinheiro + !Dinheiro
                        l_his_nota_firma = l_his_nota_firma + !Nota
                        l_his_total = l_his_total + !total
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub CalculaLubrificanteBomba()
    Dim i As Integer
    'loop movimento de lubrificante
    With tbl_movimento_lubrificante
        .Index = "id_data"
        .Seek ">=", g_empresa, CDate(msk_data_i), l_periodo_i, "1", 0, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                    Exit Do
                End If
                If !Periodo >= l_periodo_i And !Periodo <= l_periodo_f Then
                    If ![Tipo do Movimento] = 1 Then
                        l_valor_l = l_valor_l + ![Valor Total]
                    ElseIf ![Tipo do Movimento] = 2 Then
                        l_valor_b = l_valor_b + ![Valor Total]
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpResumoDiferencaPrecoBomba()
    Dim x_linha As String
    Dim i As Integer
    Dim x_valor As Currency
    If lBombaAvista > 0 Then
        x_valor = 0
        Printer.FontName = "Sans Serif 17cpi"
        x_linha = "| VENDAS DE BOMBAS A VISTA.:                           Cheques A Vista + Dinheiro.:                          Diferenca.:                |"
        i = Len(Format(lBombaAvista, "###,###,##0.00"))
        Mid(x_linha, 30 + 14 - i, i) = Format(lBombaAvista, "###,###,##0.00")
        x_valor = x_valor + l_hist_ch_vista(1) + l_hist_dinheiro(1)
        i = Len(Format(x_valor, "###,###,##0.00"))
        Mid(x_linha, 85 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
        x_valor = x_valor - lBombaAvista
        i = Len(Format(x_valor, "###,###,##0.00"))
        Mid(x_linha, 122 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
        Printer.Print x_linha
        x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
        If g_usuario = 8 Then
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
        End If
        Printer.Print x_linha
        Printer.FontName = "Draft 10cpi"
    End If
End Sub
Private Sub ImpResumoHistoricosBomba()
    Dim x_linha As String
    Dim i As Integer
    l_hist_ch_predatado(3) = l_hist_ch_predatado(3) + l_hist_ch_predatado(2) + l_hist_ch_predatado(1)
    l_hist_ch_vista(3) = l_hist_ch_vista(3) + l_hist_ch_vista(2) + l_hist_ch_vista(1)
    l_hist_dinheiro(3) = l_hist_dinheiro(3) + l_hist_dinheiro(2) + l_hist_dinheiro(1)
    l_hist_nota_firma(3) = l_hist_nota_firma(3) + l_hist_nota_firma(2) + l_hist_nota_firma(1)
    l_hist_assalto(3) = l_hist_assalto(3) + l_hist_assalto(2) + l_hist_assalto(1)
    l_hist_total(3) = l_hist_total(3) + l_hist_total(2) + l_hist_total(1)
    Printer.Print "+--+----------+----------+---------+----+---------+----------------------------+"
    Printer.Print "|  HISTÓRICO DO CAIXA DE ÓLEO/LUBRIF.   |   HISTÓRICO DOS CAIXAS UNIFICADOS    |"
    Printer.Print "+---------------------------------------+--------------------------------------+"
    x_linha = "| Cheque Pré-Datado....:                | Tot. Cheques Pré-Dat:                |"
    If l_hist_ch_predatado(2) > 0 Then
        i = Len(Format(l_hist_ch_predatado(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_ch_predatado(2), "#,###,##0.00")
    End If
    If l_hist_ch_predatado(3) > 0 Then
        i = Len(Format(l_hist_ch_predatado(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_ch_predatado(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| Cheques à Vista......:                | Tot. Cheques à Vista:                |"
    If l_hist_ch_vista(2) > 0 Then
        i = Len(Format(l_hist_ch_vista(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_ch_vista(2), "#,###,##0.00")
    End If
    If l_hist_ch_vista(3) > 0 Then
        i = Len(Format(l_hist_ch_vista(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_ch_vista(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| Dinheiro.............:                | Tot. Dinheiro.......:                |"
    If l_hist_dinheiro(2) > 0 Then
        i = Len(Format(l_hist_dinheiro(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_dinheiro(2), "#,###,##0.00")
    End If
    If l_hist_dinheiro(3) > 0 Then
        i = Len(Format(l_hist_dinheiro(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_dinheiro(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| Notas de Firmas......:                | Tot. Notas de Firmas:                |"
    If l_hist_nota_firma(2) > 0 Then
        i = Len(Format(l_hist_nota_firma(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_nota_firma(2), "#,###,##0.00")
    End If
    If l_hist_nota_firma(3) > 0 Then
        i = Len(Format(l_hist_nota_firma(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_nota_firma(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| Assalto..............:                | Tot. Assalto........:                |"
    If l_hist_assalto(2) > 0 Then
        i = Len(Format(l_hist_assalto(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_assalto(2), "#,###,##0.00")
    End If
    If l_hist_assalto(3) > 0 Then
        i = Len(Format(l_hist_assalto(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_assalto(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| Total................:                | Total Geral.........:                |"
    If l_hist_total(2) > 0 Then
        i = Len(Format(l_hist_total(2), "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(l_hist_total(2), "#,###,##0.00")
    End If
    If l_hist_total(3) > 0 Then
        i = Len(Format(l_hist_total(3), "#,###,##0.00"))
        Mid(x_linha, 68 + 12 - i, i) = Format(l_hist_total(3), "#,###,##0.00")
    End If
    Printer.Print x_linha
End Sub
Private Sub ImpResumoCombustiveisBomba()
    Dim x_linha As String
    Dim i As Integer
    Printer.Print "+-----------+-------------+-------------+------------+------------+------------+"
    Printer.Print "|COMBUSTÍVEL|    LITROS   |    VALOR    |EST. MEDIÇÃO| EST. EM R$ |  DIFERENCA |"
    Printer.Print "+-----------+-------------+-------------+------------+------------+------------+"
    Call ImpDetCombustivelBomba("ÁLCOOL    ", l_litros_a, l_valor_a)
    Call ImpDetCombustivelBomba("ÁLCOOL +  ", l_litros_aa, l_valor_aa)
    Call ImpDetCombustivelBomba("DIESEL    ", l_litros_d, l_valor_d)
    Call ImpDetCombustivelBomba("DIESEL +  ", l_litros_da, l_valor_da)
    Call ImpDetCombustivelBomba("GASOLINA  ", l_litros_g, l_valor_g)
    Call ImpDetCombustivelBomba("GASOLINA +", l_litros_ga, l_valor_ga)
    Call ImpDetCombustivelBomba("ÓLEOS/LUBR", l_litros_l, l_valor_l)
    Call ImpDetCombustivelBomba("BORRA/LAV.", l_litros_b, l_valor_b)
    Printer.Print "+-----------+-------------+-------------+------------+------------+------------+"
    x_linha = "| ** TOTAL  |             |             | DIFERENÇA DE CAIXA...:               |"
    i = Len(Format(l_litros_t, "##,###,##0.0"))
    Mid(x_linha, 15 + 12 - i, i) = Format(l_litros_t, "##,###,##0.0")
    i = Len(Format(l_valor_t, "#,###,##0.00"))
    Mid(x_linha, 28 + 12 - i, i) = Format(l_valor_t, "#,###,##0.00")
    l_dif_caixa(1) = l_hist_total(1) - l_valor_t - l_valor_afericao
    If l_dif_caixa(1) <> 0 Then
        i = Len(Format(l_dif_caixa(1), "#,###,##0.00;####,##0.00-"))
        Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(1), "#,###,##0.00;####,##0.00-")
    End If
    Printer.Print x_linha
    x_linha = "|           |             |             | DIFERENÇA CAIXA ÓLEOS:               |"
    l_dif_caixa(2) = l_hist_total(2) - l_valor_l
    If l_dif_caixa(2) <> 0 Then
        i = Len(Format(l_dif_caixa(2), "#,###,##0.00;####,##0.00-"))
        Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(2), "#,###,##0.00;####,##0.00-")
    End If
    Printer.Print x_linha
    x_linha = "|           |             |             | DIFERENÇA CAIXA BORR.:               |"
    l_hist_total(3) = l_hist_total(3) - l_hist_total(2) - l_hist_total(1)
    l_dif_caixa(3) = l_hist_total(3) - l_valor_b
    If l_dif_caixa(3) <> 0 Then
        i = Len(Format(l_dif_caixa(3), "#,###,##0.00;####,##0.00-"))
        Mid(x_linha, 67 + 12 - i, i) = Format(l_dif_caixa(3), "#,###,##0.00;####,##0.00-")
    End If
    Printer.Print x_linha
    Printer.Print "+-----------+-------------+-------------+--------------------------------------+"
End Sub
Private Sub ImpHistoricoLubrificante()
    Dim x_linha As String
    Dim i As Integer
    l_dife_caixa = l_his_total - l_valor
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 61, 1) = "+"
    Mid(x_linha, 80, 1) = "+"
    Mid(x_linha, 96, 1) = "+"
    Mid(x_linha, 115, 1) = "+"
    Mid(x_linha, 137, 1) = "+"
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
    Printer.Print "|                             RESUMO DO CAIXA                                  |"
    Printer.Print "+--------------------+------------------+--------------------------------------+"
    Printer.Print "| NOME DA OPERACAO   |     V A L O R    |                                      |"
    Printer.Print "+--------------------+------------------+--------------------------------------+"
    x_linha = "| CHEQUE PRÉ-DATADO. |                  |                                      |"
    If l_his_ch_predatado > 0 Then
        i = Len(Format(l_his_ch_predatado, "##,###,##0.00"))
        Mid(x_linha, 27 + 13 - i, i) = Format(l_his_ch_predatado, "##,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| CHEQUE À VISTA.... |                  |                                      |"
    If l_his_ch_vista > 0 Then
        i = Len(Format(l_his_ch_vista, "##,###,##0.00"))
        Mid(x_linha, 27 + 13 - i, i) = Format(l_his_ch_vista, "##,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| DINHEIRO.......... |                  |                                      |"
    If l_his_dinheiro > 0 Then
        i = Len(Format(l_his_dinheiro, "##,###,##0.00"))
        Mid(x_linha, 27 + 13 - i, i) = Format(l_his_dinheiro, "##,###,##0.00")
    End If
    If l_dife_caixa <> 0 Then
        Mid(x_linha, 45, 20) = "DIFERENÇA DE CAIXA.:"
        i = Len(Format(l_dife_caixa, "###,###,##0.00;##,###,##0.00-"))
        Mid(x_linha, 65 + 14 - i, i) = Format(l_dife_caixa, "###,###,##0.00;##,###,##0.00-")
    End If
    Printer.Print x_linha
    x_linha = "| NOTA FIRMA........ |                  |                                      |"
    If l_his_nota_firma > 0 Then
        i = Len(Format(l_his_nota_firma, "##,###,##0.00"))
        Mid(x_linha, 27 + 13 - i, i) = Format(l_his_nota_firma, "##,###,##0.00")
    End If
    Printer.Print x_linha
    x_linha = "| TOTAL............. |                  | RESPONSÁVEL:                         |"
    If l_his_total > 0 Then
        i = Len(Format(l_his_total, "##,###,##0.00"))
        Mid(x_linha, 27 + 13 - i, i) = Format(l_his_total, "##,###,##0.00")
    End If
    Mid(x_linha, 56, 24) = l_nome_funcionario
    Printer.Print x_linha
    x_linha = "+--------------------+------------------+--------------------------------------+"
    Printer.Print x_linha
    Printer.FontName = "Draft 10cpi"
    Printer.Print " "
End Sub
Private Sub ImpProdutoNota()
    Dim x_linha As String
    Dim x_nome_produto As String
    Dim i As Integer
    x_linha = Space(137)
    With tbl_movimento_nota
        Mid(x_linha, 1, 1) = "|"
        Mid(x_linha, 2, 10) = Format(![Data do Abastecimento], "dd/mm/yyyy")
        tbl_produto.Seek "=", ![Codigo do Produto2]
        If Not tbl_produto.NoMatch Then
            x_nome_produto = tbl_produto!Nome
        Else
            x_nome_produto = "** Não Cadastrado **"
        End If
        Mid(x_linha, 12, 1) = "|"
        i = Len(Format(![Codigo do Produto2], "#000"))
        Mid(x_linha, 16 + 4 - i, i) = Format(![Codigo do Produto2], "#000")
        Mid(x_linha, 21, 1) = "|"
        Mid(x_linha, 23, 40) = x_nome_produto
        Mid(x_linha, 64, 1) = "|"
        i = Len(Format(!Quantidade, "####,##0.00"))
        Mid(x_linha, 65 + 11 - i, i) = Format(!Quantidade, "####,##0.00")
        Mid(x_linha, 77, 1) = "|"
        Mid(x_linha, 87, 1) = "|"
        i = Len(Format(![Valor Total], "###,###,##0.00"))
        Mid(x_linha, 91 + 14 - i, i) = Format(![Valor Total], "###,###,##0.00")
        Mid(x_linha, 106, 1) = "|"
        Mid(x_linha, 137, 1) = "|"
    End With
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpResumoChequePreDatadoBomba()
    Dim x_linha As String
    Dim i As Integer
    Dim x_total As Currency
    x_total = MovCheque.TotalEmissaoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), l_periodo_i, l_periodo_f, "0", "P")
    If x_total > 0 Then
        x_linha = "| CHEQUES PRÉ-DATADOS PARA VENCIMENTO EM.:             TOTAL.:                 |"
        Mid(x_linha, 44, 10) = Format(CDate(msk_data_i) + 1, "dd/mm/yyyy")
        i = Len(Format(x_total, "###,###,##0.00"))
        Mid(x_linha, 64 + 14 - i, i) = Format(x_total, "###,###,##0.00")
        Printer.Print x_linha
        x_linha = "+------------------------------------------------------------------------------+"
        If g_usuario = 8 Then
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
        End If
        Printer.Print x_linha
    End If
End Sub
Private Sub ImpDetCombustivelBomba(x_combustivel As String, x_litros As Currency, x_valor As Currency)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|           |             |             |            |            |            |"
    Mid(x_linha, 3, 10) = x_combustivel
    If CCur(x_litros) > 0 Then
        i = Len(Format(x_litros, "##,###,##0.0"))
        Mid(x_linha, 15 + 12 - i, i) = Format(x_litros, "##,###,##0.0")
    End If
    If CCur(x_valor) > 0 Then
        i = Len(Format(x_valor, "#,###,##0.00"))
        Mid(x_linha, 28 + 12 - i, i) = Format(x_valor, "#,###,##0.00")
    End If
    Printer.Print x_linha
End Sub
Private Sub ImpDetLubrificante(x_codigo_produto As Long, x_nome As String, x_unidade As String, x_valor_venda As Currency, x_quantidade As Currency, x_valor_total As Currency, x_data As Date, x_periodo As String, x_codigo_funcionario As Integer)
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 12, 1) = "|"
    Mid(x_linha, 55, 1) = "|"
    Mid(x_linha, 61, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    Mid(x_linha, 96, 1) = "|"
    Mid(x_linha, 115, 1) = "|"
    Mid(x_linha, 126, 1) = "|"
    Mid(x_linha, 130, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    i = Len(Format(x_codigo_produto, "#000"))
    Mid(x_linha, 5 + 4 - i, i) = Format(x_codigo_produto, "#000")
    Mid(x_linha, 14, 40) = x_nome
    Mid(x_linha, 57, 3) = x_unidade
    i = Len(Format(x_valor_venda, "##,###,##0.00"))
    Mid(x_linha, 66 + 13 - i, i) = Format(x_valor_venda, "##,###,##0.00")
    i = Len(Format(x_quantidade, "####,##0.00"))
    Mid(x_linha, 84 + 11 - i, i) = Format(x_quantidade, "####,##0.00")
    i = Len(Format(x_valor_total, "##,###,##0.00"))
    Mid(x_linha, 101 + 13 - i, i) = Format(x_valor_total, "##,###,##0.00")
    Mid(x_linha, 116, 10) = Format(x_data, "dd/mm/yyyy")
    Mid(x_linha, 128, 1) = x_periodo
    i = Len(Format(x_codigo_funcionario, "#00"))
    Mid(x_linha, 132 + 3 - i, i) = Format(x_codigo_funcionario, "#00")
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetBomba(x_i As Integer, x_abertura As Currency, x_encerrante As Currency, x_litros As Currency, x_valor As Currency, x_historico As String, x_variavel As String)
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(80)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 4, 1) = "|"
    Mid(x_linha, 15, 1) = "|"
    Mid(x_linha, 26, 1) = "|"
    Mid(x_linha, 36, 1) = "|"
    Mid(x_linha, 51, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    If x_i > 0 Then
        Mid(x_linha, 2, 2) = Format(x_i, "00")
    End If
    If CCur(x_abertura) Or CCur(x_encerrante) > 0 Then
        i = Len(Format(x_abertura, "####,##0.0"))
        Mid(x_linha, 5 + 10 - i, i) = Format(x_abertura, "####,##0.0")
        i = Len(Format(x_encerrante, "####,##0.0"))
        Mid(x_linha, 16 + 10 - i, i) = Format(x_encerrante, "####,##0.0")
        i = Len(Format(x_litros, "###,##0.0"))
        Mid(x_linha, 27 + 9 - i, i) = Format(x_litros, "###,##0.0")
        i = Len(Format(x_valor, "##,###,##0.00"))
        Mid(x_linha, 37 + 13 - i, i) = Format(x_valor, "##,###,##0.00")
    End If
    Mid(x_linha, 52, 13) = x_historico
    If Mid(x_variavel, 1, 3) = "@N@" Then
        If CCur(Mid(x_variavel, 4, Len(x_variavel) - 3)) > 0 Then
            x_variavel = Mid(x_variavel, 4, Len(x_variavel) - 3)
            i = Len(Format(x_variavel, "##,###,##0.00"))
            Mid(x_linha, 66 + 13 - i, i) = Format(x_variavel, "##,###,##0.00")
        End If
    Else
        Mid(x_linha, 65, 15) = Mid(x_variavel, 4, Len(x_variavel) - 3)
    End If
    Printer.FontName = "Draft 10cpi"
    Printer.Print x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetCartao(x_data_emissao As Date, x_periodo As String, x_numero_lancamento As Integer, x_nome_cartao As String, x_data_vencimento As Date, x_valor As Currency, x_numero_cartao As Integer)
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 2, 10) = Format(x_data_emissao, "dd/mm/yyyy")
    Mid(x_linha, 12, 1) = "|"
    Mid(x_linha, 16, 1) = x_periodo
    Mid(x_linha, 20, 1) = "|"
    Mid(x_linha, 22, 4) = Format(x_numero_lancamento, "0000")
    Mid(x_linha, 27, 1) = "|"
    Mid(x_linha, 29, 13) = x_nome_cartao
    Mid(x_linha, 41, 1) = "|"
    Mid(x_linha, 42, 10) = Format(x_data_vencimento, "dd/mm/yyyy")
    Mid(x_linha, 52, 1) = "|"
    i = Len(Format(x_valor, "#,###,##0.00"))
    Mid(x_linha, 57 + 12 - i, i) = Format(x_valor, "#,###,##0.00")
    Mid(x_linha, 70, 1) = "|"
    Mid(x_linha, 72, 4) = Format(x_numero_cartao, "0000")
    Mid(x_linha, 93, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetCheque()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|            |       |           |          |     |               |            |                                          |             |"
    Mid(x_linha, 3, 10) = Format(rsCheque("Data de Emissao").Value, "dd/mm/yyyy")
    Mid(x_linha, 18, 1) = rsCheque("Periodo").Value
    i = Len(Format(rsCheque("Numero da Conta").Value, "##########"))
    Mid(x_linha, 24 + 10 - i, i) = Format(rsCheque("Numero da Conta").Value, "##########")
    i = Len(Format(rsCheque("Numero do Cheque").Value, "######"))
    Mid(x_linha, 38 + 6 - i, i) = Format(rsCheque("Numero do Cheque").Value, "######")
    i = Len(Format(DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value), "##"))
    Mid(x_linha, 48 + 2 - i, i) = Format(DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value, "##"))
    i = Len(Format(rsCheque("Valor").Value, "###,###,##0.00"))
    Mid(x_linha, 52 + 14 - i, i) = Format(rsCheque("Valor").Value, "###,###,##0.00")
    Mid(x_linha, 69, 10) = Format(rsCheque("Data do Vencimento").Value, "dd/mm/yyyy")
    Mid(x_linha, 82, 40) = rsCheque("Emitente").Value
    If rsCheque("Tipo do Movimento").Value = "1" Then
        Mid(x_linha, 125, 12) = "Combustível "
    ElseIf rsCheque("Tipo do Movimento").Value = "2" Then
        Mid(x_linha, 125, 12) = "Oleo/Diverso"
    ElseIf rsCheque("Tipo do Movimento").Value = "3" Then
        Mid(x_linha, 125, 12) = "Borr.Lavador"
    End If
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCabBomba()
    Dim x_linha As String
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        Printer.PaperSize = 1
        Printer.FontName = "Draft 10cpi"
        Printer.FontName = "Draft 10cpi"
        'teste para imprimir letra correta
        Printer.FontBold = False
        ImprimeTexto "  ", 1, 2, 2, 1
    End If
    lPagina = lPagina + 1
    lLinha = 0
    Printer.FontName = "Draft 5cpi"
    Printer.FontName = "Draft 10cpi"
    Printer.CurrentY = 0
    Printer.Print "+------------------------------------------------------------------------------+"
    Printer.FontBold = True
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Printer.Print x_linha
    Printer.FontBold = False
    Printer.Print "| MOVIMENTO DAS BOMBAS                                     Goiânia, " & msk_data & " |"
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    Printer.Print x_linha
    x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_linha, 29, 1) = l_periodo_i
    Mid(x_linha, 49, 1) = l_periodo_f
    Printer.Print x_linha
    Printer.Print "+--+----------+----------+---------+--------------+----------------------------+"
    Printer.Print "|N.| ABERTURA |ENCERRANTE|LTS.SAIDA|VALOR DA SAIDA| HISTORICO DO CAIXA         |"
    Printer.Print "+--+----------+----------+---------+--------------+----------------------------+"
End Sub
Private Sub ImpCabCartao()
    Dim x_string_137 As String
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        Printer.PaperSize = 1
        Printer.FontName = "Draft 10cpi"
        Printer.FontName = "Draft 10cpi"
        'teste para imprimir letra correta
        Printer.FontBold = False
        ImprimeTexto "  ", 1, 2, 2, 1
    End If
    lPagina = lPagina + 1
    lLinha = 0
    Printer.FontName = "Draft 5cpi"
    Printer.FontName = "Draft 10cpi"
    Printer.CurrentY = 0
    Printer.Print "+------------------------------------------------------------------------------+"
    Printer.FontBold = True
    x_string_137 = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_string_137, 3, 40) = g_nome_empresa
    Printer.Print x_string_137
    Printer.FontBold = False
    Printer.Print "| RELAÇÃO DO MOVIMENTO DE CARTÃO DE CRÉDITO                Goiânia, " & msk_data & " |"
    x_string_137 = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_string_137, 29, 10) = msk_data_i
    Mid(x_string_137, 42, 10) = msk_data_f
    Printer.Print x_string_137
    x_string_137 = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_string_137, 29, 1) = l_periodo_i
    Mid(x_string_137, 49, 1) = l_periodo_f
    Printer.Print x_string_137
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print "+----------+-------+------+-------------+----------+-----------------+----------------------+-------------------------------------------+"
    Printer.Print "| EMISSÃO  |PERIODO|N.LANC|    CARTÃO   |VENCIMENTO| VALOR DO CARTÃO | NÚMERO DO CARTÃO     |                                           |"
    Printer.Print "+----------+-------+------+-------------+----------+-----------------+----------------------+-------------------------------------------+"
End Sub
Private Sub ImpCabCheque(x_tipo_movimento As Integer)
    Dim x_linha As String
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        Printer.PaperSize = 1
        Printer.FontName = "Draft 10cpi"
        Printer.FontName = "Draft 10cpi"
        'teste para imprimir letra correta
        Printer.FontBold = False
        ImprimeTexto "  ", 1, 2, 2, 1
    End If
    Dim x_string_40 As String * 40
    lPagina = lPagina + 1
    lLinha = 0
    Printer.FontName = "Draft 5cpi"
    Printer.FontName = "Draft 10cpi"
    Printer.CurrentY = 0
    Printer.Print "+------------------------------------------------------------------------------+"
    x_string_40 = g_nome_empresa
    Printer.Print "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    Printer.Print "| RELAÇÃO DOS CHEQUES                                      Goiânia, " & msk_data & " |"
    If x_tipo_movimento = 0 Then
        x_string_40 = "Todos os Caixas"
    ElseIf x_tipo_movimento = 1 Then
        x_string_40 = "Caixa de combustíveis"
    ElseIf x_tipo_movimento = 1 Then
        x_string_40 = "Caixa de óleo/diversos"
    End If
    Printer.Print "| Tipo de Movimento.: " & x_string_40 & "                 |"
    x_linha = "| Referente a.: " & msk_data_i & " a " & msk_data_f & "       Período   ao                     |"
    Mid(x_linha, 55, 1) = l_periodo_i
    If x_tipo_movimento = 2 Then
        Mid(x_linha, 55, 1) = l_periodo_f
    End If
    Mid(x_linha, 60, 1) = l_periodo_f
    Printer.Print x_linha
    Printer.FontName = "Sans Serif 17cpi"
    Printer.FontBold = False
    Printer.Print "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    Printer.Print "|   EMISSÃO  |PERIODO|N. DA CONTA| N.CHEQUE |PRAZO|VALOR DO CHEQUE| VENCIMENTO | NOME DO EMITENTE                         | TIPO   MOV. |"
    Printer.Print "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
End Sub
Private Sub ImpCabLubrificante(x_tipo_movimento As Integer)
    Dim x_linha As String
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        Printer.PaperSize = 1
        Printer.FontName = "Draft 10cpi"
        Printer.FontName = "Draft 10cpi"
        'teste para imprimir letra correta
        Printer.FontBold = False
        ImprimeTexto "  ", 1, 2, 2, 1
    End If
    lPagina = lPagina + 1
    lLinha = 0
    Printer.FontName = "Draft 5cpi"
    Printer.FontName = "Draft 10cpi"
    Printer.CurrentY = 0
    Printer.Print "+------------------------------------------------------------------------------+"
    Printer.FontBold = True
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Printer.Print x_linha
    Printer.FontBold = False
    Printer.Print "| VENDAS DE PRODUTOS                                       Goiânia, " & msk_data & " |"
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    Printer.Print x_linha
    x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_linha, 29, 1) = l_periodo_i
    Mid(x_linha, 49, 1) = l_periodo_f
    Printer.Print x_linha
    x_linha = "| FUNCIONARIO.............:                                                    |"
    Mid(x_linha, 29, 3) = "000"
    Mid(x_linha, 33, 40) = "Todos os Funcionários"
    Printer.Print x_linha
    x_linha = "| GRUPO...................:                                                    |"
    Mid(x_linha, 29, 40) = "Todos os Grupos"
    Printer.Print x_linha
    x_linha = "| PRODUTO.................:                                                    |"
    Mid(x_linha, 29, 40) = "Todos os Grupos"
    Printer.Print x_linha
    x_linha = "| TIPO DO MOVIMENTO.......:                                                    |"
    If x_tipo_movimento = 0 Then
        Mid(x_linha, 29, 40) = "Todos os Caixas"
    ElseIf x_tipo_movimento = 1 Then
        Mid(x_linha, 29, 40) = "Caixa de combustíveis"
    ElseIf x_tipo_movimento = 1 Then
        Mid(x_linha, 29, 40) = "Caixa de óleo/diversos"
    End If
    Printer.Print x_linha
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print "+----------+------------------------------------------+-----+------------------+---------------+------------------+----------+---+------+"
    Printer.Print "|  CODIGO  | DISCRIMINACAO DOS PRODUTOS               | UN  |  PRECO DE VENDA  |   QUANTIDADE  |  TOTAL DE VENDA  |DATA SAIDA|PER| FUNC.|"
    Printer.Print "+----------+------------------------------------------+-----+------------------+---------------+------------------+----------+---+------+"
End Sub
Private Sub ImpCabNota(x_tipo_movimento As Integer)
    Dim x_string_40 As String * 40
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        Printer.ScaleMode = 7
        Printer.PaperSize = 1
        Printer.FontName = "Draft 10cpi"
        Printer.FontName = "Draft 10cpi"
        'teste para imprimir letra correta
        Printer.FontBold = False
        ImprimeTexto "  ", 1, 2, 2, 1
    End If
    lPagina = lPagina + 1
    lLinha = 0
    Printer.FontName = "Draft 5cpi"
    Printer.FontName = "Draft 10cpi"
    Printer.CurrentY = 0
    Printer.Print "+------------------------------------------------------------------------------+"
    x_string_40 = g_nome_empresa
    Printer.Print "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    Printer.Print "| RELAÇÃO DAS NOTAS DE ABASTECIMENTO POR EMISSÃO           Goiânia, " & msk_data & " |"
    If x_tipo_movimento = 0 Then
        x_string_40 = "Todos os Caixas"
    ElseIf x_tipo_movimento = 1 Then
        x_string_40 = "Caixa de combustíveis"
    ElseIf x_tipo_movimento = 1 Then
        x_string_40 = "Caixa de óleo/diversos"
    End If
    Printer.Print "| Tipo de Movimento.: " & x_string_40 & "                 |"
    Printer.Print "| Referente a.: " & msk_data_i & " a " & msk_data_f & "       Período " & l_periodo_i & " ao " & l_periodo_f & "                   |"
    Printer.FontName = "Sans Serif 17cpi"
    Printer.FontBold = False
    Printer.FontName = "Sans Serif 17cpi"
    Printer.Print "+----------+--------+------------------------------------------+------------+---------+------------------+------------------------------+"
    Printer.Print "|   DATA   | CÓDIGO | DESCRIÇÃO DOS PRODUTOS                   | QUANTIDADE |         | VALOR DO PRODUTO |                              |"
    Printer.Print "+----------+--------+------------------------------------------+------------+---------+------------------+------------------------------+"
End Sub
Private Sub ImpClienteNota()
    Dim x_linha As String
    Dim i As Integer
    If lSubTotal > 0 Then
        ImpSubTotalNota
    End If
    x_linha = Space(137)
    With tbl_movimento_nota
        Mid(x_linha, 1, 12) = "| Cliente..:"
        i = Len(Format(![Codigo do Cliente], "#####"))
        Mid(x_linha, 15 + 5 - i, i) = Format(![Codigo do Cliente], "#####")
        tbl_cliente.Seek "=", ![Codigo do Cliente]
        If Not tbl_cliente.NoMatch Then
            Mid(x_linha, 22, 37) = tbl_cliente![Razao Social]
        Else
            Mid(x_linha, 22, 37) = "** Não Cadastrado **"
        End If
        Mid(x_linha, 60, 8) = "Pedido.:"
        i = Len(Format(![Numero da Nota], "####,###"))
        Mid(x_linha, 68 + 8 - i, i) = Format(![Numero da Nota], "####,###")
        Mid(x_linha, 77, 1) = "|"
        i = Len(Format(![Codigo do Conveniado], "###,###"))
        Mid(x_linha, 92 + 7 - i, i) = Format(![Codigo do Conveniado], "###,###")
        Mid(x_linha, 100, 37) = Space(37)
        If tbl_movimento_nota![Codigo do Conveniado] > 0 Then
            Mid(x_linha, 79, 12) = "Conveniado.:"
            tbl_cliente_conveniado.Seek "=", ![Codigo do Cliente], tbl_movimento_nota![Codigo do Conveniado]
            If Not tbl_cliente_conveniado.NoMatch Then
                Mid(x_linha, 100, 37) = tbl_cliente_conveniado!Nome
            End If
        End If
        Mid(x_linha, 137, 1) = "|"
        Printer.FontName = "Sans Serif 17cpi"
        Printer.Print x_linha
        l_cliente = ![Codigo do Cliente]
        l_conveniado = ![Codigo do Conveniado]
        l_numero_nota = ![Numero da Nota]
    End With
    lLinha = lLinha + 1
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorios
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
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
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
    
    Set tbl_bomba = bd_sgp.OpenTable("Bomba")
    Set tbl_cartao_credito = bd_sgp.OpenTable("Cartao_Credito")
    Set tbl_cliente = bd_sgp.OpenTable("Cliente")
    Set tbl_cliente_conveniado = bd_sgp.OpenTable("Cliente_Conveniado")
    Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_afericao = bd_sgp.OpenTable("Movimento_Afericao")
    Set tbl_movimento_bomba = bd_sgp.OpenTable("Movimento_Bomba")
    Set tbl_movimento_cartao_credito = bd_sgp.OpenTable("Movimento_Cartao_Credito")
    Set tbl_movimento_historico = bd_sgp.OpenTable("Movimento_Historico")
    Set tbl_movimento_lubrificante = bd_sgp.OpenTable("Movimento_Lubrificante")
    Set tbl_movimento_nota = bd_sgp.OpenTable("Movimento_Nota_Abastecimento")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    Set tbl_tabela_premiacao = bd_sgp.OpenTable("Tabela_Premiacao")
    tbl_bomba.Index = "id_codigo"
    tbl_cartao_credito.Index = "id_codigo"
    tbl_cliente.Index = "id_codigo"
    tbl_cliente_conveniado.Index = "id_codigo"
    tbl_combustivel.Index = "id_codigo"
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_afericao.Index = "id_data"
    tbl_movimento_bomba.Index = "id_data"
    tbl_movimento_cartao_credito.Index = "id_data_emissao"
    tbl_movimento_historico.Index = "id_data"
    tbl_movimento_lubrificante.Index = "id_data"
    tbl_movimento_nota.Index = "id_data_abastecimento"
    tbl_produto.Index = "id_codigo"
    tbl_tabela_premiacao.Index = "id_mes_ano"
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
        cmd_imprimir.SetFocus
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
