VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_resumo_movimentacao_postos 
   Caption         =   "Emissão do Resumo da Movimentação dos Postos"
   ClientHeight    =   2715
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6675
   Icon            =   "lst_resumo_movimentacao_postos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   6675
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4740
      Picture         =   "lst_resumo_movimentacao_postos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2940
      Picture         =   "lst_resumo_movimentacao_postos.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime análise do resumo da movimentação dos postos."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_resumo_movimentacao_postos.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza análise do resumo da movimentação dos postos."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_resumo_movimentacao_postos.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_resumo_movimentacao_postos.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5820
         Picture         =   "lst_resumo_movimentacao_postos.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
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
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4680
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
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3660
         TabIndex        =   12
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3660
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
Attribute VB_Name = "emissao_resumo_movimentacao_postos"
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
Dim l_cheque_predatado(1 To 12) As Currency
Dim l_amex(1 To 12) As Currency
Dim l_credicard(1 To 12) As Currency
Dim l_visa(1 To 12) As Currency
Dim l_hipercheque(1 To 12) As Currency
Dim l_nota(1 To 12) As Currency
Dim l_cheque_avista(1 To 12) As Currency
Dim l_dinheiro(1 To 12) As Currency
Dim l_total(1 To 12) As Currency
Dim l_percentual(1 To 4) As Currency
Dim lSQl As String

Dim tbl_baixa_cheque As Table
Dim tbl_cartao_credito As Table
Dim tbl_duplicata_receber As Table
Dim tbl_empresa As Table
Dim tbl_movimento_cartao_credito As Table
Dim tbl_movimento_cheque_avista As Table
Dim tbl_movimento_nota_abastecimento As Table


Private rsMovComposicaoCaixa As New adodb.Recordset
Private MovCheque As New cMovimentoCheque
Private Sub BuscaPercentualCartoes()
    Dim i As Integer
    For i = 1 To 4
        tbl_cartao_credito.Seek "=", i
        If Not tbl_cartao_credito.NoMatch Then
            l_percentual(i) = tbl_cartao_credito![Taxa de Custo]
        End If
    Next
End Sub
Function CalculaCartao(x_empresa As Integer)
    With tbl_movimento_cartao_credito
        .Index = "id_data_emissao"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i), cbo_periodo_i, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        If ![Codigo do Cartao] = 1 Then
                            l_visa(x_empresa) = l_visa(x_empresa) + !valor
                            l_visa(12) = l_visa(12) + !valor
                        ElseIf ![Codigo do Cartao] = 2 Then
                            l_credicard(x_empresa) = l_credicard(x_empresa) + !valor
                            l_credicard(12) = l_credicard(12) + !valor
                        ElseIf ![Codigo do Cartao] = 3 Then
                            l_amex(x_empresa) = l_amex(x_empresa) + !valor
                            l_amex(12) = l_amex(12) + !valor
                        ElseIf ![Codigo do Cartao] = 4 Then
                            l_hipercheque(x_empresa) = l_hipercheque(x_empresa) + !valor
                            l_hipercheque(12) = l_hipercheque(12) + !valor
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function CalculaCartaoCreditos(x_empresa As Integer)
    Dim x_valor As Currency
    With tbl_movimento_cartao_credito
        .Index = "id_data_vencimento2"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i), "0", 0, CDate("01/01/1900")
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data do Vencimento] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        x_valor = Format(!valor - !valor * l_percentual(![Codigo do Cartao]) / 100, "00000000.00")
                        If ![Codigo do Cartao] = 1 Then
                            l_visa(x_empresa) = l_visa(x_empresa) + x_valor
                            l_visa(12) = l_visa(12) + x_valor
                        ElseIf ![Codigo do Cartao] = 2 Then
                            l_credicard(x_empresa) = l_credicard(x_empresa) + x_valor
                            l_credicard(12) = l_credicard(12) + x_valor
                        ElseIf ![Codigo do Cartao] = 3 Then
                            l_amex(x_empresa) = l_amex(x_empresa) + x_valor
                            l_amex(12) = l_amex(12) + x_valor
                        ElseIf ![Codigo do Cartao] = 4 Then
                            l_hipercheque(x_empresa) = l_hipercheque(x_empresa) + x_valor
                            l_hipercheque(12) = l_hipercheque(12) + x_valor
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function CalculaChequeAvista(x_empresa As Integer)
    Dim xValor As Currency
    
    With tbl_movimento_cheque_avista
        .Index = "id_digitacao"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i.Text), cbo_periodo_i.Text, "0", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        l_cheque_avista(x_empresa) = l_cheque_avista(x_empresa) + !valor
                        l_cheque_avista(12) = l_cheque_avista(12) + !valor
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    
    xValor = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), cbo_periodo_i.Text, cbo_periodo_f.Text, "0", "P")
    l_cheque_avista(x_empresa) = l_cheque_avista(x_empresa) + xValor
    l_cheque_avista(12) = l_cheque_avista(12) + l_cheque_avista(x_empresa)
End Function
Function CalculaChequePreDatado(x_empresa As Integer)
    l_cheque_predatado(x_empresa) = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), cbo_periodo_i.Text, cbo_periodo_f.Text, "0", "P")
    l_cheque_predatado(12) = l_cheque_predatado(12) + l_cheque_predatado(x_empresa)
End Function
Function CalculaChequePreDatadoCreditos(x_empresa As Integer)
    l_cheque_predatado(x_empresa) = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), cbo_periodo_i.Text, cbo_periodo_f.Text, "0", "P")
    l_cheque_predatado(12) = l_cheque_predatado(12) + l_cheque_predatado(x_empresa)
End Function
Function CalculaChequePreDatadoBaixadoCreditos(x_empresa As Integer)
    With tbl_baixa_cheque
        .Index = "id_vencimento"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i), CDate("01/01/1900"), " ", 0, "          ", "      "
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data do Vencimento] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        l_cheque_predatado(x_empresa) = l_cheque_predatado(x_empresa) + !valor
                        l_cheque_predatado(12) = l_cheque_predatado(12) + !valor
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function CalculaDinheiro(x_empresa As Integer)
    Dim xCodigo As Integer
    
    xCodigo = 1
    
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "   SELECT Codigo, Nome"
    lSQl = lSQl & "     FROM Composicao_Caixa"
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQl)
    
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        rsMovComposicaoCaixa.MoveFirst
        Do Until rsMovComposicaoCaixa.EOF
            If UCase(rsMovComposicaoCaixa("Nome").Value) Like "*DINHEIRO*" Then
                xCodigo = rsMovComposicaoCaixa("Codigo").Value
                Exit Do
            End If
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    
    'loop movimento do historico
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "   SELECT SUM(Valor) AS Total"
    lSQl = lSQl & "     FROM Movimento_Composicao_Caixa"
    lSQl = lSQl & "    WHERE Movimento_Composicao_Caixa.Empresa = " & x_empresa
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Data >= #" & Format(msk_data_i.Text, "mm/dd/yyyy") & "#"
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Data <= #" & Format(msk_data_f.Text, "mm/dd/yyyy") & "#"
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Periodo >= " & Val(cbo_periodo_i.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.Periodo <= " & Val(cbo_periodo_f.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] >= " & 1 'Val(txt_ilha_i.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.[Numero da Ilha] <= " & 1 'Val(txt_ilha_f.Text)
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.[Tipo do Movimento] >= " & 1
    lSQl = lSQl & "      AND Movimento_Composicao_Caixa.[Codigo da Composicao] = " & xCodigo
    'Abre RecordSet
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(lSQl)
    If rsMovComposicaoCaixa.RecordCount > 0 Then
        rsMovComposicaoCaixa.MoveFirst
        If Not IsNull(rsMovComposicaoCaixa("Total").Value) Then
            l_dinheiro(x_empresa) = rsMovComposicaoCaixa("Total").Value
            l_dinheiro(12) = rsMovComposicaoCaixa("Total").Value
        End If
    End If
End Function
Function CalculaDuplicataCreditos(x_empresa As Integer)
    With tbl_duplicata_receber
        .Index = "id_empresa_vencimento"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i), 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data do Vencimento] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    l_nota(x_empresa) = l_nota(x_empresa) + ![Valor do Vencimento]
                    l_nota(12) = l_nota(12) + ![Valor do Vencimento]
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function CalculaNota(x_empresa As Integer)
    With tbl_movimento_nota_abastecimento
        .Index = "id_data_abastecimento"
        If .RecordCount > 0 Then
            .Seek ">", x_empresa, CDate(msk_data_i), cbo_periodo_i, 0, 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data do Abastecimento] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
                        l_nota(x_empresa) = l_nota(x_empresa) + ![Valor Total]
                        l_nota(12) = l_nota(12) + ![Valor Total]
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_baixa_cheque.Close
    tbl_cartao_credito.Close
    tbl_duplicata_receber.Close
    tbl_empresa.Close
    tbl_movimento_cartao_credito.Close
    tbl_movimento_cheque_avista.Close
    tbl_movimento_nota_abastecimento.Close
    
    Set MovCheque = Nothing
    Set rsMovComposicaoCaixa = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 1 To 12
        l_cheque_predatado(i) = 0
        l_amex(i) = 0
        l_credicard(i) = 0
        l_visa(i) = 0
        l_hipercheque(i) = 0
        l_nota(i) = 0
        l_cheque_avista(i) = 0
        l_dinheiro(i) = 0
        l_total(i) = 0
    Next
    For i = 1 To 3
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
    Dim i As Integer
    ZeraVariaveis
    ImpCab
    ImpCab2
    Call LoopEmpresaDebitos
    Call ImpTotal
    ZeraVariaveis
    BuscaPercentualCartoes
    ImpCab3
    Call LoopEmpresaCreditos
    Call ImpTotal
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Análise do Resumo da Movimentação dos Postos|@|"
    frm_preview.Show 1
    cmd_sair.SetFocus
End Sub
Private Sub LoopEmpresaCreditos()
    Dim i As Integer
    With tbl_empresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If !Codigo > 11 Then
                    Exit Do
                End If
                i = !Codigo
                CalculaChequePreDatadoCreditos i
                CalculaChequePreDatadoBaixadoCreditos i
                CalculaCartaoCreditos i
                CalculaDuplicataCreditos i
                CalculaChequeAvista i
                CalculaDinheiro i
                l_total(i) = l_total(i) + l_cheque_predatado(i) + l_amex(i) + l_credicard(i) + l_visa(i) + l_nota(i) + l_cheque_avista(i) + l_dinheiro(i)
                l_total(12) = l_total(12) + l_total(i)
                Call ImpDet(!Nome, l_cheque_predatado(i), l_amex(i), l_credicard(i), l_visa(i), l_nota(i), l_cheque_avista(i), l_dinheiro(i), l_total(i))
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresaDebitos()
    Dim i As Integer
    With tbl_empresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If !Codigo > 11 Then
                    Exit Do
                End If
                i = !Codigo
                CalculaChequePreDatado i
                CalculaCartao i
                CalculaNota i
                CalculaChequeAvista i
                CalculaDinheiro i
                l_nota(i) = l_nota(i) + l_hipercheque(i)
                l_nota(12) = l_nota(12) + l_hipercheque(12)
                l_total(i) = l_total(i) + l_cheque_predatado(i) + l_amex(i) + l_credicard(i) + l_visa(i) + l_nota(i) + l_cheque_avista(i) + l_dinheiro(i) + l_hipercheque(i)
                l_total(12) = l_total(12) + l_total(i)
                Call ImpDet(!Nome, l_cheque_predatado(i), l_amex(i), l_credicard(i), l_visa(i), l_nota(i), l_cheque_avista(i), l_dinheiro(i), l_total(i))
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpTotal()
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "|                               |            |            |            |            |            |            |            |            |"
    Call ImpDet("** Total Geral", l_cheque_predatado(12), l_amex(12), l_credicard(12), l_visa(12), l_nota(12), l_cheque_avista(12), l_dinheiro(12), l_total(12))
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(x_empresa As String, x_cheque_predatado As Currency, x_amex As Currency, x_credicard As Currency, x_visa As Currency, x_nota As Currency, x_cheque_avista As Currency, x_dinheiro As Currency, x_total As Currency)
    Dim x_linha As String
    Dim i As Integer
    If x_total > 0 Then
        x_linha = "|                               |            |            |            |            |            |            |            |            |"
        Mid(x_linha, 3, 29) = x_empresa
        i = Len(Format(x_cheque_predatado, "#,###,##0.00"))
        Mid(x_linha, 34 + 12 - i, i) = Format(x_cheque_predatado, "#,###,##0.00")
        i = Len(Format(x_amex, "#,###,##0.00"))
        Mid(x_linha, 47 + 12 - i, i) = Format(x_amex, "#,###,##0.00")
        i = Len(Format(x_credicard, "#,###,##0.00"))
        Mid(x_linha, 60 + 12 - i, i) = Format(x_credicard, "#,###,##0.00")
        i = Len(Format(x_visa, "#,###,##0.00"))
        Mid(x_linha, 73 + 12 - i, i) = Format(x_visa, "#,###,##0.00")
        i = Len(Format(x_nota, "#,###,##0.00"))
        Mid(x_linha, 86 + 12 - i, i) = Format(x_nota, "#,###,##0.00")
        i = Len(Format(x_cheque_avista, "#,###,##0.00"))
        Mid(x_linha, 99 + 12 - i, i) = Format(x_cheque_avista, "#,###,##0.00")
        i = Len(Format(x_dinheiro, "#,###,##0.00"))
        Mid(x_linha, 112 + 12 - i, i) = Format(x_dinheiro, "#,###,##0.00")
        i = Len(Format(x_total, "#,###,##0.00"))
        Mid(x_linha, 125 + 12 - i, i) = Format(x_total, "#,###,##0.00")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If Mid(x_empresa, 1, 2) = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@Printer.Print " & "|                               |            |            |            |            |            |            |            |            |"
        lLinha = lLinha + 2
    End If
End Sub
Private Sub ImpCab()
    Dim x_linha As String * 137
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
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RESUMO DA MOVIMENTAÇÃO DOS POSTOS                               , __/__/____ |"
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
    x_linha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpCab2()
    Dim x_linha As String
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & "|                                                          D  É  B  I  T  O  S                                                          |"
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                       |  CH. PRED. |    AMEX    |  CREDICARD |    VISA    | NOTA ABAST.| CH.  VISTA |  DINHEIRO  |    TOTAL   |"
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "|                               |            |            |            |            |            |            |            |            |"
End Sub
Private Sub ImpCab3()
    Dim x_linha As String
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & " "
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & "|                                                         C  R  É  D  I  T  O  S                                                        |"
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                       |  CH. PRED. |    AMEX    |  CREDICARD |    VISA    |  DUPLICATA | CH.  VISTA |  DINHEIRO  |    TOTAL   |"
    BioImprime "@Printer.Print " & "+-------------------------------+------------+------------+------------+------------+------------+------------+------------+------------+"
    BioImprime "@Printer.Print " & "|                               |            |            |            |            |            |            |            |            |"
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
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
    cbo_periodo_i.SetFocus
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
    
    Set tbl_baixa_cheque = bd_sgp.OpenTable("Baixa_Cheque")
    Set tbl_cartao_credito = bd_sgp.OpenTable("Cartao_Credito")
    Set tbl_duplicata_receber = bd_sgp.OpenTable("Duplicata_Receber")
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_movimento_cartao_credito = bd_sgp.OpenTable("Movimento_Cartao_Credito")
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    Set tbl_movimento_nota_abastecimento = bd_sgp.OpenTable("Movimento_Nota_Abastecimento")
    tbl_baixa_cheque.Index = "id_vencimento"
    tbl_cartao_credito.Index = "id_codigo"
    tbl_duplicata_receber.Index = "id_empresa_vencimento"
    tbl_empresa.Index = "id_codigo"
    tbl_movimento_cartao_credito.Index = "id_data_emissao"
    tbl_movimento_cheque_avista.Index = "id_digitacao"
    tbl_movimento_nota_abastecimento.Index = "id_data_abastecimento"
    PreencheCboPeriodo
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
        cbo_periodo_i.SetFocus
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
