VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form emissao_cupom_complementar 
   Caption         =   "Emissão do Cupom Complementar"
   ClientHeight    =   2295
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_cupom_complementar.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_cupom_complementar.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_cupom_complementar.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_cupom_complementar.frx":25FA
      Style           =   1  'Graphical
      TabIndex        =   12
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
         Picture         =   "emissao_cupom_complementar.frx":38D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar.frx":4BAE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar.frx":5E88
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
         _Version        =   327680
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
         _Version        =   327680
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
         _Version        =   327680
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
      _Version        =   327680
   End
End
Attribute VB_Name = "emissao_cupom_complementar"
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
Dim lQtdCupomA As Currency
Dim lQtdCupomAA As Currency
Dim lQtdCupomD As Currency
Dim lQtdCupomDA As Currency
Dim lQtdCupomG As Currency
Dim lQtdCupomGA As Currency
Dim lTotalQtdCupom As Currency
Dim lQtdBombaA As Currency
Dim lQtdBombaAA As Currency
Dim lQtdBombaD As Currency
Dim lQtdBombaDA As Currency
Dim lQtdBombaG As Currency
Dim lQtdBombaGA As Currency
Dim lTotalQtdBomba As Currency
Dim lTotalCupomA As Currency
Dim lTotalCupomAA As Currency
Dim lTotalCupomD As Currency
Dim lTotalCupomDA As Currency
Dim lTotalCupomG As Currency
Dim lTotalCupomGA As Currency
Dim lTotalCupom As Currency
Dim lTotalBombaA As Currency
Dim lTotalBombaAA As Currency
Dim lTotalBombaD As Currency
Dim lTotalBombaDA As Currency
Dim lTotalBombaG As Currency
Dim lTotalBombaGA As Currency
Dim lTotalBomba As Currency
Dim tbl_aliquota As Table
Dim tbl_estoque As Table
Dim tbl_movimento_bomba As Table
Dim tbl_movimento_cupom_fiscal As Table
Dim tbl_produto As Table
Private Sub AtualizaTabelaCupomFiscal(x_numero_cupom As Long, x_ordem As Integer, x_data As Date, x_hora As Date, x_codigo_produto As Long, x_valor_unitario As Currency, x_quantidade As Currency, x_valor_total As Currency, x_codigo_aliquota As Integer)
    With tbl_movimento_cupom_fiscal
        !Empresa = g_empresa
        ![Numero do Cupom] = x_numero_cupom
        !Ordem = x_ordem
        !Data = x_data
        !Hora = x_hora
        ![Data do Cupom] = x_data
        !Periodo = 5
        ![Tipo do Movimento] = 1
        ![Codigo do Cliente] = 0
        ![Codigo do Conveniado] = 0
        ![Codigo do Produto] = x_codigo_produto
        ![Valor Unitario] = x_valor_unitario
        !Quantidade = x_quantidade
        ![Valor Total] = x_valor_total
        ![Forma de Pagamento] = 1
        ![Valor Recebido] = x_valor_total
        ![Numero do Cheque] = ""
        !Telefone = ""
        !Operador = 0
        ![Cupom Cancelado] = False
        ![Item Cancelado] = False
        ![Codigo da Aliquota] = x_codigo_aliquota
    End With
End Sub
Private Sub BuscaNumeroCupom()
    Dim x_string As String
    Dim NumeroArquivo As Integer
    On Error GoTo FileError
    If lExisteImpressora Then
        If Not Testa_ImpressoraCF Then
            NumeroArquivo = 99999
        End If
        If l_flag_cupom_fiscal = "F" Then
            'busca numero do cupom da impressora fiscal
            Call Abre_ProtocoloCF(1)
            ComandoCF = Chr(27) + "|30|" + Chr(27)
            Envia_ComandoCF
            Fecha_ProtocoloCF
            NumeroArquivo = FreeFile
            Open "MP20FI.RET" For Input As NumeroArquivo
            Input #NumeroArquivo, x_string
            Close NumeroArquivo
            If Val(x_string) > 0 Then
                txt_numero_cupom = CLng(x_string) + 1
            End If
            'busca item da impressora fiscal
            Call Abre_ProtocoloCF(1)
            ComandoCF = Chr(27) + "|35|12|" + Chr(27)
            Envia_ComandoCF
            Fecha_ProtocoloCF
            NumeroArquivo = FreeFile
            Open "MP20FI.RET" For Input As NumeroArquivo
            Input #NumeroArquivo, x_string
            If Val(x_string) > 0 Then
                txt_ordem = CLng(x_string)
            End If
            Close NumeroArquivo
            txt_ordem = 1
        Else
            txt_numero_cupom = tbl_movimento_cupom_fiscal![Numero do Cupom]
            txt_ordem = tbl_movimento_cupom_fiscal!Ordem + 1
        End If
        'busca data/hora da impressora fiscal
        Call Abre_ProtocoloCF(1)
        ComandoCF = Chr(27) + "|35|23|" + Chr(27)
        Envia_ComandoCF
        Fecha_ProtocoloCF
        NumeroArquivo = FreeFile
        Open "MP20FI.RET" For Input As NumeroArquivo
        Input #NumeroArquivo, x_string
        Close NumeroArquivo
        msk_data = CDate(Mid(x_string, 1, 2) & "/" & Mid(x_string, 3, 2) & "/19" & Mid(x_string, 5, 2))
        l_data_cupom = CDate(msk_data)
        msk_hora = Format(Mid(x_string, 7, 2), "00") & ":" & Format(Mid(x_string, 9, 2), "00") & ":" & Format(Mid(x_string, 11, 2), "00")
    Else
        If l_flag_cupom_fiscal = "F" Then
            txt_numero_cupom = 1
            If tbl_movimento_cupom_fiscal.RecordCount > 0 Then
                tbl_movimento_cupom_fiscal.MoveLast
                txt_numero_cupom = tbl_movimento_cupom_fiscal![Numero do Cupom] + 1
            End If
            txt_ordem = 1
        Else
            txt_numero_cupom = tbl_movimento_cupom_fiscal![Numero do Cupom]
            txt_ordem = tbl_movimento_cupom_fiscal!Ordem + 1
        End If
        msk_data = g_data_def
        l_data_cupom = g_data_def
        msk_hora = Format(Time, "hh:mm:ss")
    End If
    Exit Sub
FileError:
    MsgBox "Não foi possível criar o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub Finaliza()
    tbl_aliquota.Close
    tbl_estoque.Close
    tbl_movimento_bomba.Close
    tbl_movimento_cupom_fiscal.Close
    tbl_produto.Close
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lQtdCupomA = 0
    lQtdCupomAA = 0
    lQtdCupomD = 0
    lQtdCupomDA = 0
    lQtdCupomG = 0
    lQtdCupomGA = 0
    lTotalQtdCupom = 0
    lQtdBombaA = 0
    lQtdBombaAA = 0
    lQtdBombaD = 0
    lQtdBombaDA = 0
    lQtdBombaG = 0
    lQtdBombaGA = 0
    lTotalQtdBomba = 0
    lTotalCupomA = 0
    lTotalCupomAA = 0
    lTotalCupomD = 0
    lTotalCupomDA = 0
    lTotalCupomG = 0
    lTotalCupomGA = 0
    lTotalCupom = 0
    lTotalBombaA = 0
    lTotalBombaAA = 0
    lTotalBombaD = 0
    lTotalBombaDA = 0
    lTotalBombaG = 0
    lTotalBombaGA = 0
    lTotalBomba = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento_bomba
    With tbl_movimento_bomba
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, CDate(msk_data_i), 0, 0
            If Not .NoMatch Then
                ImpDados
            End If
        End If
    End With
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    LoopMovimentoBomba
    LoopTabelaProduto
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        If lLocal = 1 Then
            If (MsgBox("Após a emissão do cupom complementar será impresso a REDUÇÃO Z." & Chr(13) & "E não será mais aceito a emissão de cupom fiscal nesta data." & Chr(13) & Chr(13) & "Deseja realmente imprimir o cupom complementar?", vbQuestion + vbYesNo + vbDefaultButton2, "Emissão do Cupom Complementar")) = 6 Then
                MsgBox "Aguarde Imprimindo o cupom complementar"
                ImpCupomComplementar
            End If
        Else
            g_string = lLocal & lNomeArquivo
            frm_preview.Show 1
        End If
    End If
End Sub
Private Sub LoopMovimentoBomba()
    Dim i As Integer
    'loop movimento das bombas
    With tbl_movimento_bomba
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, CDate(msk_data_i), 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    lTotalQtdBomba = lTotalQtdBomba + ![Quantidade da Saida]
                    lTotalBomba = lTotalBomba + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    Select Case Trim(![Tipo de Combustivel])
                        Case "A"
                            lQtdBombaA = lQtdBombaA + ![Quantidade da Saida]
                            lTotalBombaA = lTotalBombaA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                        Case "AA"
                            lQtdBombaAA = lQtdBombaAA + ![Quantidade da Saida]
                            lTotalBombaAA = lTotalBombaAA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                        Case "D"
                            lQtdBombaD = lQtdBombaD + ![Quantidade da Saida]
                            lTotalBombaD = lTotalBombaD + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                        Case "DA"
                            lQtdBombaDA = lQtdBombaDA + ![Quantidade da Saida]
                            lTotalBombaDA = lTotalBombaDA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                        Case "G"
                            lQtdBombaG = lQtdBombaG + ![Quantidade da Saida]
                            lTotalBombaG = lTotalBombaG + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                        Case "GA"
                            lQtdBombaGA = lQtdBombaGA + ![Quantidade da Saida]
                            lTotalBombaGA = lTotalBombaGA + Format(![Quantidade da Saida] * ![Preco de Venda], "#########0.00")
                    End Select
                    .MoveNext
                Loop
            End If
        End If
    End With
End Sub
Private Sub LoopMovimentoCupomFiscal()
    Dim x_linha As String
    Dim x_qtd_cupom As Currency
    Dim x_qtd_bomba As Currency
    Dim x_total_cupom As Currency
    Dim x_total_bomba As Currency
    Dim x_tipo_combustivel As String
    'loop Movimento do Cupom Fiscal
    x_qtd_cupom = 0
    x_qtd_bomba = 0
    x_total_cupom = 0
    x_total_bomba = 0
    x_tipo_combustivel = ""
    With tbl_movimento_cupom_fiscal
        .Seek ">=", g_empresa, CDate(msk_data_i), tbl_produto!Codigo2, 0, 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Or ![Codigo do Produto] <> tbl_produto!Codigo2 Then
                    Exit Do
                End If
                x_qtd_cupom = x_qtd_cupom + !Quantidade
                x_total_cupom = x_total_cupom + ![Valor Total]
                lTotalQtdCupom = lTotalQtdCupom + !Quantidade
                lTotalCupom = lTotalCupom + ![Valor Total]
                If tbl_produto!Nome Like "*ALCOOL*" Then
                    If tbl_produto!Nome Like "*ADITIVADO*" Then
                        x_tipo_combustivel = "AA"
                        lQtdCupomAA = lQtdCupomAA + !Quantidade
                        lTotalCupomAA = lTotalCupomAA + ![Valor Total]
                    Else
                        x_tipo_combustivel = "A"
                        lQtdCupomA = lQtdCupomA + !Quantidade
                        lTotalCupomA = lTotalCupomA + ![Valor Total]
                    End If
                ElseIf tbl_produto!Nome Like "*DIESEL*" Then
                    If tbl_produto!Nome Like "*ADITIVADO*" Then
                        x_tipo_combustivel = "DA"
                        lQtdCupomDA = lQtdCupomDA + !Quantidade
                        lTotalCupomDA = lTotalCupomDA + ![Valor Total]
                    Else
                        x_tipo_combustivel = "D"
                        lQtdCupomD = lQtdCupomD + !Quantidade
                        lTotalCupomD = lTotalCupomD + ![Valor Total]
                    End If
                ElseIf tbl_produto!Nome Like "*GASOLINA*" Then
                    If tbl_produto!Nome Like "*ADITIVADO*" Then
                        x_tipo_combustivel = "GA"
                        lQtdCupomGA = lQtdCupomGA + !Quantidade
                        lTotalCupomGA = lTotalCupomGA + ![Valor Total]
                    Else
                        x_tipo_combustivel = "G"
                        lQtdCupomG = lQtdCupomG + !Quantidade
                        lTotalCupomG = lTotalCupomG + ![Valor Total]
                    End If
                Else
                    MsgBox "Itém não identificado " & tbl_produto!Nome, vbInformation, "Erro de Integridade!"
                End If
                .MoveNext
            Loop
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 60 Then
                x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
                Mid(x_linha, 12, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            If x_tipo_combustivel = "A" Then
                x_qtd_bomba = lQtdBombaA
                x_total_bomba = lTotalBombaA
            ElseIf x_tipo_combustivel = "AA" Then
                x_qtd_bomba = lQtdBombaAA
                x_total_bomba = lTotalBombaAA
            ElseIf x_tipo_combustivel = "D" Then
                x_qtd_bomba = lQtdBombaD
                x_total_bomba = lTotalBombaD
            ElseIf x_tipo_combustivel = "DA" Then
                x_qtd_bomba = lQtdBombaDA
                x_total_bomba = lTotalBombaDA
            ElseIf x_tipo_combustivel = "G" Then
                x_qtd_bomba = lQtdBombaG
                x_total_bomba = lTotalBombaG
            ElseIf x_tipo_combustivel = "GA" Then
                x_qtd_bomba = lQtdBombaGA
                x_total_bomba = lTotalBombaGA
            End If
            Call ImpDet(tbl_produto!Codigo2, tbl_produto!Nome, tbl_produto!unidade, x_qtd_cupom, x_total_cupom, x_qtd_bomba, x_total_bomba)
        End If
    End With
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela produto
    With tbl_produto
        .Seek ">=", 4, 0
        If Not .NoMatch Then
            Do Until .EOF
                If ![Codigo do Grupo] <> 4 Then
                    Exit Do
                End If
                LoopMovimentoCupomFiscal
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpDet(x_codigo As Long, x_nome As String, x_unidade As String, x_qtd_cupom As Currency, x_valor_cupom As Currency, x_qtd_bomba As Currency, x_valor_bomba As Currency)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|      |                                           |   |          |               |          |               |          |               |"
    i = Len(Format(x_codigo, "#,000"))
    Mid(x_linha, 2 + 5 - i, i) = Format(x_codigo, "#,000")
    Mid(x_linha, 10, 40) = x_nome
    Mid(x_linha, 53, 3) = x_unidade
    i = Len(Format(x_qtd_cupom, "###,##0.00"))
    Mid(x_linha, 57 + 10 - i, i) = Format(x_qtd_cupom, "###,##0.00")
    i = Len(Format(x_valor_cupom, "###,###,##0.00"))
    Mid(x_linha, 69 + 14 - i, i) = Format(x_valor_cupom, "###,###,##0.00")
    i = Len(Format(x_qtd_bomba, "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(x_qtd_bomba, "###,###,##0.00")
    i = Len(Format(x_valor_bomba, "###,###,##0.00"))
    Mid(x_linha, 96 + 14 - i, i) = Format(x_valor_bomba, "###,###,##0.00")
    i = Len(Format(x_qtd_bomba - x_qtd_cupom, "###,###,##0.00"))
    Mid(x_linha, 107 + 14 - i, i) = Format(x_qtd_bomba - x_qtd_cupom, "###,###,##0.00")
    i = Len(Format(x_valor_bomba - x_valor_cupom, "###,###,##0.00"))
    Mid(x_linha, 123 + 14 - i, i) = Format(x_valor_bomba - x_valor_cupom, "###,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
    If lLocal = 1 Then
        If (x_qtd_bomba - x_qtd_cupom) > 0 Then
            x_linha = Format(tbl_produto!Codigo2, "00")
            x_linha = x_linha & tbl_produto!Nome
            x_linha = x_linha & tbl_produto!unidade
            x_linha = x_linha & Format(tbl_produto![Codigo da Aliquota], "00")
            x_linha = x_linha & Format(x_qtd_bomba - x_qtd_cupom, "0000000000.00")
            x_linha = x_linha & Format(tbl_produto![Preco de Venda], "0000000000.0000")
            x_linha = x_linha & Format(x_valor_bomba - x_valor_cupom, "0000000000.00")
            Print #3, x_linha
        End If
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    
    If lLocal = 1 Then
        Print #3, "FIM"
    End If
    
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                               *** TOTAL DO RELATORIO |          |               |          |               |          |               |"
    i = Len(Format(lTotalQtdCupom, "###,##0.00"))
    Mid(x_linha, 57 + 10 - i, i) = Format(lTotalQtdCupom, "###,##0.00")
    i = Len(Format(lTotalCupom, "###,###,##0.00"))
    Mid(x_linha, 69 + 14 - i, i) = Format(lTotalCupom, "###,###,##0.00")
    i = Len(Format(lTotalQtdBomba, "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(lTotalQtdBomba, "###,###,##0.00")
    i = Len(Format(lTotalBomba, "###,###,##0.00"))
    Mid(x_linha, 96 + 14 - i, i) = Format(lTotalBomba, "###,###,##0.00")
    i = Len(Format(lTotalQtdBomba - lTotalQtdCupom, "###,###,##0.00"))
    Mid(x_linha, 107 + 14 - i, i) = Format(lTotalQtdBomba - lTotalQtdCupom, "###,###,##0.00")
    i = Len(Format(lTotalBomba - lTotalCupom, "###,###,##0.00"))
    Mid(x_linha, 123 + 14 - i, i) = Format(lTotalBomba - lTotalCupom, "###,###,##0.00")
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print x_linha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+------------------------------------------------------+----------+---------------+----------+---------------+----------+---------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                                                                           Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| CUPOM COMPLEMENTAR DO PERIODO: __/__/____ A __/__/____.                                                           Goiânia, __/__/____ |"
    Mid(x_linha, 34, 10) = msk_data_i
    Mid(x_linha, 47, 10) = msk_data_f
    Mid(x_linha, 126, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    x_linha = "+------+-------------------------------------------+---+--------------------------+--------------------------+--------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|CODIGO|                                           |   | C U P O M    F I S C A L |        V E N D A S       |   CUPOM   COMPLEMENTAR   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|  DO  | DISCRIMINAÇÃO DOS PRODUTOS                |UN.+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PROD.|                                           |   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|      |                                           |   |          |               |          |               |          |               |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpCupomComplementar()
    Dim x_linha As String
    Dim x_string As String
    Dim i As Integer
    Dim x_valor_acrescimo As Currency
    Dim x_valor_desconto As Currency
    Dim x_total As Currency
    Close #3
    x_total = 0
    'Abre o cupom fiscal
    Call Abre_ProtocoloCF(1)
    ComandoCF = Chr(27) + "|00|" + Chr(27)
    Envia_ComandoCF
    Fecha_ProtocoloCF
    Open "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT" For Input As #3
    Do Until EOF(3)
        Line Input #3, x_linha
        If Mid(x_linha, 1, 3) = "FIM" Then
            Exit Do
        End If
        tbl_aliquota.Seek "=", Mid(x_linha, 47, 2)
        'Venda de Item com entrada de departamento,
        'Verifica se há diferença do total
        x_string = Format(Format(fValidaValor(Mid(x_linha, 62, 15)) * fValidaValor(Mid(x_linha, 49, 13)), "###,##0.0000"), "###,##0.0000")
        i = Len(x_string)
        x_string = Mid(x_string, 1, i - 2)
        x_valor_acrescimo = 0
        x_valor_desconto = 0
        If fValidaValor(Mid(x_linha, 77, 13)) > fValidaValor(x_string) Then
            x_valor_acrescimo = fValidaValor(Mid(x_linha, 77, 13)) - fValidaValor(x_string)
        ElseIf fValidaValor(Mid(x_linha, 77, 13)) < fValidaValor(x_string) Then
            x_valor_desconto = fValidaValor(x_string) - fValidaValor(Mid(x_linha, 77, 13))
        Else
        End If
        'desconto e unidade de medida
        Call Abre_ProtocoloCF(1)
        ComandoCF = Chr(27) + "|63|"
        'tipo de tributação
        If Not tbl_aliquota.NoMatch Then
            ComandoCF = ComandoCF + tbl_aliquota![Codigo Fiscal] + "|"
        Else
            ComandoCF = ComandoCF + "II" + "|"
        End If
        'Valor Unitário
        x_string = Format(Mid(x_linha, 62, 15), "000000.000")
        ComandoCF = ComandoCF + Mid(x_string, 1, 6) + Mid(x_string, 8, 3) + "|"
        'Quantidade
        x_string = Format(Mid(x_linha, 49, 13), "0000.000")
        ComandoCF = ComandoCF + Mid(x_string, 1, 4) + Mid(x_string, 6, 3) + "|"
        'Valor do Desconto
        x_string = Format(x_valor_desconto, "00000000.00")
        ComandoCF = ComandoCF + Mid(x_string, 1, 8) + Mid(x_string, 10, 2) + "|"
        'Valor do Acréscimo
        x_string = Format(x_valor_acrescimo, "00000000.00")
        ComandoCF = ComandoCF + Mid(x_string, 1, 8) + Mid(x_string, 10, 2) + "|"
        'Departamento
        ComandoCF = ComandoCF + Format(1, "00") + "|"
        'Não Usado
        ComandoCF = ComandoCF + "00000000000000000000" + "|"
        'Unidade de Medida
         x_string = Mid(x_linha, 44, 3)
        ComandoCF = ComandoCF + Mid(x_string, 1, 2) + "|"
        'código do produto
        ComandoCF = ComandoCF + Format(Mid(x_linha, 1, 3), "#,##0") + "|"
        'nome do produto
        x_string = Mid(x_linha, 4, 40)
        ComandoCF = ComandoCF + Mid(x_string, 1, 40) + "|"
        ComandoCF = ComandoCF + Chr(27)
        Envia_ComandoCF
        Fecha_ProtocoloCF
        Call AtualizaTabelaCupomFiscal
        '(x_numero_cupom As Long, x_ordem As Integer, x_data As Date, x_hora As Date, x_codigo_produto As Long, x_valor_unitario As Currency, x_quantidade As Currency, x_valor_total As Currency, x_codigo_aliquota As Integer)
    Loop
    'Desconto para o Cupom Fiscal
    Call Abre_ProtocoloCF(1)
    ComandoCF = Chr(27) + "|32|A|0000|" + Chr(27)
    Envia_ComandoCF
    Fecha_ProtocoloCF
    'Efetua Forma de Pagamento
    Call Abre_ProtocoloCF(1)
    ComandoCF = Chr(27) + "|72|01|" + Format(x_total, "000000000000.00") + "|" + Chr(27)
    Envia_ComandoCF
    Fecha_ProtocoloCF
    'Fecha Cupom Fiscal
    Call Abre_ProtocoloCF(1)
    ComandoCF = Chr(27) + "|34|Cerrado_Informatica - (062) 941-3044            Sistemas para Automacao Comercial               |" + Chr(27)
    Envia_ComandoCF
    Fecha_ProtocoloCF
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
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
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    Open "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT" For Output As #3
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Relatorio
        End If
    End If
    Close #3
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
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data) Then
        msk_data = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f = Format(g_data_def, "dd/mm/yyyy")
        cmd_visualizar.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_aliquota = bd_sgp.OpenTable("Aliquota")
    Set tbl_estoque = bd_sgp.OpenTable("Estoque")
    Set tbl_movimento_bomba = bd_sgp.OpenTable("Movimento_Bomba")
    Set tbl_movimento_cupom_fiscal = bd_sgp.OpenTable("Movimento_Cupom_Fiscal")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    tbl_aliquota.Index = "id_codigo"
    tbl_estoque.Index = "id_codigo2"
    tbl_movimento_bomba.Index = "id_data"
    tbl_movimento_cupom_fiscal.Index = "id_data_produto"
    tbl_produto.Index = "id_codigo"
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
        cmd_visualizar.SetFocus
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
