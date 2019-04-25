VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_resumo_folha_pagamento 
   Caption         =   "Emite Resumo da Folha de Pagamento"
   ClientHeight    =   2235
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   5355
   Icon            =   "lst_resumo_folha_pagamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_resumo_folha_pagamento.frx":030A
   ScaleHeight     =   2235
   ScaleWidth      =   5355
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   780
      Picture         =   "lst_resumo_folha_pagamento.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Visualiza o resumo da folha de pagamento."
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2280
      Picture         =   "lst_resumo_folha_pagamento.frx":1E6A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprime o resumo da folha de pagamento."
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3780
      Picture         =   "lst_resumo_folha_pagamento.frx":3474
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1260
      Width           =   795
   End
   Begin VB.Frame frmDados
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5115
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2880
         Picture         =   "lst_resumo_folha_pagamento.frx":4B06
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_registro 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1635
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo de registro"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "&Mês/Ano"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1875
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_resumo_folha_pagamento"
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
Dim l_ano_mes As String * 6
Dim l_provento_desconto As String * 1
Dim l_desconto As Currency
Dim l_provento As Currency
Dim l_total As Currency
Dim l_fgts As Currency
Dim l_base_calculo_inss As Currency
Dim i_provento As Integer
Dim i_desconto As Integer
Dim l_codigo_provento(1 To 20) As Integer
Dim l_codigo_desconto(1 To 20) As Integer
Dim l_nome_provento(1 To 20) As String
Dim l_nome_desconto(1 To 20) As String
Dim l_quantidade_provento(1 To 20) As Currency
Dim l_quantidade_desconto(1 To 20) As Currency
Dim l_valor_provento(1 To 20) As Currency
Dim l_valor_desconto(1 To 20) As Currency
Dim tbl_funcionario As Table
Dim tbl_movimento_folha As Table
Dim tbl_tabela_provento_desconto As Table
Function ExisteMovimento() As Boolean
    Dim x_mes_ano As String
    Dim x_data As Date
    x_data = "01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
    If Val(Mid(l_ano_mes, 5, 2)) < 13 Then
        x_mes_ano = Format(x_data, "mmmm") & " / " & Format(x_data, "yyyy")
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        x_mes_ano = "13o Salario - 1a Parcela"
    End If
    ExisteMovimento = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Index = "id_data"
            .Seek ">=", g_empresa, l_ano_mes, 0, 0
            If Not .NoMatch Then
                If !Empresa = g_empresa And ![Mes Ano] = l_ano_mes Then
                    ExisteMovimento = True
                    Exit Function
                End If
            End If
        End If
    End With
    MsgBox "Não existe movimento no período " & x_mes_ano & ".", vbInformation, "Sem Movimento!"
End Function
Function CalculaValor(x_codigo As Integer) As Currency
    Dim x_tipo As Integer
    CalculaValor = 0
    x_tipo = cbo_tipo_registro.ItemData(cbo_tipo_registro.ListIndex)
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, x_codigo, l_ano_mes, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Movimento] <> x_codigo Or ![Mes Ano] <> l_ano_mes Then
                        Exit Do
                    End If
                    tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
                    If Not tbl_funcionario.NoMatch Then
                        If (x_tipo = 1 And tbl_funcionario![Serie da Carteira de Trabalho] <> "NR") Or (x_tipo = 2 And tbl_funcionario![Serie da Carteira de Trabalho] = "NR") Or x_tipo = 3 Then
                            CalculaValor = CalculaValor + !valor
                            If ![Codigo do Movimento] < 500 And ![Codigo do Movimento] <> 15 Then
                                l_base_calculo_inss = l_base_calculo_inss + !valor
                            ElseIf ![Codigo do Movimento] = 510 Or ![Codigo do Movimento] = 515 Then
                                l_base_calculo_inss = l_base_calculo_inss - !valor
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_funcionario.Close
    tbl_movimento_folha.Close
    tbl_tabela_provento_desconto.Close
End Sub
Private Sub cbo_tipo_registro_GotFocus()
    SendMessageLong cbo_tipo_registro.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_registro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = Mid(RetiraGString(1), 4, 7)
    cbo_tipo_registro.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If ExisteMovimento Then
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                Relatorio
            End If
        End If
    End If
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    Dim i As Integer
    Dim x_mes_ano As String
    If Val(Mid(l_ano_mes, 5, 2)) < 13 Then
        x_mes_ano = UCase(Format(CDate("01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)), "mmmm")) & " de " & Format(Format(CDate("01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)), "yyyy"), "#,###")
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        x_mes_ano = "13o Salario - 1a Parcela"
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        x_mes_ano = "13o Salario - 2a Parcela"
    End If
    x_mes_ano = x_mes_ano & " - " & Trim(cbo_tipo_registro)
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
    x_linha = "| RESUMO DA FOLHA DE PAGAMENTO                                    , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = g_data_def
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE A.:                                                                |"
    Mid(x_linha, 17, 40) = x_mes_ano
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+-------------+-------------+"
    BioImprime "@Printer.Print " & "| CODIGO| PROVENTO/DESCONTO                        |   PROVENTO  |   DESCONTO  |"
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+-------------+-------------+"
End Sub
Private Sub ImpDet(x_codigo As Integer, x_nome As String, x_valor As Currency)
    Dim x_linha As String
    Dim i As Integer
    If lPagina = 0 Then
        ImpCab
    End If
    If l_provento_desconto = " " Then
        l_provento_desconto = "P"
    End If
    If tbl_tabela_provento_desconto![Provento ou Desconto] = "D" And l_provento_desconto = "P" Then
        Call ImpSubTotal("P")
        l_provento_desconto = "D"
    End If
    If lLinha >= 60 Then
        x_linha = "+-------+------------------------------------------+-------------+-------------+"
        Mid(x_linha, 13, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 4, 3) = Format(x_codigo, "000")
    Mid(x_linha, 11, 40) = x_nome
    i = Len(Format(x_valor, "####,##0.00"))
    If tbl_tabela_provento_desconto![Provento ou Desconto] = "P" Then
        Mid(x_linha, 54 + 11 - i, i) = Format(x_valor, "####,##0.00")
    Else
        Mid(x_linha, 68 + 11 - i, i) = Format(x_valor, "####,##0.00")
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal(x_tipo As String)
    Dim x_linha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+-------------+-------------+"
    x_linha = "|       |                                          |             |             |"
    If x_tipo = "P" Then
        Mid(x_linha, 11, 40) = "Total dos Proventos"
        i = Len(Format(l_provento, "####,##0.00"))
        Mid(x_linha, 54 + 11 - i, i) = Format(l_provento, "####,##0.00")
    ElseIf x_tipo = "D" Then
        Mid(x_linha, 11, 40) = "Total dos Descontos"
        i = Len(Format(l_desconto, "####,##0.00"))
        Mid(x_linha, 68 + 11 - i, i) = Format(l_desconto, "####,##0.00")
    End If
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-------+------------------------------------------+-------------+-------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    'Total Líquido
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 11, 40) = "Total Líquido"
    i = Len(Format(l_total, "####,##0.00"))
    Mid(x_linha, 54 + 11 - i, i) = Format(l_total, "####,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-------+------------------------------------------+-------------+-------------+"
    BioImprime "@Printer.Print " & x_linha
    'FGTS
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 11, 40) = "Total de FGTS do Mês"
    i = Len(Format(l_fgts, "####,##0.00"))
    Mid(x_linha, 40 + 11 - i, i) = Format(l_fgts, "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-------+------------------------------------------+-------------+-------------+"
    BioImprime "@Printer.Print " & x_linha
    'Base de Cálculo do INSS
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 11, 40) = "Base Calculo INSS   "
    i = Len(Format(l_base_calculo_inss, "####,##0.00"))
    Mid(x_linha, 40 + 11 - i, i) = Format(l_base_calculo_inss, "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-------+------------------------------------------+-------------+-------------+"
    Mid(x_linha, 13, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub LoopFuncionario()
    Dim x_tipo As Integer
    x_tipo = cbo_tipo_registro.ItemData(cbo_tipo_registro.ListIndex)
    With tbl_funcionario
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, "", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Then
                        Exit Do
                    End If
                    If (x_tipo = 1 And ![Serie da Carteira de Trabalho] <> "NR") Or (x_tipo = 2 And ![Serie da Carteira de Trabalho] = "NR") Or x_tipo = 3 Then
                        Call LoopMovimentoFuncionario(!Codigo)
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Sub
Private Sub LoopImprimeMovimentoFuncionario()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    If lPagina = 0 Then
        ImpCab
    End If
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 4, 3) = Format(tbl_funcionario!Codigo, "000")
    Mid(x_linha, 11, 40) = tbl_funcionario!Nome
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    lLinha = lLinha + 1
    For i2 = 1 To 20
        If i2 > i_provento Then
            LoopImprimeTotalMovimentoFuncionario
            Exit For
        End If
        If lLinha >= 60 Then
            x_linha = "+-------+------------------------------------------+-------------+-------------+"
            Mid(x_linha, 13, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        x_linha = "|       |                                          |             |             |"
        Mid(x_linha, 4, 3) = Format(l_codigo_provento(i2), "000")
        Mid(x_linha, 11, 40) = l_nome_provento(i2)
        i = Len(Format(l_valor_provento(i2), "####,##0.00"))
        If l_codigo_provento(i2) < 500 Then
            Mid(x_linha, 54 + 11 - i, i) = Format(l_valor_provento(i2), "####,##0.00")
        Else
            Mid(x_linha, 68 + 11 - i, i) = Format(l_valor_provento(i2), "####,##0.00")
        End If
        BioImprime "@Printer.Print " & x_linha
        lLinha = lLinha + 1
    Next
End Sub
Private Sub LoopImprimeTotalMovimentoFuncionario()
    Dim x_linha As String
    Dim i As Integer
    Dim x_provento, x_desconto, x_fgts As Currency
    x_provento = 0
    x_desconto = 0
    x_fgts = 0
    For i = 1 To 20
        If i > i_provento Then
            Exit For
        End If
        If l_codigo_provento(i) < 500 Then
            x_provento = x_provento + l_valor_provento(i)
        Else
            x_desconto = x_desconto + l_valor_provento(i)
        End If
        If l_codigo_provento(i) < 500 And l_codigo_provento(i) <> 15 Then
            x_fgts = x_fgts + l_valor_provento(i)
        End If
    Next
    If lLinha >= 60 Then
        x_linha = "+-------+------------------------------------------+-------------+-------------+"
        Mid(x_linha, 13, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    'Total dos Proventos/Descontos
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 11, 40) = "Total dos Proventos/Descontos"
    i = Len(Format(x_provento, "####,##0.00"))
    Mid(x_linha, 54 + 11 - i, i) = Format(x_provento, "####,##0.00")
    i = Len(Format(x_desconto, "####,##0.00"))
    Mid(x_linha, 68 + 11 - i, i) = Format(x_desconto, "####,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|       |                                          |             |             |"
    'Total Líquido
    Mid(x_linha, 11, 40) = "Total Líquido"
    x_provento = x_provento - x_desconto
    If x_provento > 0 Then
        i = Len(Format(x_provento, "####,##0.00"))
        Mid(x_linha, 54 + 11 - i, i) = Format(x_provento, "####,##0.00")
    Else
        i = Len(Format(x_provento, "####,##0.00"))
        Mid(x_linha, 68 + 11 - i, i) = Format(x_provento, "####,##0.00")
    End If
    BioImprime "@Printer.Print " & x_linha
    'FGTS
    x_fgts = Format(x_fgts * 8 / 100, "00000000.00")
    If Mid(l_ano_mes, 5, 2) = 14 Then
        x_fgts = Format(x_fgts / 2, "00000000.00")
    End If
    l_fgts = l_fgts + x_fgts
    x_linha = "|       |                                          |             |             |"
    Mid(x_linha, 11, 40) = "FGTS do Mês"
    i = Len(Format(x_fgts, "####,##0.00"))
    Mid(x_linha, 40 + 11 - i, i) = Format(x_fgts, "####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+-------------+-------------+"
    lLinha = lLinha + 4
End Sub
Private Sub LoopMovimentoFuncionario(x_codigo As Integer)
    Dim i As Integer
    i_provento = 0
    i_desconto = 0
    For i = 1 To 20
        l_codigo_provento(i) = 0
        l_codigo_desconto(i) = 0
        l_nome_provento(i) = ""
        l_nome_desconto(i) = ""
        l_quantidade_provento(i) = 0
        l_quantidade_desconto(i) = 0
        l_valor_provento(i) = 0
        l_valor_desconto(i) = 0
    Next
    With tbl_movimento_folha
        .Index = "id_data"
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, l_ano_mes, x_codigo, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> x_codigo Or ![Mes Ano] <> l_ano_mes Then
                        Exit Do
                    End If
                    tbl_tabela_provento_desconto.Seek "=", ![Codigo do Movimento]
                    If Not tbl_tabela_provento_desconto.NoMatch Then
'                        If ![Codigo do Movimento] < 500 Then
                            i_provento = i_provento + 1
                            l_codigo_provento(i_provento) = ![Codigo do Movimento]
                            l_nome_provento(i_provento) = tbl_tabela_provento_desconto!Nome
                            l_quantidade_provento(i_provento) = !Quantidade
                            l_valor_provento(i_provento) = !valor
'                        Else
'                            i_desconto = i_desconto + 1
'                            l_codigo_desconto(i_desconto) = ![Codigo do Movimento]
'                            l_nome_desconto(i_desconto) = tbl_tabela_provento_desconto!Nome
'                            l_quantidade_desconto(i_desconto) = !Quantidade
'                            l_valor_desconto(i_desconto) = !Valor
'                        End If
                    Else
                        MsgBox "Provento/Desconto inexistente: " & !Codigo, vbInformation, "Erro de Integridade!"
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    If i_provento > 0 Or i_desconto > 0 Then
        LoopImprimeMovimentoFuncionario
    End If
End Sub
Private Sub PreencheCboTipoRegistro()
    cbo_tipo_registro.Clear
    cbo_tipo_registro.AddItem "Registrados"
    cbo_tipo_registro.ItemData(cbo_tipo_registro.NewIndex) = 1
    cbo_tipo_registro.AddItem "Não Registrados"
    cbo_tipo_registro.ItemData(cbo_tipo_registro.NewIndex) = 2
    cbo_tipo_registro.AddItem "Geral"
    cbo_tipo_registro.ItemData(cbo_tipo_registro.NewIndex) = 3
End Sub
Private Sub Relatorio()
    Dim x_valor As Currency
    ZeraVariaveis
    With tbl_tabela_provento_desconto
        If .RecordCount > 0 Then
            LoopFuncionario
            lLinha = 65
            tbl_movimento_folha.Index = "id_movimento"
            tbl_funcionario.Index = "id_codigo"
            .MoveFirst
            Do Until .EOF
                x_valor = CalculaValor(!Codigo)
                If ![Provento ou Desconto] = "D" Then
                    l_desconto = l_desconto + x_valor
                    l_total = l_total - x_valor
                Else
                    l_provento = l_provento + x_valor
                    l_total = l_total + x_valor
                End If
                Call ImpDet(!Codigo, !Nome, x_valor)
                .MoveNext
            Loop
            tbl_movimento_folha.Index = "id_data"
            tbl_funcionario.Index = "id_nome"
        End If
    End With
    If l_provento > 0 Then
        Call ImpSubTotal("D")
        Call ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Resumo da Folha de Pagamento|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate("01/" & msk_data) Then
        MsgBox "Informe o mês/ano.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_tipo_registro.ListIndex = -1 Then
        MsgBox "Selecione o tipo de registro.", 64, "Atenção!"
        cbo_tipo_registro.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    l_desconto = 0
    l_provento = 0
    l_total = 0
    l_fgts = 0
    l_base_calculo_inss = 0
    l_provento_desconto = ""
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If ExisteMovimento Then
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 6, "")
                Relatorio
            End If
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    Screen.MousePointer = 1
    If msk_data.Text = "__/____" Then
        msk_data.Text = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
        msk_data.SetFocus
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
    CentraForm Me
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_folha = bd_sgp.OpenTable("Movimento_Folha")
    Set tbl_tabela_provento_desconto = bd_sgp.OpenTable("Tabela_Provento_Desconto")
    tbl_funcionario.Index = "id_nome"
    tbl_movimento_folha.Index = "id_data"
    tbl_tabela_provento_desconto.Index = "id_codigo"
    PreencheCboTipoRegistro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 2
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_registro.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If IsDate("01/" & msk_data) Then
        l_ano_mes = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    End If
End Sub
