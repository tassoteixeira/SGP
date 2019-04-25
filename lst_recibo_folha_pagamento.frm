VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_recibo_folha_pagamento 
   Caption         =   "Emite Recibo da Folha de Pagamento"
   ClientHeight    =   2175
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7875
   Icon            =   "lst_recibo_folha_pagamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_recibo_folha_pagamento.frx":030A
   ScaleHeight     =   2175
   ScaleWidth      =   7875
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2160
      Picture         =   "lst_recibo_folha_pagamento.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime recibo da folha de pagamento de funcionário."
      Top             =   1200
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "lst_recibo_folha_pagamento.frx":1D5A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1200
      Width           =   795
   End
   Begin VB.Frame frmDados
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2880
         Picture         =   "lst_recibo_folha_pagamento.frx":33EC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4860
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Funcionario"
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   300
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   555
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "lst_recibo_folha_pagamento.frx":46C6
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Mês/Ano"
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_recibo_folha_pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_ano_mes As String * 6
Dim l_codigo As Integer
Dim l_funcionario As Integer
Dim l_local As Integer
Dim tbl_empresa As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_falta_funcionario As Table
Dim tbl_movimento_folha As Table
Dim tbl_tabela_folha As Table
Dim tbl_tabela_provento_desconto As Table
Function ExisteMovimento(x_data As String, x_funcionario As Integer) As Boolean
    ExisteMovimento = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, x_data, x_funcionario, 0
            If Not .NoMatch Then
                If !Empresa = g_empresa And ![Mes Ano] = x_data And ![Codigo do Funcionario] = x_funcionario Then
                    ExisteMovimento = True
                End If
            End If
        End If
    End With
End Function
Function BuscaRegistro2(x_data As String, x_funcionario As Integer, x_codigo As Integer) As Boolean
    BuscaRegistro2 = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek "=", g_empresa, x_data, x_funcionario, x_codigo
            If Not .NoMatch Then
                BuscaRegistro2 = True
                Exit Function
            End If
        End If
    End With
End Function
Function CalculaBaseCalculo(x_codigo As Integer) As Currency
    CalculaBaseCalculo = 0
    If Len(Trim(tbl_tabela_provento_desconto![Base para Calculo])) > 0 Then
        CalculaBaseCalculo = CalculaBaseCalculoComposta
    End If
    'Salário Base
    If x_codigo = 1 Then
        CalculaBaseCalculo = tbl_funcionario![Salario Base]
    'Salário Família
    ElseIf x_codigo = 15 Then
        If CalculaBaseCalculo > tbl_tabela_folha![Salario Familia Acima De] Then
            CalculaBaseCalculo = tbl_tabela_folha![Salario Familia 2]
        Else
            CalculaBaseCalculo = tbl_tabela_folha![Salario Familia 1]
        End If
    'Cesta Básica
    ElseIf x_codigo = 560 Then
        CalculaBaseCalculo = tbl_tabela_folha![Cesta Basica]
    End If
    CalculaBaseCalculo = Format(CalculaBaseCalculo, "########0.00")
End Function
Function CalculaBaseCalculoComposta() As Currency
    Dim x_codigo As Integer
    Dim i As Integer
    Dim i2 As Integer
    CalculaBaseCalculoComposta = 0
    i2 = Len(tbl_tabela_provento_desconto![Base para Calculo]) / 5
    i = 1
    Do Until i > i2
        x_codigo = RetiraString(tbl_tabela_provento_desconto![Base para Calculo], i)
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), tbl_funcionario!Codigo, x_codigo) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta + tbl_movimento_folha!valor
        End If
        i = i + 1
    Loop
    'Se o movimento for INSS
    If tbl_tabela_provento_desconto!Codigo = 520 Then
        'Verifica se tem "Falta" e deduz na base de cálculo
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), l_funcionario, 510) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta - tbl_movimento_folha!valor
        End If
        'Verifica se tem "DSR" e deduz na base de cálculo
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), l_funcionario, 515) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta - tbl_movimento_folha!valor
        End If
    End If
    CalculaBaseCalculoComposta = Format(CalculaBaseCalculoComposta, "########0.00")
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_empresa.Close
    tbl_funcionario.Close
    tbl_movimento_falta_funcionario.Close
    tbl_movimento_folha.Close
    tbl_tabela_folha.Close
    tbl_tabela_provento_desconto.Close
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = Mid(RetiraGString(1), 4, 7)
    txt_funcionario.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If SelecionaImpressoraEpson(Me) Then
        SelecionaPaginaContraCheque
        Call GravaAuditoria(1, Me.name, 7, "")
        Relatorio
        If l_funcionario = 0 Then
            cmd_sair.SetFocus
        Else
            dbcbo_funcionario.SetFocus
        End If
    End If
End Sub
Private Sub ImpRecibo()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_linha As Currency
    Dim x_cgc As String
    Dim x_provento As Currency
    Dim x_desconto As Currency
    Dim x_tot_provento As Currency
    Dim x_tot_desconto As Currency
    Dim x_salario_base As Currency
    Dim x_base_inss As Currency
    Dim x_base_fgts As Currency
    Dim x_fgts As Currency
    Dim x_base_irrf As Currency
    Dim x_nome As String
    Dim x_data As Date
    Dim x_quantidade As String
    Dim x_imprime As Boolean
    Dim x_13_impresso As Boolean
    l_local = 1
    x_linha = -0.8
    x_tot_provento = 0
    x_tot_desconto = 0
    x_tot_desconto = 0
        x_13_impresso = True
    If Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        x_13_impresso = False
    End If
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 8
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 5.5
    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Roman 10cpi"
    tbl_empresa.Seek "=", g_empresa
    If Not tbl_empresa.NoMatch Then
        x_cgc = Mid(tbl_empresa!CGC, 1, 2) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 3, 3) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 6, 3) & "/"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 9, 4) & "-"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 13, 2)
    End If
    Printer.FontName = "Roman 10cpi"
    Printer.FontBold = True
    ImprimeTexto "C.G.C.: " & x_cgc, 0, 9, x_linha + 0.8, l_local
    Printer.FontName = "Roman 6cpi"
    Printer.FontBold = True
    ImprimeTexto UCase(tbl_empresa!Nome), 0, 13, x_linha + 1.3, l_local
    Printer.FontBold = True
    Printer.FontName = "Roman 10cpi"
    If Val(Mid(l_ano_mes, 5, 2)) <= 12 Then
        x_data = "01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
        ImprimeTexto Trim(Format(x_data, "mmmm")) & "/" & Format(x_data, "yyyy"), 13.5, 17, x_linha + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        ImprimeTexto "Novembro/" & Mid(l_ano_mes, 1, 4), 13.5, 17, x_linha + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        ImprimeTexto "Dezembro/" & Mid(l_ano_mes, 1, 4), 13.5, 17, x_linha + 1.3, l_local
    End If
    Printer.FontBold = False
    ImprimeTexto Format(tbl_funcionario!Codigo, "000"), 0, 2, x_linha + 2.4, l_local
    Printer.FontBold = True
    ImprimeTexto tbl_funcionario!Nome, 1.1, 12, x_linha + 2.4, l_local
    Printer.FontBold = False
    ImprimeTexto UCase(tbl_funcionario!Cargo), 1.1, 12, x_linha + 3, l_local
    x_linha = 2.9
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Index = "id_funcionario"
            .Seek ">=", g_empresa, tbl_funcionario!Codigo, l_ano_mes, 0
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> tbl_funcionario!Codigo Or ![Mes Ano] <> l_ano_mes Then
                    Exit Do
                End If
                x_provento = 0
                x_desconto = 0
                tbl_tabela_provento_desconto.Seek "=", ![Codigo do Movimento]
                If Not tbl_tabela_provento_desconto.NoMatch Then
                    x_nome = tbl_tabela_provento_desconto!Nome
                    If tbl_tabela_provento_desconto![Provento ou Desconto] = "P" Then
                        x_provento = !valor
                        x_tot_provento = x_tot_provento + !valor
                    Else
                        x_desconto = !valor
                        x_tot_desconto = x_tot_desconto + !valor
                    End If
                Else
                    x_nome = "** Nao Cadastrado **"
                End If
                x_quantidade = ""
                If ![Codigo do Movimento] = 1 Then
                    x_quantidade = Format(!Quantidade, "#0")
                ElseIf ![Codigo do Movimento] = 5 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 10 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 15 Then
                    x_quantidade = Format(!Quantidade, "#0")
                ElseIf ![Codigo do Movimento] = 500 Then
                    x_quantidade = Format(!Quantidade, "#0.00") & "%"
                ElseIf ![Codigo do Movimento] = 505 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 520 Then
                    x_quantidade = Format(!Quantidade, "#0.00") & "%"
                ElseIf ![Codigo do Movimento] = 530 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 560 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                End If
                x_imprime = True
                If Val(Mid(![Mes Ano], 5, 2)) = 13 Then
                    x_imprime = False
                ElseIf Val(Mid(![Mes Ano], 5, 2)) = 14 And ![Codigo do Movimento] < 500 Then
                    x_imprime = False
                End If
                'Imprime 13o Salário 2a Parcela
                If Not x_13_impresso And ![Codigo do Movimento] >= 500 Then
                    x_13_impresso = True
                    x_quantidade = ![Codigo do Movimento]
                    .Index = "id_data"
                    Call BuscaRegistro2(l_ano_mes, l_funcionario, 1)
                    x_linha = x_linha + 0.5
                    Printer.FontName = "Roman 12cpi"
                    ImprimeTexto "13o Salario", 1.1, 7.8, x_linha, l_local
                    ImprimeValor Format(!Quantidade, "#0") & "/12", 8.7, 10.2, x_linha, l_local
                    ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
                    .Index = "id_funcionario"
                    .Seek ">=", g_empresa, l_funcionario, l_ano_mes, Val(x_quantidade)
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                End If
                If x_imprime Then
                    x_linha = x_linha + 0.5
                    ImprimeTexto Format(![Codigo do Movimento], "000"), 0, 2, x_linha, l_local
                    Printer.FontName = "Roman 12cpi"
                    ImprimeTexto x_nome, 1.1, 7.8, x_linha, l_local
                    ImprimeValor x_quantidade, 8.7, 10.2, x_linha, l_local
                    Printer.FontName = "Roman 10cpi"
                    If x_provento > 0 Then
                        ImprimeValor Format(x_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
                    Else
                        ImprimeValor Format(x_desconto, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
                    End If
                End If
                .MoveNext
            Loop
            If Val(Mid(l_ano_mes, 5, 2)) = 13 Then
                .Index = "id_data"
                Call BuscaRegistro2(l_ano_mes, tbl_funcionario!Codigo, 1)
                x_linha = 2.9
                x_linha = x_linha + 0.5
                Printer.FontName = "Roman 12cpi"
                ImprimeTexto "13o Salario - 1a PARCELA", 1.1, 7.8, x_linha, l_local
                ImprimeValor Format(!Quantidade, "#0") & "/12", 8.7, 10.2, x_linha, l_local
                ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
            End If
            .Index = "id_data"
            Call BuscaRegistro2(l_ano_mes, tbl_funcionario!Codigo, l_codigo)
            x_linha = 10.3
            Printer.FontBold = True
            ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
            ImprimeValor Format(x_tot_desconto, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
            x_tot_provento = x_tot_provento - x_tot_desconto
            x_linha = 11.1
            ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
            Printer.FontBold = False
            x_linha = -0.8
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 1]), 0, 10.2, x_linha + 10.8, l_local
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 2]), 0, 10.2, x_linha + 11.3, l_local
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 3]), 0, 10.2, x_linha + 11.8, l_local
        End If
    End With
    tbl_tabela_provento_desconto.Seek "=", 1
    x_salario_base = CalculaBaseCalculo(1)
    tbl_tabela_provento_desconto.Seek "=", 520
    x_base_inss = CalculaBaseCalculo(520)
    x_base_fgts = CalculaBaseCalculo(520)
    If Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        x_base_fgts = x_base_fgts / 2
    End If
    x_fgts = Format(x_base_fgts * 8 / 100, "###,##0.00")
    x_base_irrf = 0
    x_linha = 11.9
    ImprimeValor Format(x_salario_base, "##,###,##0.00"), 0, 2.9, x_linha, l_local
    ImprimeValor Format(x_base_inss, "##,###,##0.00"), 3.4, 6.3, x_linha, l_local
    ImprimeValor Format(x_base_fgts, "##,###,##0.00"), 6.5, 9.4, x_linha, l_local
    ImprimeValor Format(x_fgts, "##,###,##0.00"), 9.2, 12.1, x_linha, l_local
    ImprimeValor Format(x_base_irrf, "##,###,##0.00"), 12.1, 15, x_linha, l_local
    Printer.EndDoc
End Sub
Private Sub Relatorio()
    Dim x_mes_ano As String
    If Val(Mid(l_ano_mes, 5, 2)) < 13 Then
        x_mes_ano = Format(CDate("01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)), "mmmm") & " / " & Format(CDate("01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)), "yyyy")
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        x_mes_ano = "13o Salario - 1a Parcela"
    End If
    With tbl_funcionario
        If .RecordCount > 0 Then
            If l_funcionario = 0 Then
                If (MsgBox("Será impresso todos os recibos da folha de pagamento." & Chr(10) & "Empresa: " & g_nome_empresa & Chr(10) & "Referência: " & x_mes_ano & Chr(10) & Chr(10) & "Confirme para imprimi-los.", vbOKCancel + vbDefaultButton1, "Emite Todos os Recibos!") <> 1) Then
                    Exit Sub
                End If
            End If
            If l_funcionario = 0 Then
                .Index = "id_nome"
                .Seek ">=", g_empresa, " ", 0
            Else
                .Seek "=", g_empresa, l_funcionario
            End If
            If Not .NoMatch Then
                Do Until .EOF
                    If ExisteMovimento(l_ano_mes, !Codigo) Then
                        If ![Serie da Carteira de Trabalho] <> "NR" Then
                            ImpRecibo
                        End If
                    Else
                        If l_funcionario > 0 Then
                            MsgBox "O funcionário " & Trim(tbl_funcionario!Nome) & ", não tem movimento no período informado.", vbInformation, "Sem Movimento!"
                        End If
                    End If
                    If l_funcionario > 0 Then
                        Exit Do
                    End If
                    If !Empresa <> g_empresa Then
                        Exit Do
                    End If
                    .MoveNext
                Loop
            End If
        End If
        .Index = "id_codigo"
    End With
End Sub
Function RetiraString(x_string As String, numero As Integer) As Integer
    RetiraString = 0
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(x_string) > 0 Then
        Do Until x_index > Len(x_string)
            x_index = x_index + 1
            If Mid(x_string, x_index, 1) = "@" Then
                If x_numero = numero Then
                    RetiraString = Mid(x_string, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Private Sub SelecionaPaginaContraCheque()
Dim Impressora As Printer
Dim NumeroFormulario As Integer
    On Error GoTo PrinterError
    NumeroFormulario = 256
    Printer.PaperSize = NumeroFormulario
    Printer.ScaleMode = 5
'    For NumeroFormulario = 1 To 256
'        Printer.PaperSize = NumeroFormulario
'        'Seleciona largura do formulário
'        Printer.ScaleWidth = 8
'        'Seleciona altura do formulário
'        Printer.ScaleHeight = 5.5
'    Next
    Exit Sub
PrinterError:
    NumeroFormulario = NumeroFormulario + 1
    Resume



    NumeroFormulario = 256
    Resume
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate("01/" & msk_data) Then
        MsgBox "Informe o mês/ano.", 64, "Atenção!"
        msk_data.SetFocus
    'ElseIf Not Val(dbcbo_funcionario.BoundText) > 0 Then
    '    MsgBox "Selecione o funcionário.", 64, "Atenção!"
    '    dbcbo_funcionario.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If dbcbo_funcionario.BoundText <> "" Then
        l_funcionario = Val(dbcbo_funcionario.BoundText)
        txt_funcionario = dbcbo_funcionario.BoundText
        txt_funcionario_LostFocus
        cmd_imprimir.SetFocus
    Else
        l_funcionario = 0
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If msk_data.Text = "__/____" Then
        dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = " & Chr(34) & "A" & Chr(34) & " And [Salario Base] > " & 0 & " Order By [Nome]"
        dta_funcionario.Refresh
        Screen.MousePointer = 1
        msk_data.Text = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
        dbcbo_funcionario.BoundText = ""
        msk_data.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_falta_funcionario = bd_sgp.OpenRecordset("Movimento_Falta_Funcionario", dbOpenTable)
    Set tbl_movimento_folha = bd_sgp.OpenTable("Movimento_Folha")
    Set tbl_tabela_folha = bd_sgp.OpenTable("Tabela_Folha")
    Set tbl_tabela_provento_desconto = bd_sgp.OpenTable("Tabela_Provento_Desconto")
    tbl_empresa.Index = "id_codigo"
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_falta_funcionario.Index = "id_funcionario"
    tbl_movimento_folha.Index = "id_data"
    tbl_tabela_folha.Index = "id_mes_ano"
    tbl_tabela_provento_desconto.Index = "id_codigo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    tbl_tabela_folha.Seek "=", Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    If tbl_tabela_folha.NoMatch Then
        MsgBox "Tabela da folha não cadastrada.", 64, "Erro de Consistência!"
        msk_data.SetFocus
    Else
        l_ano_mes = tbl_tabela_folha![Mes Ano]
    End If
End Sub
Private Sub txt_funcionario_GotFocus()
    txt_funcionario.SelStart = 0
    txt_funcionario.SelLength = Len(txt_funcionario)
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    l_funcionario = Val(txt_funcionario)
    If Val(txt_funcionario) > 0 Then
        tbl_funcionario.Seek "=", g_empresa, Val(txt_funcionario)
        If Not tbl_funcionario.NoMatch Then
            If tbl_funcionario!Situacao = "I" Then
                MsgBox "O funcionário " & Trim(tbl_funcionario!Nome) & " está inativo.", 64, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dbcbo_funcionario.BoundText = tbl_funcionario!Codigo
                cmd_imprimir.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", 64, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    Else
        dbcbo_funcionario.BoundText = ""
    End If
End Sub
