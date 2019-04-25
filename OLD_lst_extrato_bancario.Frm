VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form emissao_extrato_bancario 
   Caption         =   "Listagem do Extrato Bancário"
   ClientHeight    =   2355
   ClientLeft      =   90
   ClientTop       =   5400
   ClientWidth     =   6975
   Icon            =   "lst_extrato_bancario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_extrato_bancario.frx":030A
   ScaleHeight     =   2355
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1800
      Picture         =   "lst_extrato_bancario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime o estrato bancário."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4380
      Picture         =   "lst_extrato_bancario.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1380
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6120
         Picture         =   "lst_extrato_bancario.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_extrato_bancario.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_extrato_bancario.frx":55A0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   5100
         TabIndex        =   8
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
      Begin MSMask.MaskEdBox msk_data 
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
         Caption         =   "Data Final"
         Height          =   315
         Left            =   4020
         TabIndex        =   7
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "emissao_extrato_bancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lColuna(0 To 5) As Currency
Dim lColunaI As Currency
Dim lEntreLinha As Currency
Dim lLinha As Integer
Dim lLinhaI As Currency
Dim lLinhaTab As Currency
Dim lLocal As Integer
Dim lPagina As Integer
Dim lTamanhoCabDados As Currency
Dim lTamanhoCabTitulo As Currency
'Fim de variáveis padrão para relatório
Dim l_saldo_anterior As Currency
Dim l_saldo_atual As Currency
Dim l_data_anterior As String
Dim l_data_atual As String
Dim l_historico As String
Dim l_codigo_historico As Integer
Dim tbl_movimento_bancario As Table
Private Sub ZeraVariaveis()
    lColunaI = 0
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 1.7
    lColuna(2) = lColunaI + 9.5
    lColuna(3) = lColunaI + 15
    lColuna(4) = lColunaI + 17
    lColuna(5) = lColunaI + 20
    lEntreLinha = 0.5
    lLinha = 0
    lLinhaI = 1
    lLinhaTab = 0
    lLocal = 1
    lPagina = 0
    lTamanhoCabDados = 0.8
    lTamanhoCabTitulo = 2.5
End Sub
Private Sub Finaliza()
    tbl_historico.Close
    tbl_saldo_conta.Close
    tbl_movimento_bancario.Close
End Sub
Private Sub ImpDados()
    Dim x_valor As Currency
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = True
    'calcula linha
    lLinha = 1
    lLinhaTab = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados + lLinha * lEntreLinha - 0.3
    'teste para imprimir letra correta
    Printer.FontBold = False
    ImprimeTexto "  ", lColuna(3), lColuna(4), lLinhaTab, lLocal
    Printer.FontBold = True
    'imprime saldo de abertura
    ImprimeCentralizado l_data_anterior, lColuna(0), lColuna(1), lLinhaTab, lLocal
    ImprimeTexto "*** Saldo Anterior...", lColuna(1), lColuna(2), lLinhaTab, lLocal
    ImprimeValor Format(l_saldo_anterior, "###,##0.00;-###,##0.00"), lColuna(4), lColuna(5), lLinhaTab, lLocal
    'loop movimento de conta bancária
    Do Until tbl_movimento_bancario.EOF
        If tbl_movimento_bancario!Empresa <> g_empresa Or tbl_movimento_bancario![Codigo da Conta] <> g_conta_bancaria Then
            Exit Do
        End If
        If tbl_movimento_bancario!Data > CDate(msk_data_f) Then
            Exit Do
        End If
        If lLinha = 30 Then
            ImpGrade
            Printer.NewPage
        End If
        'calcula linha
        lLinha = lLinha + 1
        lLinhaTab = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados + lLinha * lEntreLinha - 0.3
        'Calcula saldo atual
        x_valor = tbl_movimento_bancario!valor
        If tbl_movimento_bancario![Debito ou Credito] = "D" Then
            x_valor = -x_valor
        End If
        l_saldo_atual = l_saldo_atual + x_valor
        l_data_atual = tbl_movimento_bancario!Data
        'le tabela auxiliar
        tbl_historico.Index = "id_codigo"
        l_codigo_historico = tbl_movimento_bancario![Codigo do Historico]
        tbl_historico.Seek "=", g_empresa, l_codigo_historico
        If Not tbl_historico.NoMatch Then
            MontaHistorico g_empresa, tbl_historico!Codigo
        End If
        'imprime detalhe
        If lLinha = 1 Then
            Printer.FontBold = False
            ImprimeTexto "  ", lColuna(1), lColuna(2), lLinhaTab, lLocal
            Printer.FontBold = True
        End If
        ImprimeCentralizado Format(tbl_movimento_bancario!Data, "dd/mm/yyyy"), lColuna(0), lColuna(1), lLinhaTab, lLocal
        ImprimeTexto l_historico, lColuna(1), lColuna(2), lLinhaTab, lLocal
        ImprimeTexto tbl_movimento_bancario!Complemento, lColuna(2), lColuna(3), lLinhaTab, lLocal
        ImprimeTexto tbl_movimento_bancario![Numero do Documento], lColuna(3), lColuna(4), lLinhaTab, lLocal
        ImprimeValor Format(x_valor, "###,##0.00;-###,##0.00"), lColuna(4), lColuna(5), lLinhaTab, lLocal
        tbl_movimento_bancario.MoveNext
    Loop
    'calcula linha
    lLinha = lLinha + 1
    lLinhaTab = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados + lLinha * lEntreLinha - 0.3
    'imprime saldo atual
    ImprimeCentralizado l_data_atual, lColuna(0), lColuna(1), lLinhaTab, lLocal
    Printer.FontSize = 12
    Printer.DrawWidth = 2
    Printer.FontBold = True
    ImprimeTexto "*** Saldo Atual...", lColuna(1), lColuna(2), lLinhaTab, lLocal
    ImprimeValor Format(l_saldo_atual, "###,##0.00;-###,##0.00"), lColuna(4), lColuna(5), lLinhaTab, lLocal
    ImpGrade
    Printer.EndDoc
End Sub
Private Sub MontaHistorico(empresa_x As Integer, conta_x As Integer)
    Dim pai As String
    Dim filho As String
    tbl_historico.Index = "id_codigo"
    tbl_historico.Seek "=", empresa_x, conta_x
    If Not tbl_historico.NoMatch Then
        l_historico = Trim(tbl_historico![Nome Filho])
        pai = tbl_historico![Nome Pai]
        Do Until Trim(tbl_historico![Nome Pai]) = "\"
            tbl_historico.Index = "id_filho2"
            tbl_historico.Seek ">=", empresa_x, pai, filho
            If Not tbl_historico.NoMatch Then
                If Trim(tbl_historico![Nome Filho]) = Trim(pai) Then
                    l_historico = Trim(tbl_historico![Nome Filho]) & "\" & l_historico
                    pai = tbl_historico![Nome Pai]
                End If
            End If
        Loop
    End If
End Sub
Private Sub ImpGrade()
    Dim x_linha_i As Currency
    Dim x_linha_f As Currency
    lPagina = lPagina + 1
    ImpCabTitulo
    ImpCabDados
    x_linha_i = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados
    x_linha_f = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados + lLinha * lEntreLinha + 0.3
    Printer.DrawWidth = 2
    Printer.Line (lColuna(0), x_linha_i)-(lColuna(0), x_linha_f)
    Printer.Line (lColuna(1), x_linha_i)-(lColuna(1), x_linha_f)
    Printer.Line (lColuna(2), x_linha_i)-(lColuna(2), x_linha_f)
    Printer.Line (lColuna(3), x_linha_i)-(lColuna(3), x_linha_f)
    Printer.Line (lColuna(4), x_linha_i)-(lColuna(4), x_linha_f)
    Printer.Line (lColuna(5), x_linha_i)-(lColuna(5), x_linha_f)
    Printer.Line (lColuna(0), x_linha_f)-(lColuna(5), x_linha_f)
    lLinha = 0
End Sub
Private Sub ImpCabDados()
    Dim x_linha_i As Currency
    Dim x_linha_f As Currency
    x_linha_i = lLinhaI + lTamanhoCabTitulo
    x_linha_f = lLinhaI + lTamanhoCabTitulo + lTamanhoCabDados
    Printer.DrawWidth = 2
    Printer.Line (lColuna(0), x_linha_i)-(lColuna(0), x_linha_f)
    Printer.Line (lColuna(1), x_linha_i)-(lColuna(1), x_linha_f)
    Printer.Line (lColuna(2), x_linha_i)-(lColuna(2), x_linha_f)
    Printer.Line (lColuna(3), x_linha_i)-(lColuna(3), x_linha_f)
    Printer.Line (lColuna(4), x_linha_i)-(lColuna(4), x_linha_f)
    Printer.Line (lColuna(5), x_linha_i)-(lColuna(5), x_linha_f)
    Printer.Line (lColuna(0), x_linha_f)-(lColuna(5), x_linha_f)
    x_linha_i = lLinhaI + lTamanhoCabTitulo + 0.2
    Printer.FontSize = 10
    Printer.DrawWidth = 2
    Printer.FontBold = True
    ImprimeCentralizado "Data", lColuna(0), lColuna(1), x_linha_i, lLocal
    ImprimeCentralizado "Histórico", lColuna(1), lColuna(2), x_linha_i, lLocal
    ImprimeCentralizado "Complemento", lColuna(2), lColuna(3), x_linha_i, lLocal
    ImprimeCentralizado "N. Doc.", lColuna(3), lColuna(4), x_linha_i, lLocal
    ImprimeCentralizado "Valor", lColuna(4), lColuna(5), x_linha_i, lLocal
End Sub
Private Sub ImpCabTitulo()
    Printer.DrawWidth = 2
    Printer.Line (lColuna(0), lLinhaI)-(lColuna(5), lLinhaI)
    Printer.Line (lColuna(0), lLinhaI)-(lColuna(0), lLinhaI + lTamanhoCabTitulo)
    Printer.Line (lColuna(5), lLinhaI)-(lColuna(5), lLinhaI + lTamanhoCabTitulo)
    Printer.Line (lColuna(0), lLinhaI + lTamanhoCabTitulo)-(lColuna(5), lLinhaI + lTamanhoCabTitulo)
    Printer.FontSize = 10
    Printer.DrawWidth = 2
    Printer.FontBold = True
    ImprimeTexto "EXTRATO BANCÁRIO", lColuna(0) + 1, lColuna(5), lLinhaI + 0.3, lLocal
    ImprimeTexto "Empresa.: " & g_nome_empresa, lColuna(0) + 1, lColuna(5), lLinhaI + 0.9, lLocal
    ImprimeTexto "Conta......: " & g_nome_conta, lColuna(0) + 1, lColuna(5), lLinhaI + 1.5, lLocal
    ImprimeTexto "Período...: " & msk_data_i & " a " & msk_data_f, lColuna(0) + 1, lColuna(5), lLinhaI + 2.1, lLocal
    ImprimeTexto "Página,   " & Format(lPagina, "000"), lColuna(0) + 16.5, lColuna(5), lLinhaI + 0.3, lLocal
    ImprimeTexto "Goiânia, " & Format(msk_data, "dd/mm/yyyy"), lColuna(0) + 16.5, lColuna(5), lLinhaI + 0.9, lLocal
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'busca saldo anterior
    tbl_saldo_conta.Index = "id_data"
    tbl_saldo_conta.Seek "<", g_empresa, g_conta_bancaria, msk_data_i
    If Not tbl_saldo_conta.NoMatch Then
        If tbl_saldo_conta!Empresa = g_empresa And tbl_saldo_conta![Codigo da Conta] = g_conta_bancaria Then
            l_saldo_anterior = tbl_saldo_conta!Saldo
            l_data_anterior = tbl_saldo_conta!Data
        Else
            l_saldo_anterior = 0
        End If
    End If
    l_saldo_atual = l_saldo_anterior
    'busca saldo anterior
    tbl_movimento_bancario.Index = "id_data"
    tbl_movimento_bancario.Seek ">", g_empresa, g_conta_bancaria, msk_data_i, 0
    If Not tbl_movimento_bancario.NoMatch Then
        If tbl_movimento_bancario!Empresa = g_empresa And tbl_movimento_bancario![Codigo da Conta] = g_conta_bancaria Then
            ImpDados
        End If
    End If
    cmd_sair.SetFocus
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
        If SelecionaImpressoraHP(Me) Then
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
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data) Then
        msk_data = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.SetFocus
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
    Set tbl_historico = bd_sgp.OpenTable("Historico")
    Set tbl_saldo_conta = bd_sgp.OpenTable("Saldo_Conta")
    Set tbl_movimento_bancario = bd_sgp.OpenTable("Movimento_Bancario")
    tbl_historico.Index = "id_codigo"
    tbl_saldo_conta.Index = "id_data"
    tbl_movimento_bancario.Index = "id_conta"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
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
