VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_movimento_cartao 
   Caption         =   "Emissão dos Cartões de Crédito"
   ClientHeight    =   3975
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_movimento_cartao.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_movimento_cartao.frx":030A
   ScaleHeight     =   3975
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_movimento_cartao.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Visualiza movimento de cartão de crédito."
      Top             =   3000
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_movimento_cartao.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime movimento de cartão de crédito."
      Top             =   3000
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_movimento_cartao.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3000
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_movimento_cartao.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_movimento_cartao.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_movimento_cartao.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_cartao_credito 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2280
         Width           =   4755
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1860
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
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
      Begin VB.Label Label7 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Período &Final"
         Height          =   315
         Left            =   4380
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Cartão"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Período Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_movimento_cartao"
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
Dim lTotal As Currency

Dim lSQL As String
Private rsTabela As New adodb.Recordset

Private CartaoCredito As New cCartaoCredito

Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set CartaoCredito = Nothing
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
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "0 Todos os Caixas"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 0
    cbo_tipo_movimento.AddItem "1 Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de Óleo/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Caixa da Borr./Lavador"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub PreencheCboCartaoCredito()
    cbo_cartao_credito.Clear
    cbo_cartao_credito.AddItem "Todos os Cartões"
    cbo_cartao_credito.ItemData(cbo_cartao_credito.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM Cartao_Credito"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_cartao_credito.AddItem rsTabela("Nome").Value
            cbo_cartao_credito.ItemData(cbo_cartao_credito.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    lSQL = ""
    lSQL = lSQL & "SELECT [Data de Emissao], Periodo, [Numero do Lancamento], [Data do Vencimento], valor, [Numero do Cartao], [Codigo do Cartao]"
    lSQL = lSQL & "  FROM Movimento_Cartao_Credito"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data de Emissao] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data de Emissao] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND Periodo >= " & preparaTexto(Val(cbo_periodo_i.Text))
    lSQL = lSQL & "   AND Periodo <= " & preparaTexto(Val(cbo_periodo_f.Text))
    If cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex) > 0 Then
        lSQL = lSQL & "   AND [Codigo do Cartao] = " & cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex)
    End If
    If cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) > 0 Then
        lSQL = lSQL & "   AND [Tipo do Movimento] = " & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    End If
    lSQL = lSQL & " ORDER BY [Data de Emissao]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        ImpDados
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    LoopMovimentoCartaoCredito
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Movimento com Cartão de Crédito|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopMovimentoCartaoCredito()
    'loop movimento do cartao de credito
    Dim x_linha As String
    Dim x_nome_cartao As String * 40
    
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
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
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        'Le tabela auxiliar
        If CartaoCredito.LocalizarCodigo(rsTabela("Codigo do Cartao").Value) Then
            x_nome_cartao = CartaoCredito.Nome
        Else
            x_nome_cartao = "** Não Cadastrado **"
        End If
        Call ImpDet(rsTabela("Data de Emissao").Value, rsTabela("Periodo").Value, rsTabela("Numero do Lancamento").Value, x_nome_cartao, rsTabela("Data do Vencimento").Value, rsTabela("valor").Value, rsTabela("Numero do Cartao").Value)
        lTotal = lTotal + rsTabela("valor").Value
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpDet(x_data_emissao As Date, x_periodo As String, x_numero_lancamento As Integer, x_nome_cartao As String, x_data_vencimento As Date, x_valor As Currency, x_numero_cartao As Integer)
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
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
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
    BioImprime "@Printer.Print " & x_linha
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 40, 11) = "*** TOTAL.:"
    Mid(x_linha, 52, 1) = "|"
    i = Len(Format(lTotal, "#,###,##0.00"))
    Mid(x_linha, 57 + 12 - i, i) = Format(lTotal, "#,###,##0.00")
    Mid(x_linha, 70, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print x_linha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & "  "
    BioImprime "@@Printer.FontBold = False"
    x_linha = String(137, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 52, 1) = "+"
    Mid(x_linha, 70, 1) = "+"
    Mid(x_linha, 95, 22) = " Cerrado Informática. "
    Mid(x_linha, 137, 1) = "+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_string_137 As String
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
    x_string_137 = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_string_137, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_string_137
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "| RELAÇÃO DO MOVIMENTO DE CARTÃO DE CRÉDITO                Goiânia, " & msk_data & " |"
    x_string_137 = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_string_137, 29, 10) = msk_data_i
    Mid(x_string_137, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_string_137
    x_string_137 = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_string_137, 29, 1) = cbo_periodo_i
    Mid(x_string_137, 49, 1) = cbo_periodo_f
    BioImprime "@Printer.Print " & x_string_137
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+-------+------+-------------+----------+-----------------+----------------------+-------------------------------------------+"
    BioImprime "@Printer.Print " & "| EMISSÃO  |PERIODO|N.LANC|    CARTÃO   |VENCIMENTO| VALOR DO CARTÃO | NÚMERO DO CARTÃO     |                                           |"
    BioImprime "@Printer.Print " & "+----------+-------+------+-------------+----------+-----------------+----------------------+-------------------------------------------+"
End Sub
Private Sub cbo_cartao_credito_GotFocus()
    SendMessageLong cbo_cartao_credito.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_cartao_credito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_GotFocus()
    SendMessageLong cbo_periodo_i.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_cartao_credito.SetFocus
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
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
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
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Escolha o período inicial.", 64, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Escolha o período final.", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "O periodo final deve ser maior que " & Val(cbo_periodo_i) - 1 & ".", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
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
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
        cbo_cartao_credito.ListIndex = 0
        cbo_periodo_i.SetFocus
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
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    PreencheCboCartaoCredito
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
