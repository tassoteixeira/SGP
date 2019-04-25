VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmissaoAnaliseMovEstoque 
   Caption         =   "Emissão da Análise da Movimentação do Estoque"
   ClientHeight    =   3075
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "EmissaoAnaliseMovEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "EmissaoAnaliseMovEstoque.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "EmissaoAnaliseMovEstoque.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime análise da venda de produtos."
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "EmissaoAnaliseMovEstoque.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza análise da venda de produtos."
      Top             =   2100
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "EmissaoAnaliseMovEstoque.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "EmissaoAnaliseMovEstoque.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6060
         Picture         =   "EmissaoAnaliseMovEstoque.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
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
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   4875
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
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
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
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1500
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
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoAnaliseMovEstoque"
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
Dim l_total_quantidade As Currency
Dim l_total_custo As Currency
Dim l_total_venda As Currency
Dim lSQL As String

Dim rstTabela As adodb.Recordset
Private Estoque2 As New cEstoque2
Private EntradaProduto As New cEntradaProduto
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
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
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    l_total_quantidade = 0
    l_total_custo = 0
    l_total_venda = 0
End Sub
Private Sub PreencheCboGrupo()
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    
    lSQL = "SELECT Codigo, Nome FROM Grupo ORDER BY Nome"
    Set rstTabela = Conectar.RsConexao(lSQL)
    With rstTabela
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_grupo.AddItem !Nome
                cbo_grupo.ItemData(cbo_grupo.NewIndex) = !Codigo
                .MoveNext
            Loop
            rstTabela.Close
        End If
    End With
    Set rstTabela = Nothing
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
    ZeraVariaveis
    
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Produto.Unidade"
    lSQL = lSQL & "  FROM Produto"
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & " WHERE Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
    lSQL = lSQL & " ORDER BY Produto.Nome"
    'Abre RecordSet
    Set rstTabela = New adodb.Recordset
    Set rstTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rstTabela.RecordCount > 0 Then
        LoopTabelaProduto
        If lPagina > 0 Then
            ImpTotal
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Análise da Movimentação do Estoque|@|"
            frm_preview.Show 1
        End If
    End If
    If rstTabela.State = 1 Then
        rstTabela.Close
    End If
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela produto
    Dim xLinha As String
    Dim xEstoqueAnterior As Currency
    Dim xEntrada As Currency
    Dim xSaida As Currency
    Dim xEstoqueAtual As Currency
    With rstTabela
        .MoveFirst
        Do Until .EOF
            xEstoqueAnterior = 0
            If Estoque2.LocalizarCodigo(g_empresa, CDate(msk_data_i.Text) - 1, rstTabela("Codigo").Value) Then
                xEstoqueAnterior = Estoque2.Quantidade
            End If
            xEntrada = EntradaProduto.TotalQtdProdutoDatas(g_empresa, rstTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text))
            xSaida = MovimentoCupomFiscal.QuantidadeProdutoVendaData(g_empresa, rstTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 1, 9, 0)
            xEstoqueAtual = xEstoqueAnterior + xEntrada - xSaida
            If xEntrada > 0 Or xSaida > 0 Or xEstoqueAtual > 0 Then
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 60 Then
                    xLinha = "+------+-------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
                    Mid(xLinha, 12, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                Call ImpDet(xEstoqueAnterior, xEntrada, xSaida, xEstoqueAtual)
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    
    xLinha = "+------+-------------------------------+---+--------+--------+--------+--------+"
    Mid(xLinha, 10, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(ByVal pAbertura As Currency, ByVal pEntrada As Currency, ByVal pSaida As Currency, ByVal pAtual As Currency)
    Dim xLinha As String
    Dim x_percentual As Currency
    Dim i As Integer
    'xLinha= "         1         2         3         4         5         6         7         8"
    'xLinha= "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    xLinha = "|      |                               |   |        |        |        |        |"
    
    i = Len(Format(rstTabela("Codigo").Value, "##,000"))
    Mid(xLinha, 2 + 6 - i, i) = Format(rstTabela("Codigo").Value, "##,000")
    Mid(xLinha, 9, 31) = rstTabela("Nome").Value
    Mid(xLinha, 41, 3) = rstTabela("Unidade").Value
    i = Len(Format(pAbertura, "####,##0"))
    Mid(xLinha, 45 + 8 - i, i) = Format(pAbertura, "####,##0")
    i = Len(Format(pEntrada, "####,##0"))
    Mid(xLinha, 54 + 8 - i, i) = Format(pEntrada, "####,##0")
    i = Len(Format(pSaida, "####,##0"))
    Mid(xLinha, 63 + 8 - i, i) = Format(pSaida, "####,##0")
    i = Len(Format(pAtual, "####,##0"))
    Mid(xLinha, 72 + 8 - i, i) = Format(pAtual, "####,##0")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    If lPagina = 0 Then
        'seleciona medidas para centímetros
        lNomeArquivo = BioCriaImprime
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
    Printer.CurrentY = 0
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| ANÁLISE DA MOVIMENTAÇÃO DO ESTOQUE                              , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____       PERÍODO.: _ AO _     |"
    Mid(xLinha, 29, 10) = msk_data_i.Text
    Mid(xLinha, 42, 10) = msk_data_f.Text
    Mid(xLinha, 69, 1) = cbo_periodo_i.Text
    Mid(xLinha, 74, 1) = cbo_periodo_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| GRUPO.:                                                                      |"
    Mid(xLinha, 11, 3) = Format(cbo_grupo.ItemData(cbo_grupo.ListIndex), "000")
    Mid(xLinha, 15, 30) = cbo_grupo.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------+-------------------------------+---+--------+--------+--------+--------+"
    BioImprime "@Printer.Print " & "|      |                               |   | QUANT. | ENTRADA| SAIDAS | QUANT. |"
    BioImprime "@Printer.Print " & "|CÓDIGO|DISCRIMINAÇÃO DOS PRODUTOS     |UN.| ESTOQUE|   NO   |   NO   | ESTOQUE|"
    BioImprime "@Printer.Print " & "|      |                               |   |ANTERIOR| PERIODO| PERIODO|  ATUAL |"
    BioImprime "@Printer.Print " & "+------+-------------------------------+---+--------+--------+--------+--------+"
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
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
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", vbInformation, "Atenção!"
        cbo_grupo.SetFocus
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
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 3
        cbo_grupo.ListIndex = 0
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
    PreencheCboPeriodo
    PreencheCboGrupo
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
