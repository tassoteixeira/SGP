VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lstComparaCupomLubrif 
   Caption         =   "Compara Venda Lubrificante x Cupom Fiscal"
   ClientHeight    =   3885
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lstComparaCupomLubrif.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lstComparaCupomLubrif.frx":030A
   ScaleHeight     =   3885
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lstComparaCupomLubrif.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lstComparaCupomLubrif.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lstComparaCupomLubrif.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2940
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lstComparaCupomLubrif.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lstComparaCupomLubrif.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cboTipoEstoque 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Width           =   4755
      End
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo da Loja"
         Height          =   255
         Left            =   3660
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo do Posto"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   1080
         Width           =   1875
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lstComparaCupomLubrif.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   4755
      End
      Begin VB.ComboBox cbo_produto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1860
         Width           =   4755
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
      Begin MSMask.MaskEdBox txtDataFinal 
         Height          =   315
         Left            =   4860
         TabIndex        =   18
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
      Begin MSMask.MaskEdBox txtDataInicial 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo de Estoque"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "I&mprimir Produto"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Produto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lstComparaCupomLubrif"
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
Dim lQuantidade As Currency
Dim lTotal As Currency
Dim lCodigoGrupo As Integer

Dim lSQL As String
Private rsTabela As New adodb.Recordset

Private Aliquota As New cAliquota
Private Grupo As New cGrupo
Private Produto As New cProduto
Private SubEstoque As New cSubEstoque



Private MovCupomFiscal As New cMovimentoCupomFiscal
Private MovLubrificante As New cMovimentoLubrificante
Dim lQtdCupom As Currency
Dim lQtdLubrif As Currency

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
    Set Aliquota = Nothing
    Set Grupo = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
End Sub
Private Sub PreencheCboGrupo()
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_grupo.AddItem rsTabela("Nome").Value
            cbo_grupo.ItemData(cbo_grupo.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub PreencheCboProduto()
    cbo_produto.Clear
    cbo_produto.AddItem "Todos os Produtos"
    cbo_produto.ItemData(cbo_produto.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_produto.AddItem rsTabela("Nome").Value
            cbo_produto.ItemData(cbo_produto.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub PreencheCboTipoEstoque()
    cboTipoEstoque.Clear
    cboTipoEstoque.AddItem "Geral"
    cboTipoEstoque.ItemData(cboTipoEstoque.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM TipoSubEstoque"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cboTipoEstoque.AddItem rsTabela("Nome").Value
            cboTipoEstoque.ItemData(cboTipoEstoque.NewIndex) = rsTabela("Codigo").Value
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
    lQuantidade = 0
    lTotal = 0
    lCodigoGrupo = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Estoque.Quantidade, Produto.[Preco de Custo], Estoque.[Preco de Venda], Produto.Unidade, Produto.[Codigo da Aliquota]"
    lSQL = lSQL & "  FROM Produto, Estoque"
    lSQL = lSQL & " WHERE Estoque.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Estoque.[Codigo do Produto2] = Produto.Codigo"
    lSQL = lSQL & "   AND Produto.Inativo = " & preparaBooleano(False)
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
    If cbo_produto.ItemData(cbo_produto.ListIndex) > 0 Then
        lSQL = lSQL & "   AND Produto.Codigo = " & cbo_produto.ItemData(cbo_produto.ListIndex)
    End If
    If chkExclusivoPosto.Value = 1 And chkExclusivoLoja.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    If chkExclusivoLoja.Value = 1 And chkExclusivoPosto.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
    lSQL = lSQL & " ORDER BY [Codigo do Grupo], Produto.Nome"
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
    LoopTabelaProduto
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Inventário de Produtos|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela produto
    Dim x_linha As String
    Dim x_quantidade As Currency
    Dim x_valor1 As Currency
    Dim x_valor2 As Currency
    Dim xAliquota As Currency
    Dim xEstoque As Currency
    
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 60 Then
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            Mid(x_linha, 15, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        x_valor1 = rsTabela("Preco de Venda").Value
        
        xEstoque = rsTabela("Quantidade").Value
        If cboTipoEstoque.ItemData(cboTipoEstoque.ListIndex) > 0 Then
            If SubEstoque.LocalizarCodigo(g_empresa, rsTabela("Codigo").Value, cboTipoEstoque.ItemData(cboTipoEstoque.ListIndex)) Then
                xEstoque = SubEstoque.Quantidade
            End If
        End If
        
        x_valor2 = x_valor1 * xEstoque
        If lCodigoGrupo <> rsTabela("Codigo do Grupo").Value Then
            lCodigoGrupo = rsTabela("Codigo do Grupo").Value
            ImpGrupo
        End If
        xAliquota = 0
        If Aliquota.LocalizarCodigoAliquota(rsTabela("Codigo da Aliquota").Value) Then
            xAliquota = Aliquota.Aliquota
        End If
        Call ImpDet(rsTabela("Codigo").Value, rsTabela("Nome").Value, rsTabela("unidade").Value, xEstoque, x_valor1, x_valor2, xAliquota)
        If xEstoque > 0 Then
            lTotal = lTotal + x_valor2
            lQuantidade = lQuantidade + x_quantidade
        End If
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpDet(x_codigo As Long, x_nome As String, x_unidade As String, x_quantidade As Currency, x_valor1 As Currency, x_valor2 As Currency, xAliquotaImposto As Currency)
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "|           |                                              |         |                 |                    |                    |      |"
    i = Len(Format(x_codigo, "#,000"))
    Mid(xLinha, 5 + 5 - i, i) = Format(x_codigo, "#,000")
    Mid(xLinha, 17, 40) = x_nome
    Mid(xLinha, vbInformation, 3) = x_unidade
    i = Len(Format(x_quantidade, "###,###,##0"))
    Mid(xLinha, 74 + 11 - i, i) = Format(x_quantidade, "###,###,##0")
    
        
    
    lQtdCupom = MovCupomFiscal.QuantidadeProdutoVendaData(g_empresa, x_codigo, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), 1, 4, 0)
    lQtdLubrif = MovLubrificante.TotalQtd(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), x_codigo)
        
    If lQtdCupom > 0 Or lQtdLubrif > 0 Then
        
        i = Len(Format(lQtdCupom, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(lQtdCupom, "###,###,##0.00")
        i = Len(Format(lQtdLubrif, "###,###,##0.00"))
        Mid(xLinha, 113 + 14 - i, i) = Format(lQtdLubrif, "###,###,##0.00")
        lQtdLubrif = lQtdCupom - lQtdLubrif
        If lQtdLubrif = 0 Then
            Exit Sub
        End If
        i = Len(Format(lQtdLubrif, "##0"))
        Mid(xLinha, 132 + 3 - i, i) = Format(lQtdLubrif, "##0")
        
        
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpGrupo()
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|           |                                              |         |                 |                    |                    |      |"
    If Grupo.LocalizarCodigo(lCodigoGrupo) Then
        Mid(xLinha, 15, 40) = "GRUPO: " & Grupo.Nome
    Else
        Mid(xLinha, 15, 40) = "GRUPO: " & "** Grupo Não Cadastrado **"
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 2
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                 TOTAL DO RELATÓRIO |                 |                    |                    |      |"
    i = Len(Format(lQuantidade, "###,###,###"))
    Mid(xLinha, 74 + 11 - i, i) = Format(lQuantidade, "###,###,###")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print xLinha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontBold = False"
    xLinha = "+--------------------------------------------------------------------+-----------------+--------------------+--------------------+------+"
    Mid(xLinha, 3, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                                                                           Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    '                        1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '               12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '                                                                                                            123456789012345678901234567890
    xLinha = "| COMPARA VENDA LUBRIFICANTE X CUPOM FISCAL                GRUPO.:                                                   CIDADE, __/__/____ |"
    Mid(xLinha, 68, 30) = cbo_grupo.Text
    i = Len(g_cidade_empresa)
    Mid(xLinha, 94 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| ESTOQUE.:                                                                                                                             |"
    Mid(xLinha, 13, 30) = cboTipoEstoque.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  CODIGO   |                                              |         |                 |           V  A  L  O  R  E  S           |      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|    DO     |   DISCRIMINAÇÃO DOS PRODUTOS                 | UNIDADE |    QUANTIDADE   +--------------------+--------------------+      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  PRODUTO  |                                              |         |    EM ESTOQUE   |   VENDA DO CUPOM   |   QUANTIDADE DA    |      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|           |                                              |         |                 |      FISCAL        |   VENDA  INTERNA   |      |"
    BioImprime "@Printer.Print " & xLinha
    lCodigoGrupo = 0
'    If chk_linha_separadora.Value = 0 Then
'        xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
'        BioImprime "@Printer.Print " & xLinha
'    End If
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_produto.SetFocus
    End If
End Sub
Private Sub cbo_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoEstoque.SetFocus
    End If
End Sub
Private Sub cboTipoEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chk_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub chkExclusivoLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
Private Sub chkExclusivoPosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoLoja.SetFocus
    End If
End Sub
Private Sub chkImprimeEstoqueZerado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    chkExclusivoPosto.SetFocus
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
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Escolha o grupo.", vbInformation, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf cbo_produto.ListIndex = -1 Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        cbo_produto.SetFocus
    ElseIf cboTipoEstoque.ListIndex = -1 Then
        MsgBox "Escolha o tipo de estoque.", vbInformation, "Atenção!"
        cboTipoEstoque.SetFocus
    ElseIf chkExclusivoPosto.Value = 0 And chkExclusivoLoja.Value = 0 Then
        MsgBox "Selecione produto a imprimir.", vbInformation, "Atenção!"
        chkExclusivoPosto.SetFocus
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
        txtDataInicial.Text = Format(g_data_def, "dd/mm/yyyy")
        txtDataFinal.Text = Format(g_data_def, "dd/mm/yyyy")
        chkExclusivoPosto.Value = 1
        chkExclusivoLoja.Value = 0
        If UCase(g_nome_empresa) Like "*LOJA*" Then
            chkExclusivoPosto.Value = 0
            chkExclusivoLoja.Value = 1
        End If
        cbo_grupo.ListIndex = 0
        cbo_produto.ListIndex = 0
        txtDataInicial.SetFocus
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
    
    PreencheCboGrupo
    PreencheCboProduto
    PreencheCboTipoEstoque
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoPosto.SetFocus
    End If
End Sub
Private Sub txtDataFinal_GotFocus()
    txtDataFinal.SelStart = 0
    txtDataFinal.SelLength = 5
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoPosto.SetFocus
    End If
End Sub
Private Sub txtDataInicial_GotFocus()
    txtDataInicial.SelStart = 0
    txtDataInicial.SelLength = 5
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataFinal.SetFocus
    End If
End Sub
