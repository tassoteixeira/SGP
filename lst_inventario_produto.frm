VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_inventario_produto 
   Caption         =   "Emissão do Inventário de Produtos"
   ClientHeight    =   4740
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_inventario_produto.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_inventario_produto.frx":030A
   ScaleHeight     =   4740
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_inventario_produto.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   3780
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_inventario_produto.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   3780
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_inventario_produto.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3780
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox cboLocalizacao 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2340
         Width           =   4755
      End
      Begin VB.ComboBox cboTipoEstoque 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1920
         Width           =   4755
      End
      Begin VB.CheckBox chk_linha_separadora 
         Caption         =   "Imprime linha separadora"
         Height          =   255
         Left            =   3660
         TabIndex        =   17
         Top             =   2760
         Width           =   2235
      End
      Begin VB.CheckBox chkImprimeEstoqueZerado 
         Caption         =   "Imprime produtos sem estoque"
         Height          =   255
         Left            =   3660
         TabIndex        =   20
         Top             =   3180
         Width           =   2715
      End
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo da Loja"
         Height          =   255
         Left            =   3660
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo do Posto"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   1875
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_inventario_produto.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chk_quantidade 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   3180
         Width           =   195
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   4755
      End
      Begin VB.ComboBox cbo_produto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1500
         Width           =   4755
      End
      Begin VB.ComboBox cbo_preco 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
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
      Begin VB.Label Label3 
         Caption         =   "&Localização"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo de Estoque"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "I&mprimir Produto"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Imprimir &Quantidade"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   3180
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Produto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo de Preço"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2760
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
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_inventario_produto"
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
Dim lQtdMaximaLinha As Integer
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
Private MovLocalizacaoProduto As New cMovLocalizacaoProduto

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
    Set MovLocalizacaoProduto = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
End Sub
Private Sub PreencheCboPreco()
    cbo_preco.Clear
    cbo_preco.AddItem "Custo"
    cbo_preco.ItemData(cbo_preco.NewIndex) = 1
    cbo_preco.AddItem "Venda"
    cbo_preco.ItemData(cbo_preco.NewIndex) = 2
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
Private Sub PreencheCboLocalizacao()
    cboLocalizacao.Clear
    cboLocalizacao.AddItem "Todas as Localizações"
    cboLocalizacao.ItemData(cboLocalizacao.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM LocalizacaoProduto"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cboLocalizacao.AddItem rsTabela("Nome").Value
            cboLocalizacao.ItemData(cboLocalizacao.NewIndex) = rsTabela("Codigo").Value
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
    If g_impressora_matricial Then
        lQtdMaximaLinha = 60
    Else
        lQtdMaximaLinha = 92
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Estoque.Quantidade, Produto.[Preco de Custo], Estoque.[Preco de Venda], Produto.Unidade, Produto.[Codigo da Aliquota], Produto.[Codigo NCM]"
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
    If chkImprimeEstoqueZerado.Value = 0 Then
        lSQL = lSQL & "   AND Estoque.Quantidade > 0 "
    End If
    lSQL = lSQL & " ORDER BY [Codigo do Grupo], Produto.Nome, Produto.[Codigo NCM]"
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
    Dim xContinua As Boolean
    
    Do Until rsTabela.EOF
    
        xContinua = False
        If cboLocalizacao.ItemData(cboLocalizacao.ListIndex) = 0 Then
            xContinua = True
        Else
            If MovLocalizacaoProduto.LocalizarCodigo(g_empresa, cboLocalizacao.ItemData(cboLocalizacao.ListIndex), rsTabela("Codigo").Value) Then
                xContinua = True
            End If
        End If
        
        If xContinua = True Then
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= lQtdMaximaLinha Then
                 x_linha = "+-----------+----------------------------------------------+----------+-----+----------+--------------------+--------------------+------+"
                'x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
                Mid(x_linha, 15, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            If cbo_preco = "Custo" Then
                x_valor1 = rsTabela("Preco de Custo").Value
            Else
                x_valor1 = rsTabela("Preco de Venda").Value
            End If
            
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
            If xEstoque > 0 Or chkImprimeEstoqueZerado.Value = 1 Then
                Call ImpDet(rsTabela("Codigo").Value, rsTabela("Nome").Value, rsTabela("unidade").Value, xEstoque, x_valor1, x_valor2, xAliquota, rsTabela("Codigo NCM").Value)
            End If
            If xEstoque > 0 Then
                lTotal = lTotal + x_valor2
                lQuantidade = lQuantidade + x_quantidade
            End If
        End If
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpDet(x_codigo As Long, x_nome As String, x_unidade As String, x_quantidade As Currency, x_valor1 As Currency, x_valor2 As Currency, xAliquotaImposto As Currency, pCodigoNCM As String)
    Dim xLinha As String
    Dim i As Integer
    If chk_linha_separadora.Value = 1 Then
        xLinha = "+-----------+----------------------------------------------+----------+-----+----------+--------------------+--------------------+------+"
        'xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    xLinha = "|           |                                              |          |     |          |                    |                    |      |"
    'xLinha = "|           |                                              |         |                 |                    |                    |      |"
    i = Len(Format(x_codigo, "#,000"))
    Mid(xLinha, 5 + 5 - i, i) = Format(x_codigo, "#,000")
    Mid(xLinha, 17, 40) = x_nome
    Mid(xLinha, 63, 8) = pCodigoNCM
    Mid(xLinha, 74, 3) = x_unidade
    'Mid(xLinha, vbInformation, 3) = x_unidade
    If chk_quantidade.Value = 0 Then
        x_quantidade = 0
        x_valor2 = 0
    End If
    i = Len(Format(x_quantidade, "###,###,##0"))
    Mid(xLinha, 77 + 10 - i, i) = Format(x_quantidade, "###,###,##0")
    'Mid(xLinha, 74 + 11 - i, i) = Format(x_quantidade, "###,###,##0")
    i = Len(Format(x_valor1, "###,###,##0.00"))
    Mid(xLinha, 92 + 14 - i, i) = Format(x_valor1, "###,###,##0.00")
    i = Len(Format(x_valor2, "###,###,##0.00"))
    Mid(xLinha, 113 + 14 - i, i) = Format(x_valor2, "###,###,##0.00")
    i = Len(Format(xAliquotaImposto, "##0"))
    Mid(xLinha, 132 + 3 - i, i) = Format(xAliquotaImposto, "##0")
    Mid(xLinha, 135, 1) = "%"
    If UCase(Aliquota.Nome) Like "*ISEN*" Then
        Mid(xLinha, 131, 6) = "ISENTA"
    End If
    If UCase(Aliquota.Nome) Like "*SUBST*" Then
        Mid(xLinha, 131, 6) = "SUB.TR"
    End If
    If UCase(Aliquota.Nome) Like "*INCID*" Then
        Mid(xLinha, 131, 6) = "N.INC."
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpGrupo()
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "+-----------+----------------------------------------------+----------+-----+----------+--------------------+--------------------+------+"
    'xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|           |                                              |          |     |          |                    |                    |      |"
    'xLinha = "|           |                                              |         |                 |                    |                    |      |"
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
    
    xLinha = "+-----------+----------------------------------------------+----------+-----+----------+--------------------+--------------------+------+"
    'xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                 TOTAL DO RELATÓRIO  |                |                    |                    |      |"
    'xLinha = "|                                                 TOTAL DO RELATÓRIO |                 |                    |                    |      |"
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
    xLinha = "+---------------------------------------------------------------------+----------------+--------------------+--------------------+------+"
   'xLinha = "+--------------------------------------------------------------------+-----------------+--------------------+--------------------+------+"
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
    xLinha = "| INVENTÁRIO DE PRODUTOS                                   GRUPO.:                                                   CIDADE, __/__/____ |"
    Mid(xLinha, 68, 30) = cbo_grupo.Text
    i = Len(g_cidade_empresa)
    Mid(xLinha, 94 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| ESTOQUE.:                                           LOCALIZAÇÃO:                                                                      |"
    Mid(xLinha, 13, 30) = cboTipoEstoque.Text
    Mid(xLinha, 68, 30) = cboLocalizacao.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    'xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
    xLinha = "+-----------+----------------------------------------------+----------+-----+----------+-----------------------------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  CODIGO   |                                              |          |     |          |           V  A  L  O  R  E  S           |      |"
    'xLinha = "|  CODIGO   |                                              |         |                 |           V  A  L  O  R  E  S           |      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|    DO     |   DISCRIMINAÇÃO DOS PRODUTOS                 |   NCM    |UNID.|  ESTOQUE +--------------------+--------------------+      |"
    'xLinha = "|    DO     |   DISCRIMINAÇÃO DOS PRODUTOS                 | UNIDADE |    QUANTIDADE   +--------------------+--------------------+      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  PRODUTO  |                                              |          |     |          |   PREÇO DE _____   |   PREÇO DE _____   |      |"
    'xLinha = "|  PRODUTO  |                                              |         |    EM ESTOQUE   |   PREÇO DE _____   |   PREÇO DE _____   |      |"
    Mid(xLinha, 101, 5) = UCase(cbo_preco.Text)
    Mid(xLinha, 122, 5) = UCase(cbo_preco.Text)
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|           |                                              |          |     |          |      UNITÁRIO      |       TOTAL        |      |"
    'xLinha = "|           |                                              |         |                 |      UNITÁRIO      |       TOTAL        |      |"
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
Private Sub cbo_preco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_quantidade.SetFocus
    End If
End Sub
Private Sub cbo_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoEstoque.SetFocus
    End If
End Sub
Private Sub cboLocalizacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_preco.SetFocus
    End If
End Sub
Private Sub cboTipoEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboLocalizacao.SetFocus
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
    ElseIf cbo_preco.ListIndex = -1 Then
        MsgBox "Escolha o tipo de preço.", vbInformation, "Atenção!"
        cbo_preco.SetFocus
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
        chkExclusivoPosto.Value = 1
        chkExclusivoLoja.Value = 0
        If UCase(g_nome_empresa) Like "*LOJA*" Then
            chkExclusivoPosto.Value = 0
            chkExclusivoLoja.Value = 1
        End If
        cbo_grupo.ListIndex = 0
        cbo_produto.ListIndex = 0
        cboTipoEstoque.ListIndex = 0
        cboLocalizacao.ListIndex = 0
        cbo_preco.ListIndex = 1
        chk_quantidade.Value = 1
        cmd_imprimir.SetFocus
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
    PreencheCboLocalizacao
    PreencheCboPreco
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
