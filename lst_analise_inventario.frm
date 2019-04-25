VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_analise_inventario 
   Caption         =   "Emissão da Análise do Inventário"
   ClientHeight    =   2505
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "lst_analise_inventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2505
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "lst_analise_inventario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "lst_analise_inventario.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprime análise do inventário."
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_analise_inventario.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Visualiza análise do inventário."
      Top             =   1560
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.CheckBox chkImprimeEstoqueZerado 
         Caption         =   "Imprime produtos sem estoque"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   2715
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_inventario.frx":46C0
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
         TabIndex        =   5
         Top             =   660
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
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   660
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
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_inventario"
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
Private rsTabela As New adodb.Recordset

Private Grupo As New cGrupo
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Grupo = Nothing
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
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Estoque.Quantidade, Produto.[Preco de Custo], Estoque.[Preco de Venda], Produto.Unidade"
    lSQL = lSQL & "  FROM Produto, Estoque"
    lSQL = lSQL & " WHERE Estoque.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Estoque.[Codigo do Produto2] = Produto.Codigo"
    lSQL = lSQL & "   AND Produto.Inativo = " & preparaBooleano(False)
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
'    If chkExclusivoPosto.Value = 1 And chkExclusivoLoja.Value = 0 Then
'        lSQL = lSQL & "   AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
'    End If
'    If chkExclusivoLoja.Value = 1 And chkExclusivoPosto.Value = 0 Then
'        lSQL = lSQL & "   AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
'    End If
'    If chkImprimeEstoqueZerado.Value = 0 Then
'        lSQL = lSQL & "   AND Estoque.Quantidade > 0 "
'    End If
    lSQL = lSQL & " ORDER BY [Codigo do Grupo], Produto.Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        LoopTabelaProduto
        If lPagina > 0 Then
            ImpTotal
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Análise do Inventário|@|"
            frm_preview.Show 1
        End If
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela produto
    Dim x_linha As String
    Dim x_total_custo As Currency
    Dim x_total_venda As Currency
    
    rsTabela.MoveFirst
    Do Until rsTabela.EOF
        If rsTabela("Quantidade").Value > 0 Then
            l_total_custo = l_total_custo + Format(rsTabela("Preco de Custo").Value * rsTabela("Quantidade").Value, "0000000000.00")
            l_total_venda = l_total_venda + Format(rsTabela("Preco de Venda").Value * rsTabela("Quantidade").Value, "0000000000.00")
            l_total_quantidade = l_total_quantidade + rsTabela("Quantidade").Value
        ElseIf chkImprimeEstoqueZerado.Value = 1 Then
        End If
        rsTabela.MoveNext
    Loop
    rsTabela.MoveFirst
    Do Until rsTabela.EOF
        If rsTabela("Quantidade").Value > 0 Or chkImprimeEstoqueZerado.Value = 1 Then
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 60 Then
                x_linha = "+------+-------------------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+"
                Mid(x_linha, 12, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            x_total_custo = Format(rsTabela("Preco de Custo").Value * rsTabela("Quantidade").Value, "0000000000.00")
            x_total_venda = Format(rsTabela("Preco de Venda").Value * rsTabela("Quantidade").Value, "0000000000.00")
            Call ImpDet(rsTabela("Codigo").Value, rsTabela("Nome").Value, rsTabela("Unidade").Value, rsTabela("Quantidade").Value, rsTabela("Preco de Custo").Value, x_total_custo, rsTabela("Preco de Venda").Value, x_total_venda)
        End If
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim x_percentual As Currency
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------+-------------------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+"
    x_linha = "|                                            *** TOTAL |        |       |           |      %|      %|       |           |      %|      %|"
    i = Len(Format(l_total_quantidade, "####,###"))
    Mid(x_linha, 57 + 8 - i, i) = Format(l_total_quantidade, "####,###")
    i = Len(Format(l_total_custo, "####,##0.00"))
    Mid(x_linha, 74 + 11 - i, i) = Format(l_total_custo, "####,##0.00")
    If l_total_custo > 0 Then
        x_percentual = (l_total_venda - l_total_custo) * 100 / l_total_custo
    Else
        x_percentual = 0
    End If
    i = Len(Format(x_percentual, "###.##"))
    Mid(x_linha, 86 + 6 - i, i) = Format(x_percentual, "###.##")
    Mid(x_linha, 94, 6) = "100,00"
    i = Len(Format(l_total_venda, "####,##0.00"))
    Mid(x_linha, 110 + 11 - i, i) = Format(l_total_venda, "####,##0.00")
    If l_total_venda > 0 Then
        x_percentual = (l_total_venda - l_total_custo) * 100 / l_total_venda
    Else
        x_percentual = 0
    End If
    i = Len(Format(x_percentual, "###.##"))
    Mid(x_linha, 122 + 6 - i, i) = Format(x_percentual, "###.##")
    Mid(x_linha, 130, 6) = "100,00"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+------------------------------------------------------+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(pCodigo As Long, pNome As String, pUnidade As String, pQuantidade As Currency, pCusto As Currency, pTotalCcusto As Currency, pVenda As Currency, pTotalVenda As Currency)
    Dim x_linha As String
    Dim x_percentual As Currency
    Dim i As Integer
    x_linha = "|      |                                           |   |        |       |           |      %|       |       |           |      %|       |"
    i = Len(Format(pCodigo, "##,000"))
    Mid(x_linha, 2 + 6 - i, i) = Format(pCodigo, "##,000")
    Mid(x_linha, 10, 40) = pNome
    Mid(x_linha, 53, 3) = pUnidade
    i = Len(Format(pQuantidade, "####,###"))
    Mid(x_linha, 57 + 8 - i, i) = Format(pQuantidade, "####,###")
    i = Len(Format(pCusto, "###0.00"))
    Mid(x_linha, 66 + 7 - i, i) = Format(pCusto, "###0.00")
    If pTotalCcusto <> 0 Then
        i = Len(Format(pTotalCcusto, "####,##0.00"))
        Mid(x_linha, 74 + 11 - i, i) = Format(pTotalCcusto, "####,##0.00")
    End If
    If pCusto > 0 Then
        x_percentual = (pVenda - pCusto) * 100 / pCusto
    Else
        x_percentual = 100
    End If
    i = Len(Format(x_percentual, "##0.00"))
    Mid(x_linha, 86 + 6 - i, i) = Format(x_percentual, "##0.00")
    If pTotalCcusto > 0 Then
        x_percentual = (pTotalCcusto * 100 / l_total_custo)
    Else
        x_percentual = 0
    End If
    If x_percentual > 0 Then
        i = Len(Format(x_percentual, "##0.00"))
        Mid(x_linha, 94 + 6 - i, i) = Format(x_percentual, "##0.00")
        Mid(x_linha, 100, 1) = "%"
    End If
    i = Len(Format(pVenda, "###0.00"))
    Mid(x_linha, 102 + 7 - i, i) = Format(pVenda, "###0.00")
    If pTotalVenda <> 0 Then
        i = Len(Format(pTotalVenda, "####,##0.00"))
        Mid(x_linha, 110 + 11 - i, i) = Format(pTotalVenda, "####,##0.00")
    End If
    If pVenda > 0 Then
        x_percentual = (pVenda - pCusto) * 100 / pVenda
    Else
        x_percentual = 0
    End If
    i = Len(Format(x_percentual, "###.##"))
    Mid(x_linha, 122 + 6 - i, i) = Format(x_percentual, "###.##")
    If pTotalVenda > 0 Then
        x_percentual = (pTotalVenda * 100 / l_total_venda)
    Else
        x_percentual = 0
    End If
    If x_percentual > 0 Then
        i = Len(Format(x_percentual, "##0.00"))
        Mid(x_linha, 130 + 6 - i, i) = Format(x_percentual, "##0.00")
        Mid(x_linha, 136, 1) = "%"
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    Printer.CurrentY = 0
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| ANÁLISE DO INVENTÁRIO  FUNCIONÁRIOS                             , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| GRUPO.:                                                                      |"
    Mid(x_linha, 11, 3) = Format(cbo_grupo.ItemData(cbo_grupo.ListIndex), "000")
    Mid(x_linha, 15, 30) = cbo_grupo
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------+-------------------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+"
    BioImprime "@Printer.Print " & "|      |                                           |   | QUANT. | PREÇO |   TOTAL   |% CUSTO|% SOBRE| PREÇO |   TOTAL   |% VENDA|% SOBRE|"
    BioImprime "@Printer.Print " & "|CÓDIGO|         DISCRIMINAÇÃO DOS PRODUTOS        |UN.|   EM   |  DE   |     DO    |  PARA | TOTAL |  DE   |     DA    |  PARA | TOTAL |"
    BioImprime "@Printer.Print " & "|      |                                           |   | ESTOQUE| CUSTO |   CUSTO   | VENDA |ESTOQUE| VENDA |   VENDA   | CUSTO |ESTOQUE|"
    BioImprime "@Printer.Print " & "+------+-------------------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+"
End Sub
Private Sub cbo_grupo_GotFocus()
    SendMessageLong cbo_grupo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
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
    cbo_grupo.SetFocus
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
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", 64, "Atenção!"
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
        cbo_grupo.ListIndex = 0
        cbo_grupo.SetFocus
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
