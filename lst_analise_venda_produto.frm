VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_analise_venda_produto 
   Caption         =   "Emiss�o da An�lise da Venda de Produtos"
   ClientHeight    =   3075
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6915
   Icon            =   "lst_analise_venda_produto.frx":0000
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
      Picture         =   "lst_analise_venda_produto.frx":030A
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
      Picture         =   "lst_analise_venda_produto.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime an�lise da venda de produtos."
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_analise_venda_produto.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza an�lise da venda de produtos."
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
         Picture         =   "lst_analise_venda_produto.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calend�rio."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_analise_venda_produto.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calend�rio."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6060
         Picture         =   "lst_analise_venda_produto.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calend�rio."
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
         Caption         =   "Per�odo &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "&Per�odo inicial"
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
         Caption         =   "Data de &emiss�o"
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
Attribute VB_Name = "emissao_analise_venda_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'In�cio de vari�veis padr�o para relat�rio
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de vari�veis padr�o para relat�rio
Dim l_total_quantidade As Currency
Dim l_total_custo As Currency
Dim l_total_venda As Currency
Dim lSQL As String

Dim rstTabela As adodb.Recordset
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relat�rio!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
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
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Produto.Unidade,"
    lSQL = lSQL & "       SUM(Movimento_Lubrificante.Quantidade) AS Quantidade,"
    lSQL = lSQL & "       SUM(Movimento_Lubrificante.Quantidade * Movimento_Lubrificante.[Valor Custo]) AS TotalCusto,"
    lSQL = lSQL & "       SUM(Movimento_Lubrificante.[Valor Total]) AS TotalVenda,"
    lSQL = lSQL & "       Max(Movimento_Lubrificante.[Valor Custo]) As PrecoCusto,"
    lSQL = lSQL & "       Max(Movimento_Lubrificante.[Valor Venda]) As PrecoVenda"
    lSQL = lSQL & "  FROM Produto, Movimento_Lubrificante"
    lSQL = lSQL & " WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Movimento_Lubrificante.[Codigo do Produto2] = Produto.Codigo"
    lSQL = lSQL & "   AND Movimento_Lubrificante.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Movimento_Lubrificante.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Movimento_Lubrificante.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Movimento_Lubrificante.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
    lSQL = lSQL & " GROUP BY Produto.Codigo, Produto.Nome, Produto.[Codigo do Grupo], Produto.[Unidade]"
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
            g_string = lLocal & lNomeArquivo & "|@|An�lise da Venda dos Produto|@|"
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
    Dim x_quantidade As Currency
    Dim x_custo As Currency
    Dim x_venda As Currency
    Dim x_total_custo As Currency
    Dim x_total_venda As Currency
    With rstTabela
        .MoveFirst
        Do Until .EOF
            l_total_quantidade = l_total_quantidade + rstTabela("Quantidade").Value
            l_total_custo = l_total_custo + rstTabela("TotalCusto").Value
            l_total_venda = l_total_venda + rstTabela("TotalVenda").Value
            .MoveNext
        Loop
        
        .MoveFirst
        Do Until .EOF
            If rstTabela("Quantidade").Value > 0 Then
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 60 Then
                    xLinha = "+------+-------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
                    Mid(xLinha, 12, 22) = " Cerrado Inform�tica. "
                    BioImprime "@Printer.Print " & xLinha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                Call ImpDet
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim x_percentual As Currency
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------+-------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
    xLinha = "|                                *** TOTAL |        |       |           |      %|100,00%|       |           |      %|100,00%|           |"
    i = Len(Format(l_total_quantidade, "####,##0"))
    Mid(xLinha, 45 + 8 - i, i) = Format(l_total_quantidade, "####,##0")
    i = Len(Format(l_total_custo, "####,##0.00"))
    Mid(xLinha, 62 + 11 - i, i) = Format(l_total_custo, "####,##0.00")
    x_percentual = (l_total_venda - l_total_custo) * 100 / l_total_custo
    i = Len(Format(x_percentual, "##0.00"))
    Mid(xLinha, 74 + 6 - i, i) = Format(x_percentual, "##0.00")
    i = Len(Format(l_total_venda, "####,##0.00"))
    Mid(xLinha, 98 + 11 - i, i) = Format(l_total_venda, "####,##0.00")
    x_percentual = (l_total_venda - l_total_custo) * 100 / l_total_venda
    i = Len(Format(x_percentual, "###.##"))
    Mid(xLinha, 110 + 6 - i, i) = Format(x_percentual, "###.##")
    x_percentual = l_total_venda - l_total_custo
    i = Len(Format(x_percentual, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(x_percentual, "####,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "+------------------------------------------+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
    Mid(xLinha, 5, 22) = " Cerrado Inform�tica. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim x_percentual As Currency
    Dim i As Integer
    xLinha = "|      |                               |   |        |       |           |      %|      %|       |           |      %|      %|           |"
    
    
', ,
    
    i = Len(Format(rstTabela("Codigo").Value, "##,000"))
    Mid(xLinha, 2 + 6 - i, i) = Format(rstTabela("Codigo").Value, "##,000")
    Mid(xLinha, 9, 31) = rstTabela("Nome").Value
    Mid(xLinha, 41, 3) = rstTabela("Unidade").Value
    i = Len(Format(rstTabela("Quantidade").Value, "####,##0"))
    Mid(xLinha, 45 + 8 - i, i) = Format(rstTabela("Quantidade").Value, "####,##0")
    i = Len(Format(rstTabela("PrecoCusto").Value, "###0.00"))
    Mid(xLinha, 54 + 7 - i, i) = Format(rstTabela("PrecoCusto").Value, "###0.00")
    i = Len(Format(rstTabela("TotalCusto").Value, "####,##0.00"))
    Mid(xLinha, 62 + 11 - i, i) = Format(rstTabela("TotalCusto").Value, "####,##0.00")
    If rstTabela("TotalCusto").Value > 0 Then
        x_percentual = (rstTabela("TotalVenda").Value - rstTabela("TotalCusto").Value) * 100 / rstTabela("TotalCusto").Value
    Else
        x_percentual = 100
    End If
    i = Len(Format(x_percentual, "##0.00"))
    Mid(xLinha, 74 + 6 - i, i) = Format(x_percentual, "##0.00")
    x_percentual = (rstTabela("TotalCusto").Value * 100 / l_total_custo)
    i = Len(Format(x_percentual, "##0.00"))
    Mid(xLinha, 82 + 6 - i, i) = Format(x_percentual, "##0.00")
    i = Len(Format(rstTabela("PrecoVenda").Value, "###0.00"))
    Mid(xLinha, 90 + 7 - i, i) = Format(rstTabela("PrecoVenda").Value, "###0.00")
    i = Len(Format(rstTabela("TotalVenda").Value, "####,##0.00"))
    Mid(xLinha, 98 + 11 - i, i) = Format(rstTabela("TotalVenda").Value, "####,##0.00")
    If rstTabela("TotalVenda").Value > 0 Then
        x_percentual = (rstTabela("TotalVenda").Value - rstTabela("TotalCusto").Value) * 100 / rstTabela("TotalVenda").Value
    Else
        x_percentual = 0
    End If
    i = Len(Format(x_percentual, "##0.00"))
    Mid(xLinha, 110 + 6 - i, i) = Format(x_percentual, "##0.00")
    x_percentual = (rstTabela("TotalVenda").Value * 100 / l_total_venda)
    i = Len(Format(x_percentual, "##0.00"))
    Mid(xLinha, 118 + 6 - i, i) = Format(x_percentual, "##0.00")
    x_percentual = rstTabela("TotalVenda").Value - rstTabela("TotalCusto").Value
    i = Len(Format(x_percentual, "####,##0.00"))
    Mid(xLinha, 126 + 11 - i, i) = Format(x_percentual, "####,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    If lPagina = 0 Then
        'seleciona medidas para cent�metros
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
    xLinha = "|                                                                  P�gina, " & Format(lPagina, "000") & " |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| AN�LISE DA VENDA DE PRODUTOS                                    , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____       PER�ODO.: _ AO _     |"
    Mid(xLinha, 29, 10) = msk_data_i.Text
    Mid(xLinha, 42, 10) = msk_data_f.Text
    Mid(xLinha, 69, 1) = cbo_periodo_i.Text
    Mid(xLinha, 74, 1) = cbo_periodo_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| GRUPO.:                                                                      |"
    Mid(xLinha, 11, 3) = Format(cbo_grupo.ItemData(cbo_grupo.ListIndex), "000")
    Mid(xLinha, 15, 30) = cbo_grupo.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------+-------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
    BioImprime "@Printer.Print " & "|      |                               |   | QUANT. | PRE�O |   TOTAL   |% CUSTO|% SOBRE| PRE�O |   TOTAL   |% VENDA|% SOBRE|   VALOR   |"
    BioImprime "@Printer.Print " & "|C�DIGO|  DISCRIMINA��O DOS PRODUTOS   |UN.|        |  DE   |     DO    |  PARA | TOTAL |  DE   |     DA    |  PARA | TOTAL |     DO    |"
    BioImprime "@Printer.Print " & "|      |                               |   |VENDIDAS| CUSTO |   CUSTO   | VENDA | CUSTO | VENDA |   VENDA   | CUSTO | VENDAS|   LUCRO   |"
    BioImprime "@Printer.Print " & "+------+-------------------------------+---+--------+-------+-----------+-------+-------+-------+-----------+-------+-------+-----------+"
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
        MsgBox "Informe a data de emiss�o.", vbInformation, "Aten��o!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Aten��o!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Aten��o!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Aten��o!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o per�odo inicial.", vbInformation, "Aten��o!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o per�odo final.", vbInformation, "Aten��o!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Aten��o!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", vbInformation, "Aten��o!"
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
