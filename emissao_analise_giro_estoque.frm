VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_analise_giro_estoque 
   Caption         =   "Emissão da Análise do Giro de Estoque"
   ClientHeight    =   2715
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_analise_giro_estoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_analise_giro_estoque.frx":030A
   ScaleHeight     =   2715
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_analise_giro_estoque.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Visualiza a análise do giro de estoque."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_analise_giro_estoque.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime a análise do giro de estoque."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_analise_giro_estoque.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_analise_giro_estoque.frx":4706
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
         Width           =   4755
      End
      Begin VB.ComboBox cbo_tipo_venda 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
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
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Vendas"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1080
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
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_giro_estoque"
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
Dim lQuantidade(1 To 3) As Currency
Dim lTotal As Currency
Dim lData(1 To 3) As Date
Dim lDataEntrada As Date
Dim lQuantidadeEntrada As Currency
Dim tbl_entrada_produto As Table
Dim tbl_estoque As Table
Dim tbl_grupo As Table
Dim tbl_movimento_lubrificante As Table
Dim tbl_produto As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_entrada_produto.Close
    tbl_estoque.Close
    tbl_grupo.Close
    tbl_movimento_lubrificante.Close
    tbl_produto.Close
End Sub
Private Sub PreencheCboTipoVenda()
    cbo_tipo_venda.Clear
    cbo_tipo_venda.AddItem "> 0"
    cbo_tipo_venda.ItemData(cbo_tipo_venda.NewIndex) = 1
    cbo_tipo_venda.AddItem "= 0"
    cbo_tipo_venda.ItemData(cbo_tipo_venda.NewIndex) = 2
End Sub
Private Sub PreencheCboGrupo()
    cbo_grupo.Clear
    With tbl_grupo
        .Index = "id_nome"
        cbo_grupo.AddItem "Todos os Grupos"
        cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
        .MoveFirst
        Do Until .EOF
            cbo_grupo.AddItem !Nome
            cbo_grupo.ItemData(cbo_grupo.NewIndex) = !Codigo
            .MoveNext
        Loop
    End With
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
    lData(3) = CDate("01/" & Month(msk_data) & "/" & Year(msk_data))
    If Month(msk_data) = 1 Then
        lData(2) = CDate("01/12/" & Year(msk_data) - 1)
    Else
        lData(2) = CDate("01/" & Month(msk_data) - 1 & "/" & Year(msk_data))
    End If
    If Month(msk_data) = 2 Then
        lData(1) = CDate("01/12/" & Year(msk_data) - 1)
    ElseIf Month(msk_data) = 1 Then
        lData(1) = CDate("01/11/" & Year(msk_data) - 1)
    Else
        lData(1) = CDate("01/" & Month(msk_data) - 2 & "/" & Year(msk_data))
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica produto
    With tbl_produto
        .Index = "id_nome"
        .Seek ">=", " ", 0, 0
        If Not .NoMatch Then
            ImpDados
        End If
    End With
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    LoopTabelaProduto
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Análise do Giro de Estoque|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela Produto
    With tbl_produto
        Do Until .EOF
            If cbo_grupo.ItemData(cbo_grupo.ListIndex) = 0 Or cbo_grupo.ItemData(cbo_grupo.ListIndex) = ![Codigo do Grupo] Then
                If !Inativo = False Then
                    LoopTabelaMovimentoLubrificante
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoopTabelaEntradaProduto()
'+Empresa;+Codigo do Produto2;+Data da Entrada;+Numero do Documento
    lDataEntrada = Null
    lQuantidadeEntrada = 0
    With tbl_entrada_produto
        .Seek "<=", g_empresa, tbl_produto!Codigo, CDate(msk_data), "9999999999"
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Codigo do Produto] <> tbl_produto!Codigo Then
                    Exit Do
                End If
                If ![Tipo de Entrada] = "1" Then
                    lDataEntrada = ![Data da Entrada]
                    lQuantidadeEntrada = !Quantidade
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopTabelaMovimentoLubrificante()
    Dim i As Integer
    'loop tabela Movimento_Lubrificante
    For i = 1 To 3
        lQuantidade(i) = 0
    Next
    With tbl_movimento_lubrificante
        .Seek ">=", g_empresa, tbl_produto!Codigo, lData(1), "0", 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Codigo do Produto2] <> tbl_produto!Codigo Then
                    Exit Do
                End If
                If Month(!Data) > Month(lData(3)) And Year(!Data) > Year(lData(3)) Then
                    Exit Do
                End If
                For i = 1 To 3
                    If Month(lData(i)) = Month(!Data) Then
                        lQuantidade(i) = lQuantidade(i) + !Quantidade
                    End If
                Next
                .MoveNext
            Loop
        End If
    End With
    'If (lQuantidade(1) + lQuantidade(2) + lQuantidade(3)) > 0 Then
        Call ImpDet
    'End If
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    Dim x_quantidade As Currency
    Dim x_media As Currency
    If lPagina = 0 Then
        ImpCab
    End If
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "|   |                                              |         |         |         |         |         |         |                        |"
    If lLinha >= 60 Then
        x_linha = "+---+----------------------------------------------+---------+---------+---------+---------+---------+---------+------------------------+"
        Mid(x_linha, 17, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    i = Len(Format(tbl_produto!Codigo, "000"))
    Mid(x_linha, 1 + 3 - i, i) = Format(tbl_produto!Codigo, "000")
    Mid(x_linha, 7, 40) = tbl_produto!Nome
    i = Len(Format(lQuantidade(1), "###,##0"))
    Mid(x_linha, 62 + 7 - i, i) = Format(lQuantidade(1), "###,##0")
    i = Len(Format(lQuantidade(2), "###,##0"))
    Mid(x_linha, 72 + 7 - i, i) = Format(lQuantidade(2), "###,##0")
    i = Len(Format(lQuantidade(3), "###,##0"))
    Mid(x_linha, 82 + 7 - i, i) = Format(lQuantidade(3), "###,##0")
    'Le tabela auxiliar
    tbl_estoque.Seek "=", g_empresa, tbl_produto!Codigo
    If Not tbl_estoque.NoMatch Then
        x_quantidade = tbl_estoque!Quantidade
    Else
        x_quantidade = 0
    End If
    x_media = lQuantidade(1) + lQuantidade(2) + lQuantidade(3)
    i = Len(Format(x_media, "###,##0"))
    Mid(x_linha, 92 + 7 - i, i) = Format(x_media, "###,##0")
    x_media = x_media / 3
    i = Len(Format(x_media, "###,##0"))
    Mid(x_linha, 102 + 7 - i, i) = Format(x_media, "###,##0")
    i = Len(Format(x_quantidade, "###,##0"))
    Mid(x_linha, 112 + 7 - i, i) = Format(x_quantidade, "###,##0")
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    x_linha = "+-----------+----------------------------------------------+---------+---------+---------+---------+---------+---------+----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|           |                                    *** TOTAL |         |         |         |         |         |         |                |"
    'i = Len(Format(lQuantidade, "###,###,###"))
    'Mid(x_linha, 74 + 11 - i, i) = Format(lQuantidade, "###,###,###")
    'i = Len(Format(lTotal, "###,###,##0.00"))
    'Mid(x_linha, 113 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
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
    x_linha = "+---+----------------------------------------------+---------+---------+---------+---------+---------+---------+------------------------+"
    Mid(x_linha, 17, 22) = " Cerrado Informática. "
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
    x_linha = "| ANALISE DE GIRO DE ESTOQUE                               GRUPO.:                                                  Goiânia, __/__/____ |"
    Mid(x_linha, 126, 10) = msk_data
    Mid(x_linha, 68, 30) = cbo_grupo
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "+---+------------------------------------------+----------+-------+----+---------+---------+---------+---------+------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|COD| DESCRIMINAÇÃO DOS PRODUTOS               |99/99/9999|123,123|DAS |  SAIDAS |  TOTAL  |  MEDIA  | ESTOQUE |                        |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|   | 1234567890123456789012345678901234567890 |          |  | __/____ | __/____ |  SAIDAS |  SAIDAS |  ATUAL  |                        |"
    Mid(x_linha, 62, 7) = Format(lData(1), "mm") & "/" & Format(lData(1), "yyyy")
    Mid(x_linha, 72, 7) = Format(lData(2), "mm") & "/" & Format(lData(2), "yyyy")
    Mid(x_linha, 82, 7) = Format(lData(3), "mm") & "/" & Format(lData(3), "yyyy")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+---+------------------------------------------+----------+---------+---------+---------+---------+---------+------------------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub cbo_grupo_GotFocus()
    SendMessageLong cbo_grupo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_venda.SetFocus
    End If
End Sub
Private Sub cbo_tipo_venda_GotFocus()
    SendMessageLong cbo_tipo_venda.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_venda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
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
        MsgBox "Escolha o grupo.", 64, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf cbo_tipo_venda.ListIndex = -1 Then
        MsgBox "Escolha o tipo de venda.", 64, "Atenção!"
        cbo_tipo_venda.SetFocus
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
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_grupo.ListIndex = 0
        cbo_tipo_venda.ListIndex = 1
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
    Set tbl_entrada_produto = bd_sgp.OpenTable("Entrada_Produto")
    Set tbl_estoque = bd_sgp.OpenTable("Estoque")
    Set tbl_grupo = bd_sgp.OpenTable("Grupo")
    Set tbl_movimento_lubrificante = bd_sgp.OpenTable("Movimento_Lubrificante")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    tbl_entrada_produto.Index = "id_produto"
    tbl_estoque.Index = "id_codigo2"
    tbl_movimento_lubrificante.Index = "id_produto"
    PreencheCboGrupo
    PreencheCboTipoVenda
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
