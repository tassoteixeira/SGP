VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_entrada_produto 
   Caption         =   "Emissão das Entradas de Produtos"
   ClientHeight    =   3915
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_entrada_produto.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_entrada_produto.frx":030A
   ScaleHeight     =   3915
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_entrada_produto.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza entrada de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_entrada_produto.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime entrada de produtos."
      Top             =   2940
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_entrada_produto.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2940
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_entrada_produto.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_entrada_produto.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_entrada_produto.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_entrada 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2310
         Width           =   2175
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         ItemData        =   "lst_entrada_produto.frx":7F94
         Left            =   1680
         List            =   "lst_entrada_produto.frx":7F96
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   4755
      End
      Begin VB.ComboBox cbo_produto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   4755
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
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
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
      Begin VB.Label Label3 
         Caption         =   "&Tipo da entrada"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Produto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
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
Attribute VB_Name = "lst_entrada_produto"
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

Dim lSQL As String
Private rsTabela As New adodb.Recordset

Private Grupo As New cGrupo
Private Produto As New cProduto

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
    Set Grupo = Nothing
    Set Produto = Nothing
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
Private Sub PreencheCboTipoEntrada()
    cbo_tipo_entrada.Clear
    cbo_tipo_entrada.AddItem "Todos os Tipos"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 0
    cbo_tipo_entrada.AddItem "Normal"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 1
    cbo_tipo_entrada.AddItem "Acerto de Estoque"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 2
    cbo_tipo_entrada.AddItem "Inventário"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 3
    cbo_tipo_entrada.AddItem "Transferência"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 4
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lQuantidade = 0
    lTotal = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    lSQL = ""
    lSQL = lSQL & "SELECT [Data da Entrada], [Numero do Documento], [Codigo do Produto], Quantidade, [Preco de Custo], Observacao"
    lSQL = lSQL & "  FROM Entrada_Produto"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data da Entrada] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data da Entrada] <= " & preparaData(msk_data_f.Text)
    'If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
    '    lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    'End If
    If cbo_produto.ItemData(cbo_produto.ListIndex) > 0 Then
        lSQL = lSQL & "   AND [Codigo do Produto] = " & cbo_produto.ItemData(cbo_produto.ListIndex)
    End If
    If cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex) > 0 Then
        lSQL = lSQL & "   AND [Tipo da Entrada] = " & cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex)
    End If
    lSQL = lSQL & " ORDER BY [Codigo do Produto]"
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
    LoopTabelaEntrada
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Entrada de Produtos|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabelaEntrada()
    'loop entrada produto
    Dim x_linha As String
    Dim x_quantidade As Currency
    Dim x_valor1 As Currency
    Dim x_valor2 As Currency
    Do Until rsTabela.EOF
        'Le tabela auxiliar "Produto"
        If Produto.LocalizarCodigo(rsTabela("Codigo do Produto").Value) Then
            'Le tabela auxiliar "Grupo"
            If Grupo.LocalizarCodigo(Produto.CodigoGrupo) Then
                If (cbo_grupo.ItemData(cbo_grupo.ListIndex) = 0 Or cbo_grupo.ItemData(cbo_grupo.ListIndex) = Produto.CodigoGrupo) Then
                    If lPagina = 0 Then
                        ImpCab
                    End If
                    If lLinha >= 60 Then
                        x_linha = "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+------------------------+"
                        Mid(x_linha, 34, 22) = " Cerrado Informática. "
                        BioImprime "@Printer.Print " & x_linha
                        BioImprime "@@Printer.NewPage"
                        ImpCab
                    End If
                    Call ImpDet(rsTabela("Data da Entrada").Value, rsTabela("Numero do Documento").Value, rsTabela("Codigo do Produto").Value, Produto.Nome, Produto.unidade, rsTabela("Quantidade").Value, rsTabela("Preco de Custo").Value, rsTabela("Quantidade").Value * rsTabela("Preco de Custo").Value, rsTabela("Observacao").Value)
                    lTotal = lTotal + rsTabela("Quantidade").Value * rsTabela("Preco de Custo").Value
                    lQuantidade = lQuantidade + rsTabela("Quantidade").Value
                End If
            Else
                MsgBox "Grupo inexistente!" & Chr(10) & Produto.CodigoGrupo, vbInformation, "Erro de integridade!"
            End If
        Else
            MsgBox "Produto inexistente!" & Chr(10) & rsTabela("Codigo do Produto").Value, vbInformation, "Erro de integridade!"
        End If
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpDet(x_data As Date, x_documento As String, x_codigo As Long, x_nome As String, x_unidade As String, x_quantidade As Currency, x_valor1 As Currency, x_valor2 As Currency, x_observacao As String)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|          |          |      |                                       |   |         |             |             |                        |"
    Mid(x_linha, 2, 10) = Format(x_data, "dd/mm/yyyy")
    Mid(x_linha, 13, 10) = x_documento
    i = Len(Format(x_codigo, "##000"))
    Mid(x_linha, 24 + 5 - i, i) = Format(x_codigo, "##000")
    Mid(x_linha, 31, 39) = x_nome
    Mid(x_linha, 71, 3) = x_unidade
    i = Len(Format(x_quantidade, "##,##0.00"))
    Mid(x_linha, 75 + 9 - i, i) = Format(x_quantidade, "##,##0.00")
    i = Len(Format(x_valor1, "#####,##0.00"))
    Mid(x_linha, 85 + 12 - i, i) = Format(x_valor1, "#####,##0.00")
    i = Len(Format(x_valor2, "#####,##0.00"))
    Mid(x_linha, 99 + 12 - i, i) = Format(x_valor2, "#####,##0.00")
    Mid(x_linha, 113, 24) = x_observacao
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    x_linha = "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                     TOTAL DO RELATÓRIO |         |             |             |                        |"
    i = Len(Format(lQuantidade, "##,##0.00"))
    Mid(x_linha, 75 + 9 - i, i) = Format(lQuantidade, "##,##0.00")
    i = Len(Format(lTotal, "#####,##0.00"))
    Mid(x_linha, 99 + 12 - i, i) = Format(lTotal, "#####,##0.00")
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
    x_linha = "+------------------------------------------------------------------------+---------+-------------+-------------+------------------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
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
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DAS ENTRADAS DE PRODUTOS                                , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PERÍODO DE MOVIMENTAÇÃO.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| GRUPO...................:                                                    |"
    Mid(x_linha, 29, 30) = cbo_grupo
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| TIPO DE ENTRADA.........:                                                    |"
    Mid(x_linha, 29, 30) = cbo_tipo_entrada
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+------------------------+"
    BioImprime "@Printer.Print " & "|  D A T A |DOCUMENTO |CODIGO|DISCRIMINACAO DOS PRODUTOS             |UN.|  QUANT. |VLR. UNITARIO| TOTAL CUSTO |OBSERVACAO              |"
    BioImprime "@Printer.Print " & "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+------------------------+"
End Sub
Private Sub cbo_grupo_GotFocus()
    SendMessageLong cbo_grupo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_produto.SetFocus
    End If
End Sub
Private Sub cbo_produto_GotFocus()
    SendMessageLong cbo_produto.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_entrada.SetFocus
    End If
End Sub
Private Sub cbo_tipo_entrada_GotFocus()
    SendMessageLong cbo_tipo_entrada.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_grupo.SetFocus
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
    cbo_grupo.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_grupo.SetFocus
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
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Escolha o grupo.", 64, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf cbo_produto.ListIndex = -1 Then
        MsgBox "Escolha o produto.", 64, "Atenção!"
        cbo_produto.SetFocus
    ElseIf cbo_tipo_entrada.ListIndex = -1 Then
        MsgBox "Escolha o tipo de entrada.", 64, "Atenção!"
        cbo_tipo_entrada.SetFocus
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
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_grupo.ListIndex = 0
        cbo_produto.ListIndex = 0
        cbo_tipo_entrada.ListIndex = 0
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
    PreencheCboGrupo
    PreencheCboProduto
    PreencheCboTipoEntrada
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
        cbo_grupo.SetFocus
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
