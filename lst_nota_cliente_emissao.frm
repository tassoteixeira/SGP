VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_nota_cliente_emissao 
   Caption         =   "Emissão das Notas de Abastecimento por Emissão"
   ClientHeight    =   3765
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_nota_cliente_emissao.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_nota_cliente_emissao.frx":030A
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_nota_cliente_emissao.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Visualiza notas de abastecimento por emissão."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_nota_cliente_emissao.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprime notas de abastecimento por emissão."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_nota_cliente_emissao.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkOrdemAlfabetica 
         Caption         =   "Ordem Alfabética"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_emissao.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_emissao.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_emissao.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
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
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   2175
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
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Caption         =   "&Período inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Período &final"
         Height          =   315
         Left            =   4380
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_nota_cliente_emissao"
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
Dim lSQL As String
'Fim de variáveis padrão para relatório
Dim l_cliente As Long
Dim l_conveniado As Long
Dim l_numero_nota As Long
Dim lSubTotal As Currency
Dim lSubQtd As Integer
Dim lTotal As Currency
Dim lSubTotalLiquido As Currency
Dim lTotalLiquido As Currency


Private ClienteConveniado As New cClienteConveniado
Private rsTabela As New adodb.Recordset
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
    Set ClienteConveniado = Nothing
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
    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Notas Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lSubQtd = 0
    lTotal = 0
    lSubTotalLiquido = 0
    lTotalLiquido = 0
    l_cliente = 0
    l_conveniado = 0
    l_numero_nota = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Nota_Abastecimento.[Codigo do Cliente], Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.[Numero da Nota],"
    lSQL = lSQL & "       Movimento_Nota_Abastecimento.[Codigo do Conveniado], Movimento_Nota_Abastecimento.[Tipo do Movimento], Movimento_Nota_Abastecimento.Periodo,"
    lSQL = lSQL & "       Movimento_Nota_Abastecimento.Quantidade, Movimento_Nota_Abastecimento.[Valor Total], Movimento_Nota_Abastecimento.[Codigo do Produto2],"
    lSQL = lSQL & "       Movimento_Nota_Abastecimento.[Numero do Cupom], Movimento_Nota_Abastecimento.[Valor Desconto Unitario],"
    lSQL = lSQL & "       Movimento_Nota_Abastecimento.[Valor Unitario], Cliente.[Razao Social] as NomeCliente,"
    lSQL = lSQL & "       Produto.Nome as NomeProduto"
    lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento, Produto, Cliente"
    lSQL = lSQL & " WHERE Movimento_Nota_Abastecimento.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Data do Abastecimento] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Data do Abastecimento] <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.Periodo >= " & preparaTexto(Val(cbo_periodo_i.Text))
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.Periodo <= " & preparaTexto(Val(cbo_periodo_f.Text))
    If Val(cbo_tipo_movimento.Text) > 0 Then
        lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Tipo do Movimento] = " & preparaTexto(Val(cbo_tipo_movimento.Text))
    End If
    lSQL = lSQL & "   AND Produto.Codigo = Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    lSQL = lSQL & "   AND Cliente.Codigo = Movimento_Nota_Abastecimento.[Codigo do Cliente]"
    If chkOrdemAlfabetica.Value = 1 Then
        lSQL = lSQL & " ORDER BY Cliente.[Razao Social], [Data do Abastecimento], [Tipo do Movimento], Periodo, [Numero da Nota], [Codigo do Produto2]"
    Else
        lSQL = lSQL & " ORDER BY [Data do Abastecimento], [Tipo do Movimento], Periodo, [Numero da Nota], [Codigo do Produto2]"
    End If
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
    Dim xLinha As String
   
    
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
            ImpCliente
        End If
        If lLinha >= 60 Then
            If lSubQtd = 1 Then
                lSubQtd = 0
                lSubTotal = 0
                lSubTotalLiquido = 0
            End If
            xLinha = "+----------+--------+------------------------------------------+------------+---------+------------------+------------------------------+"
            Mid(xLinha, 25, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If rsTabela("Codigo do Cliente").Value <> l_cliente Or rsTabela("Codigo do Conveniado").Value <> l_conveniado Or (rsTabela("Numero da Nota").Value <> l_numero_nota And rsTabela("Numero do Cupom").Value <> l_numero_nota) Then
            ImpCliente
        End If
        ImpProduto
        lSubQtd = lSubQtd + 1
        lSubTotal = lSubTotal + rsTabela("Valor Total").Value
        lTotal = lTotal + rsTabela("Valor Total").Value
        rsTabela.MoveNext
    Loop
    If lTotal > 0 Then
        ImpSubTotal
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Abastecimento por Emissão|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpCliente()
    Dim xLinha As String
    Dim xString As String
    Dim i As Integer
    If lSubTotal > 0 Then
        ImpSubTotal
    End If
    xLinha = "|          |        |                                          |            |         |                  |                  |           |"
    i = Len(Format(rsTabela("Codigo do Cliente").Value, "#####"))
    Mid(xLinha, 15 + 5 - i, i) = Format(rsTabela("Codigo do Cliente").Value, "#####")
    Mid(xLinha, 23, 37) = rsTabela("NomeCliente").Value
    If rsTabela("Numero do Cupom").Value > 0 Then
        i = Len(rsTabela("Numero da Nota").Value)
        xString = Format(Mid(rsTabela("Numero da Nota").Value, 1, i - 2), "####,###") '& "." & Mid(rsTabela("Numero da Nota").Value, i - 1, 2)
        i = Len(xString)
        Mid(xLinha, 68 + 8 - i, i) = xString
    Else
        i = Len(Format(rsTabela("Numero da Nota").Value, "####,###"))
        Mid(xLinha, 68 + 8 - i, i) = Format(rsTabela("Numero da Nota").Value, "####,###")
    End If
    i = Len(Format(rsTabela("Codigo do Conveniado").Value, "###,###"))
    Mid(xLinha, 92 + 7 - i, i) = Format(rsTabela("Codigo do Conveniado").Value, "###,###")
    If rsTabela("Codigo do Conveniado").Value > 0 Then
        Mid(xLinha, 79, 12) = "Conveniado.:"
        If ClienteConveniado.LocalizarCodigo(rsTabela("Codigo do Cliente").Value, rsTabela("Codigo do Conveniado").Value) Then
            Mid(xLinha, 100, 37) = ClienteConveniado.Nome
        End If
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    l_cliente = rsTabela("Codigo do Cliente").Value
    l_conveniado = rsTabela("Codigo do Conveniado").Value
    If rsTabela("Numero do Cupom").Value > 0 Then
        l_numero_nota = rsTabela("Numero do Cupom").Value
    Else
        l_numero_nota = rsTabela("Numero da Nota").Value
    End If
    lLinha = lLinha + 1
End Sub
Private Sub ImpProduto()
    Dim xLinha As String
    Dim i As Integer
    Dim xValorLiquido As Currency
    
    xLinha = "|          |        |                                          |            |         |                  |                  |           |"
    Mid(xLinha, 2, 10) = Format(rsTabela("Data do Abastecimento").Value, "dd/mm/yyyy")
    i = Len(Format(rsTabela("Codigo do Produto2").Value, "#000"))
    Mid(xLinha, 16 + 4 - i, i) = Format(rsTabela("Codigo do Produto2").Value, "#000")
    Mid(xLinha, 23, 40) = rsTabela("NomeProduto").Value
    i = Len(Format(rsTabela("Quantidade").Value, "####,##0.00"))
    Mid(xLinha, 65 + 11 - i, i) = Format(rsTabela("Quantidade").Value, "####,##0.00")
    Mid(xLinha, 82, 1) = rsTabela("Periodo").Value
    i = Len(Format(rsTabela("Valor Total").Value, "###,###,##0.00"))
    Mid(xLinha, 91 + 14 - i, i) = Format(rsTabela("Valor Total").Value, "###,###,##0.00")
    
    xValorLiquido = rsTabela("Valor Total").Value
    If rsTabela("Valor Desconto Unitario").Value > 0 Then
        xValorLiquido = Format((rsTabela("Valor Unitario").Value - rsTabela("Valor Desconto Unitario").Value) * rsTabela("Quantidade").Value, "0000000000.00")
    ElseIf rsTabela("Valor Desconto Unitario").Value < 0 Then
        xValorLiquido = Format((rsTabela("Valor Unitario").Value - rsTabela("Valor Desconto Unitario").Value) * rsTabela("Quantidade").Value, "0000000000.00")
    End If
    i = Len(Format(xValorLiquido, "###,###,##0.00"))
    Mid(xLinha, 110 + 14 - i, i) = Format(xValorLiquido, "###,###,##0.00")
    lSubTotalLiquido = lSubTotalLiquido + xValorLiquido
    lTotalLiquido = lTotalLiquido + xValorLiquido
    Mid(xLinha, 131, 1) = rsTabela("Tipo do Movimento").Value
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "|          |        |                                          |            | * TOTAL |                  |                  |           |"
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(xLinha, 91 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    i = Len(Format(lSubTotalLiquido, "###,###,##0.00"))
    Mid(xLinha, 110 + 14 - i, i) = Format(lSubTotalLiquido, "###,###,##0.00")
    If lSubQtd > 1 Then
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    xLinha = "+----------+--------+------------------------------------------+------------+---------+------------------+------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    lSubQtd = 0
    lSubTotal = 0
    lSubTotalLiquido = 0
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "|                                                                     *** TOTAL GERAL |                  |                  |           |"
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 91 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(lTotalLiquido, "###,###,##0.00"))
    Mid(xLinha, 110 + 14 - i, i) = Format(lTotalLiquido, "###,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------------------------------------------------------------------+------------------+------------------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    Dim x_string_40 As String * 40
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
    x_string_40 = g_nome_empresa
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    '                   1         2         3         4         5         6         7         8
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    '                                              123456789012345678901234567890
    xLinha = "| RELAÇÃO DAS NOTAS DE ABASTECIMENTO POR EMISSÃO            CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    x_string_40 = Mid(cbo_tipo_movimento, 3, Len(cbo_tipo_movimento))
    BioImprime "@Printer.Print " & "| Tipo de Movimento.: " & x_string_40 & "                 |"
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i.Text & " a " & msk_data_f.Text & "       Período " & cbo_periodo_i & " ao " & cbo_periodo_f & "                   |"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+--------+------------------------------------------+------------+---------+------------------+------------------+-----------+"
    BioImprime "@Printer.Print " & "|          | CÓDIGO | RAZÃO SOCIAL                             |NUMERO NOTA |         |                  |                  |           |"
    BioImprime "@Printer.Print " & "|   DATA   | CÓDIGO | DESCRIÇÃO DOS PRODUTOS                   | QUANTIDADE | PERÍODO | VALOR DO PRODUTO | VALOR    LIQUIDO | TIPO MOV. |"
    BioImprime "@Printer.Print " & "+----------+--------+------------------------------------------+------------+---------+------------------+------------------+-----------+"
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
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
        SendKeys "{Tab}"
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
            DoEvents
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
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Escolha o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Escolha o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "O periodo final deve ser maior que " & Val(cbo_periodo_i) - 1 & ".", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            DoEvents
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
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
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
