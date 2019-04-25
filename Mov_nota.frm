VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form mov_nota_abastecimento 
   Caption         =   "Movimento de Notas de Abastecimento"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7875
   Icon            =   "Mov_nota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Mov_nota.frx":030A
   ScaleHeight     =   6675
   ScaleWidth      =   7875
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   1515
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2672
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "Mov_nota.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cria um novo registro."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "Mov_nota.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Altera o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "Mov_nota.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "Mov_nota.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "Mov_nota.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5700
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.Frame frmTotalNota 
         Caption         =   "Total de Notas No Período"
         Height          =   675
         Left            =   5220
         TabIndex        =   24
         Top             =   3060
         Width           =   2295
         Begin VB.Label lbl_total_notas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   500
            TabIndex        =   25
            Top             =   265
            Width           =   1275
         End
      End
      Begin VB.Data dta_produto 
         Caption         =   "dta_produto"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4860
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Produto"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data dta_cliente_conveniado 
         Caption         =   "dta_cliente_conveniado"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cliente_Conveniado"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data dta_cliente 
         Caption         =   "dta_cliente"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4980
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cliente"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   300
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   21
         ToolTipText     =   "Tecle CTRL+C p/ calcular litros vendidos!"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_unitario 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   16
         Top             =   2400
         Width           =   795
      End
      Begin VB.TextBox txt_cliente_conveniado 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox txt_numero_nota 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_total 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   23
         Top             =   3480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_abastecimento 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo dbcbo_cliente 
         Bindings        =   "Mov_nota.frx":7472
         Height          =   315
         Left            =   2940
         TabIndex        =   9
         Top             =   1320
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "razao social"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcbo_cliente_conveniado 
         Bindings        =   "Mov_nota.frx":748C
         Height          =   315
         Left            =   2940
         TabIndex        =   12
         Top             =   1680
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo do Conveniado"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcbo_produto 
         Bindings        =   "Mov_nota.frx":74B1
         Height          =   315
         Left            =   2880
         TabIndex        =   17
         Top             =   2400
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da nota"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Cl&iente conveniado"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço total"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do abastecimento"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5580
      TabIndex        =   34
      Top             =   5580
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "Mov_nota.frx":74CB
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "Mov_nota.frx":89C5
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "Mov_nota.frx":9EBF
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "Mov_nota.frx":B331
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "Mov_nota.frx":C8B3
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "Mov_nota.frx":DEBD
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5700
      Width           =   795
   End
End
Attribute VB_Name = "mov_nota_abastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_nota_abastecimento As Integer
Dim l_opcao As String
Dim l_data_abastecimento As Date
Dim l_periodo As String * 1
Dim l_nota As Long
Dim lOrdem As Integer
Dim l_cliente As Long
Dim l_codigo_produto As Long
Dim l_empresa As Integer
Dim lNumeroMovimentoCaixa As Long
Dim lSQL As String
Dim l_gravados As Long
Dim l_tipo_movimento As String * 1
Dim l_vezes As Integer
Dim l_qtd_periodo As Integer
Dim lCalcLitro As Boolean

Dim tbl_baixa_nota_abastecimento As Table
Dim tbl_cliente As Table
Dim tbl_cliente_conveniado As Table
Dim tbl_configuracao As Table
Dim tbl_desconto_personalizado As Table
Dim tbl_estoque As Table
Dim tbl_produto As Table

Private Cliente As New cCliente
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovCaixa As New cMovimentoCaixa
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private rsTabela As New adodb.Recordset
Function AlteraMovimentoCaixa() As Boolean
    AlteraMovimentoCaixa = False
    If Not MovCaixa.Excluir(g_empresa, l_data_abastecimento, lNumeroMovimentoCaixa) Then
        MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
    If IncluiMovimentoCaixa Then
        AlteraMovimentoCaixa = True
    End If
End Function
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    If g_nivel_acesso > 4 Then
        If g_empresa < g_cfg_empresa_i Or g_empresa > g_cfg_empresa_f Then
            cmd_novo.Enabled = False
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub ChamaCalcLitros()
    g_valor = fValidaValor4(txt_valor_unitario.Text)
    calc_litro.Show 1
    lCalcLitro = True
    'txt_quantidade.Text = Format(g_valor, "###,##0.00")
    'txt_valor_total.Text = Format(Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor2(txt_quantidade.Text), "###,##0.0"), "###,##0.00")
    txt_quantidade.Text = Format(RetiraGString(1), "###,##0.00")
    txt_valor_total.Text = Format(RetiraGString(2), "###,##0.00")
    cmd_ok.SetFocus
End Sub
Private Sub ChamaProdutoEspecial()
    If UCase(g_nome_empresa) Like "*MARQUES*" Then
        txt_produto.Text = 1
    Else
        txt_produto.Text = 120
    End If
    tbl_produto.Seek "=", CLng(txt_produto.Text)
    If Not tbl_produto.NoMatch Then
        dbcbo_produto.BoundText = CLng(txt_produto.Text)
        txt_valor_unitario.Text = Format(tbl_produto![Preco de Venda], "###,##0.0000")
    Else
        MsgBox "Produto não cadastro.", vbInformation, "Atenção!"
        txt_produto.SetFocus
        Exit Sub
    End If
    ChamaCalcLitros
End Sub
Private Sub Inclui()
    l_opcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    
    xComplemento = "TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & cbo_periodo.ItemData(cbo_periodo.ListIndex) & " " & dbcbo_cliente.Text
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(msk_data_abastecimento.Text)
    MovCaixa.NumeroMovimento = 1
    MovCaixa.Valor = fValidaValor(txt_valor_total.Text)
    MovCaixa.NumeroDocumento = txt_numero_nota.Text
    MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
    MovCaixa.Complemento = xComplemento
    MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
    MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = False
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixa = True
    End If
End Function
Private Sub AtualizaConstantes()
    tbl_configuracao.Index = "id_codigo"
    tbl_configuracao.Seek "=", g_empresa
    If Not tbl_configuracao.NoMatch Then
        l_qtd_periodo = tbl_configuracao![Quantidade de Periodos]
    Else
        l_qtd_periodo = 1
    End If
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    Dim x_total As Currency
    LimpaMSFlexGrid
    lSQL = "Select Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.Periodo, Movimento_Nota_Abastecimento.[Tipo do Movimento], Movimento_Nota_Abastecimento.[Codigo do Cliente], Cliente.[Razao Social], Produto.Nome as NomeProduto, Movimento_Nota_Abastecimento.[Valor Total], Cliente_Conveniado.Nome as NomeConveniado"
    lSQL = lSQL & " From Movimento_Nota_Abastecimento, Cliente, Produto, Cliente_Conveniado"
    lSQL = lSQL & " Where Cliente.Codigo = Movimento_Nota_Abastecimento.[Codigo do Cliente]"
    lSQL = lSQL & " And Produto.Codigo = Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    lSQL = lSQL & " And Cliente_Conveniado.[Codigo do Conveniado] = Movimento_Nota_Abastecimento.[Codigo do Conveniado]"
    lSQL = lSQL & " And Movimento_Nota_Abastecimento.Empresa = " & g_empresa
    lSQL = lSQL & " And Movimento_Nota_Abastecimento.[Data do Abastecimento] = #" & CDate(Format(l_data_abastecimento, "mm/dd/yyyy")) & "#"
    lSQL = lSQL & " And Movimento_Nota_Abastecimento.Periodo = " & Chr(34) & l_periodo & Chr(34)
    lSQL = lSQL & " And Movimento_Nota_Abastecimento.[Tipo do Movimento] = " & Chr(34) & l_tipo_movimento & Chr(34)
    lSQL = lSQL & " Order by Movimento_Nota_Abastecimento.[Numero da Nota], Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    x_total = 0
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela("Data do Abastecimento").Value
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela("Periodo").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Tipo do Movimento").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Codigo do Cliente").Value
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela("Razao Social").Value
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela("NomeProduto").Value
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = Format(rsTabela("Valor Total").Value, "###,###,##0.00")
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela("NomeConveniado").Value
            x_total = x_total + rsTabela("Valor Total").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    lbl_total_notas = Format(x_total, "###,##0.00")
    If Val(l_tipo_movimento) = 0 Then
        l_periodo = "1"
        l_tipo_movimento = "1"
    End If
End Sub
Private Sub AtualTabe()
    MovNotaAbastecimento.Empresa = g_empresa
    MovNotaAbastecimento.DataAbastecimento = Format(msk_data_abastecimento.Text, "dd/mm/yyyy")
    MovNotaAbastecimento.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovNotaAbastecimento.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovNotaAbastecimento.CodigoCliente = CLng(dbcbo_cliente.BoundText)
    If dbcbo_cliente_conveniado.BoundText <> "" Then
        MovNotaAbastecimento.CodigoConveniado = CLng(dbcbo_cliente_conveniado.BoundText)
    Else
        MovNotaAbastecimento.CodigoConveniado = 0
        MovNotaAbastecimento.BaixadoPelaDuplicata = False
    End If
    MovNotaAbastecimento.NumeroNota = Val(txt_numero_nota.Text)
    MovNotaAbastecimento.Ordem = 1
    MovNotaAbastecimento.CodigoProduto2 = CLng(dbcbo_produto.BoundText)
    MovNotaAbastecimento.ValorUnitario = fValidaValor4(txt_valor_unitario.Text)
    MovNotaAbastecimento.Quantidade = fValidaValor2(txt_quantidade.Text)
    MovNotaAbastecimento.ValorTotal = fValidaValor2(txt_valor_total.Text)
    If l_opcao = 1 Then
        MovNotaAbastecimento.NumeroCupom = 0
        MovNotaAbastecimento.DataConferencia = "00:00:00"
    End If
    MovNotaAbastecimento.ValorDescontoUnitario = ValorDescontoUnitario(MovNotaAbastecimento.CodigoCliente, MovNotaAbastecimento.CodigoProduto2, MovNotaAbastecimento.ValorUnitario)
    MovNotaAbastecimento.NumeroMovimentoCaixa = MovCaixa.NumeroMovimento
    MovNotaAbastecimento.NumeroIlha = 1
End Sub
Private Sub AtualTela()
    Dim i As Integer
    l_data_abastecimento = MovNotaAbastecimento.DataAbastecimento
    l_periodo = MovNotaAbastecimento.Periodo
    l_nota = MovNotaAbastecimento.NumeroNota
    lOrdem = MovNotaAbastecimento.Ordem
    l_cliente = MovNotaAbastecimento.CodigoCliente
    l_codigo_produto = MovNotaAbastecimento.CodigoProduto2
    l_tipo_movimento = MovNotaAbastecimento.TipoMovimento
    lNumeroMovimentoCaixa = MovNotaAbastecimento.NumeroMovimentoCaixa
    msk_data_abastecimento.Text = Format(MovNotaAbastecimento.DataAbastecimento, "dd/mm/yyyy")
    cbo_periodo.ListIndex = -1
    For i = 0 To cbo_periodo.ListCount - 1
        If cbo_periodo.ItemData(i) = MovNotaAbastecimento.Periodo Then
            cbo_periodo.ListIndex = i
            Exit For
        End If
    Next
    cbo_tipo_movimento.ListIndex = -1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        If cbo_tipo_movimento.ItemData(i) = MovNotaAbastecimento.TipoMovimento Then
            cbo_tipo_movimento.ListIndex = i
            Exit For
        End If
    Next
    txt_cliente.Text = MovNotaAbastecimento.CodigoCliente
    dbcbo_cliente.BoundText = MovNotaAbastecimento.CodigoCliente
    txt_cliente_conveniado = Format(MovNotaAbastecimento.CodigoConveniado, "######")
    If MovNotaAbastecimento.CodigoConveniado > 0 Then
        dbcbo_cliente_conveniado.BoundText = MovNotaAbastecimento.CodigoConveniado
    Else
        dbcbo_cliente_conveniado.BoundText = ""
    End If
    txt_numero_nota.Text = MovNotaAbastecimento.NumeroNota
    txt_produto.Text = MovNotaAbastecimento.CodigoProduto2
    tbl_produto.Seek "=", l_codigo_produto
    If Not tbl_produto.NoMatch Then
        dbcbo_produto.BoundText = MovNotaAbastecimento.CodigoProduto2
    Else
        dbcbo_produto.BoundText = ""
    End If
    txt_valor_unitario.Text = Format(MovNotaAbastecimento.ValorUnitario, "###,##0.0000")
    txt_quantidade.Text = Format(MovNotaAbastecimento.Quantidade, "###,##0.00")
    txt_valor_total.Text = Format(MovNotaAbastecimento.ValorTotal, "###,##0.00")
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    tbl_baixa_nota_abastecimento.Close
    tbl_cliente.Close
    tbl_cliente_conveniado.Close
    tbl_configuracao.Close
    tbl_desconto_personalizado.Close
    tbl_estoque.Close
    tbl_produto.Close
    
    Set Cliente = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovCaixa = Nothing
    Set MovNotaAbastecimento = Nothing
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo.Clear
    cbo_periodo.AddItem 1
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 1
    cbo_periodo.AddItem 2
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 2
    cbo_periodo.AddItem 3
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 3
    cbo_periodo.AddItem 4
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 4
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Notas Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Function ValidaClienteConveniado() As Boolean
    ValidaClienteConveniado = False
    If Not tbl_cliente.NoMatch Then
        If tbl_cliente![Codigo do Convenio] > 1 And dbcbo_cliente_conveniado.BoundText <> "" Then
            ValidaClienteConveniado = True
        ElseIf tbl_cliente![Codigo do Convenio] = 1 Then
            ValidaClienteConveniado = True
        End If
    End If
End Function
Function ValorDescontoUnitario(ByVal xCliente As Long, ByVal xProduto As Long, ByVal xValorUnitario As Currency) As Currency
    ValorDescontoUnitario = 0
    With tbl_desconto_personalizado
        If .RecordCount > 0 Then
            .Seek "=", xCliente, xProduto
            If Not .NoMatch Then
                If ![Percentual a Descontar] <> 0 Then
                    ValorDescontoUnitario = CCur(xValorUnitario * ![Percentual a Descontar] / 100)
                Else
                    ValorDescontoUnitario = ![Valor a Descontar]
                End If
            End If
        End If
    End With
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovNotaAbastecimento.Empresa < g_cfg_empresa_i Or MovNotaAbastecimento.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovNotaAbastecimento.DataAbastecimento < g_cfg_data_i Or MovNotaAbastecimento.DataAbastecimento > g_cfg_data_f Then
            x_flag = False
        ElseIf MovNotaAbastecimento.Periodo < g_cfg_periodo_i Or MovNotaAbastecimento.Periodo > g_cfg_periodo_f Then
            x_flag = False
        End If
    End If
    If x_flag Then
        cmd_alterar.Enabled = True
        cmd_excluir.Enabled = True
    Else
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data_abastecimento < g_cfg_data_i Or msk_data_abastecimento > g_cfg_data_f Then
        MsgBox "A data de abastecimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data_abastecimento.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaNotaAbastecimento
    End If
End Sub
Private Sub dbcbo_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dbcbo_produto_LostFocus()
    If dbcbo_produto.BoundText <> "" And l_opcao > 0 Then
        txt_produto = dbcbo_produto.BoundText
        tbl_produto.Seek "=", CLng(txt_produto)
        If Not tbl_produto.NoMatch Then
            tbl_estoque.Seek "=", g_empresa, CLng(txt_produto)
            If Not tbl_estoque.NoMatch Then
                If tbl_estoque![Preco de Venda] <> 0 Then
                    txt_valor_unitario = Format(tbl_estoque![Preco de Venda], "###,##0.0000")
                Else
                    txt_valor_unitario = Format(tbl_produto![Preco de Venda], "###,##0.0000")
                End If
            Else
                MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                txt_valor_unitario = ""
                txt_valor_unitario.SetFocus
                Exit Sub
            End If
        End If
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    l_opcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_quantidade.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovNotaAbastecimento.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    l_opcao = 0
    LimpaTela
    If MovNotaAbastecimento.LocalizarCodigo(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
        AtualTela
        AtualizaMSFlexGrid
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 7
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data abast."
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Per."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Tipo mov."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Cliente"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Razão social"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Produto"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor Total"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = 6
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Conveniado"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = 1
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub LimpaTela()
    If l_gravados = 0 Then
        msk_data_abastecimento = "__/__/____"
        cbo_periodo.ListIndex = -1
        cbo_tipo_movimento.ListIndex = -1
        txt_cliente = ""
        dbcbo_cliente.BoundText = ""
        txt_cliente_conveniado = ""
        dbcbo_cliente_conveniado.BoundText = ""
        txt_numero_nota = ""
    End If
    txt_produto = ""
    dbcbo_produto.BoundText = ""
    txt_valor_unitario = ""
    txt_quantidade = ""
    txt_valor_total = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_numero_nota.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = vbYes Then
            l_opcao = 3
            If MovNotaAbastecimento.Excluir(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                LimpaTela
                If Not MovCaixa.Excluir(g_empresa, l_data_abastecimento, lNumeroMovimentoCaixa) Then
                    MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
                End If
                If MovNotaAbastecimento.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtualizaMSFlexGrid
                    AtivaBotoes
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
                l_opcao = 0
            Else
                MsgBox "Registro não excluido!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If l_gravados = 0 Then
        If BuscaProximoCaixa Then
            txt_cliente.SetFocus
        Else
            msk_data_abastecimento.SetFocus
        End If
    Else
        If UCase(g_nome_empresa) Like "*MARQUES*" Then
            dbcbo_cliente.SetFocus
        Else
            txt_produto.SetFocus
        End If
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            g_string = ""
            If l_opcao = 1 Then
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                End If
                AtualTabe
                If MovNotaAbastecimento.Incluir Then
                    l_data_abastecimento = msk_data_abastecimento.Text
                    l_periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
                    l_nota = txt_numero_nota.Text
                    lOrdem = 1
                    l_cliente = CLng(dbcbo_cliente.BoundText)
                    l_codigo_produto = CLng(dbcbo_produto.BoundText)
                    l_tipo_movimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
                    l_gravados = 1
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
            ElseIf l_opcao = 2 Then
                If Not AlteraMovimentoCaixa Then
                    MsgBox "Não foi possível alterar este registro do Caixa!", vbInformation, "Erro de Integridade."
                End If
                AtualTabe
                If MovNotaAbastecimento.Alterar(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                    l_data_abastecimento = msk_data_abastecimento.Text
                    l_periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
                    l_nota = txt_numero_nota.Text
                    lOrdem = 1
                    l_cliente = CLng(dbcbo_cliente.BoundText)
                    l_codigo_produto = CLng(dbcbo_produto.BoundText)
                    l_tipo_movimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
            End If
            If MovNotaAbastecimento.LocalizarCodigo(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                AtualTela
                AtualizaMSFlexGrid
            Else
                LimpaTela
                MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
            End If
            If l_opcao = 1 Then
                l_opcao = 0
                cmd_novo_Click
                If g_caixa_unificado Then
                    txt_cliente.SetFocus
                End If
                If Val(txt_cliente) = 26 Then
                    txt_cliente_conveniado.SetFocus
                End If
            Else
                l_opcao = 0
                cmd_alterar.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_abastecimento) Then
        MsgBox "Informe a data de abastecimento.", vbInformation, "Atenção!"
        msk_data_abastecimento.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf dbcbo_cliente.BoundText = "" Then
        MsgBox "Escolha o cliente.", vbInformation, "Atenção!"
        dbcbo_cliente.SetFocus
    ElseIf Not ValidaClienteConveniado Then
        MsgBox "Escolha o cliente_conveniado.", vbInformation, "Atenção!"
        dbcbo_cliente_conveniado.SetFocus
    ElseIf Not Val(txt_numero_nota) > 0 Then
        MsgBox "Informe o número da nota.", vbInformation, "Atenção!"
        txt_numero_nota.SetFocus
    ElseIf dbcbo_produto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dbcbo_produto.SetFocus
    ElseIf Not fValidaValor4(txt_valor_unitario) > 0 Then
        MsgBox "Informe o valor unitário do produto.", vbInformation, "Atenção!"
        txt_valor_unitario.SetFocus
    ElseIf Not fValidaValor2(txt_quantidade) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor2(txt_valor_total) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_movimento_nota.Show 1
    If Len(g_string) > 0 Then
        l_data_abastecimento = RetiraGString(1)
        l_periodo = RetiraGString(2)
        l_nota = RetiraGString(3)
        lOrdem = RetiraGString(4)
        lOrdem = RetiraGString(5)
        l_cliente = RetiraGString(6)
        l_codigo_produto = RetiraGString(7)
        If MovNotaAbastecimento.LocalizarCodigo(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovNotaAbastecimento.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovNotaAbastecimento.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If MovNotaAbastecimento.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dbcbo_cliente_conveniado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_numero_nota.SetFocus
    End If
End Sub
Private Sub dbcbo_cliente_conveniado_LostFocus()
    If l_opcao > 0 Then
        If txt_cliente_conveniado <> dbcbo_cliente_conveniado.BoundText Then
            l_vezes = 1
        End If
        If l_vezes = 1 Then
            l_vezes = l_vezes + 1
            If dbcbo_cliente_conveniado.BoundText <> "" Then
                tbl_cliente_conveniado.Seek "=", tbl_cliente![Codigo do Convenio], CLng(dbcbo_cliente_conveniado.BoundText)
                If Not tbl_cliente_conveniado.NoMatch Then
                    txt_cliente_conveniado = tbl_cliente_conveniado![Codigo do Conveniado]
                Else
                    dbcbo_cliente_conveniado.BoundText = ""
                End If
            End If
            If l_opcao = 1 Then
                txt_numero_nota.Text = MovNotaAbastecimento.ProximoNumeroNota(g_empresa, CDate(msk_data_abastecimento.Text))
                If dbcbo_cliente_conveniado.BoundText <> "" Then
                    If UCase(g_nome_empresa) <> "AUTO POSTO MANTIQUEIRA LTDA" Then
                        ChamaProdutoEspecial
                        Exit Sub
                    End If
                End If
            End If
            txt_produto.SetFocus
        End If
    End If
End Sub
Private Sub dbcbo_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_cliente_conveniado.SetFocus
    End If
End Sub
Private Sub dbcbo_cliente_LostFocus()
    If dbcbo_cliente.BoundText <> "" And l_opcao > 0 Then
        tbl_cliente.Seek "=", Val(dbcbo_cliente.BoundText)
        If Not tbl_cliente.NoMatch Then
            txt_cliente = tbl_cliente!Codigo
            If tbl_cliente![Codigo do Convenio] = 1 Then
                If l_opcao = 1 Then
                    txt_numero_nota.Text = MovNotaAbastecimento.ProximoNumeroNota(g_empresa, CDate(msk_data_abastecimento.Text))
                End If
                txt_cliente_conveniado.Text = ""
                dbcbo_cliente_conveniado.BoundText = ""
                If g_caixa_unificado Or UCase(g_nome_empresa) Like "*VIA 63*" Or UCase(g_nome_empresa) Like "*MARQUES*" Then
                    txt_numero_nota.SetFocus
                Else
                    txt_produto.SetFocus
                End If
                Exit Sub
            Else
                dta_cliente_conveniado.RecordSource = "Select * From Cliente_Conveniado Where [Codigo do Convenio] = " & tbl_cliente![Codigo do Convenio] & " Order By Nome"
                dta_cliente_conveniado.Refresh
                txt_cliente_conveniado.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_movimento_nota_abastecimento = 0
    End If
    If flag_movimento_nota_abastecimento = 0 Then
        AtualizaConstantes
        l_gravados = 0
        l_opcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If RetiraGString(1) = "ConsultaNotaAbastecimento" Then
            AtivaBotoes
            l_data_abastecimento = RetiraGString(2)
            l_periodo = RetiraGString(3)
            l_nota = RetiraGString(4)
            lOrdem = RetiraGString(5)
            l_cliente = RetiraGString(6)
            l_codigo_produto = RetiraGString(7)
            If MovNotaAbastecimento.LocalizarCodigo(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                AtualTela
                AtualizaMSFlexGrid
                cmd_alterar.SetFocus
            End If
        Else
            If MovNotaAbastecimento.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtualizaMSFlexGrid
                AtivaBotoes
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            If cmd_novo.Enabled Then
                cmd_novo.SetFocus
            End If
        End If
    Else
        flag_movimento_nota_abastecimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_nota_abastecimento = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF5 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF7 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    Set tbl_baixa_nota_abastecimento = bd_sgp.OpenTable("Baixa_Nota_Abastecimento")
    Set tbl_cliente = bd_sgp.OpenTable("Cliente")
    Set tbl_cliente_conveniado = bd_sgp.OpenTable("Cliente_Conveniado")
    Set tbl_configuracao = bd_sgp.OpenTable("configuracao")
    Set tbl_desconto_personalizado = bd_sgp.OpenTable("Movimento_Desconto_Personalizado")
    Set tbl_estoque = bd_sgp.OpenTable("Estoque")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    
    tbl_baixa_nota_abastecimento.Index = "id_cliente_pagamento"
    tbl_cliente.Index = "id_codigo"
    tbl_cliente_conveniado.Index = "id_codigo"
    tbl_desconto_personalizado.Index = "id_cliente"
    tbl_estoque.Index = "id_codigo2"
    tbl_produto.Index = "id_codigo"
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    dta_cliente.RecordSource = "Select * From Cliente Where Inativo = False Order By [Razao Social]"
    dta_cliente.Refresh
    dta_cliente_conveniado.RecordSource = "Select * From Cliente_Conveniado Order By Nome"
    dta_cliente_conveniado.Refresh
    dta_produto.RecordSource = "Select * From Produto Where Inativo = False Order By Nome"
    dta_produto.Refresh
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "NOTA ABASTECIMENTO") Then
        MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
    End If
    lCalcLitro = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_abastecimento_GotFocus()
    msk_data_abastecimento.SelStart = 0
    msk_data_abastecimento.SelLength = 2
End Sub
Private Sub msk_data_abastecimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub txt_cliente_conveniado_GotFocus()
    l_vezes = 0
    If l_opcao = 1 Then
        If IsNumeric(txt_cliente_conveniado) Then
            dbcbo_cliente_conveniado.BoundText = CLng(txt_cliente_conveniado)
        End If
        txt_cliente_conveniado = ""
    End If
End Sub
Private Sub txt_cliente_conveniado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_cliente_conveniado.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_conveniado_LostFocus()
    l_vezes = l_vezes + 1
    If Val(txt_cliente_conveniado) > 0 Then
        tbl_cliente_conveniado.Seek "=", tbl_cliente![Codigo do Convenio], CLng(txt_cliente_conveniado)
        If Not tbl_cliente_conveniado.NoMatch Then
            dbcbo_cliente_conveniado.BoundText = CLng(txt_cliente_conveniado)
            If l_opcao = 1 Then
                txt_numero_nota.Text = MovNotaAbastecimento.ProximoNumeroNota(g_empresa, CDate(msk_data_abastecimento.Text))
            End If
            dbcbo_cliente_conveniado_LostFocus
            Exit Sub
        Else
            MsgBox "Cliente conveniado não cadastro.", vbInformation, "Atenção!"
            dbcbo_cliente_conveniado.BoundText = ""
            txt_cliente_conveniado.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    If l_opcao = 1 Then
        txt_cliente = ""
    End If
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_cliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    If Val(txt_cliente) > 0 Then
        tbl_cliente.Seek "=", Val(txt_cliente)
        If Not tbl_cliente.NoMatch Then
            If tbl_cliente!Inativo = True Then
                MsgBox "O cliente " & Trim(tbl_cliente![Razao Social]) & " está inativo.", vbInformation, "Cliente Inativo!"
                txt_cliente.SetFocus
                Exit Sub
            Else
                dbcbo_cliente.BoundText = Val(txt_cliente)
                dbcbo_cliente_LostFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            dbcbo_cliente.BoundText = ""
            txt_cliente.SetFocus
        End If
    End If
End Sub
Private Sub txt_numero_nota_GotFocus()
    txt_numero_nota.SelStart = 0
    txt_numero_nota.SelLength = Len(txt_numero_nota.Text)
End Sub
Private Sub txt_numero_nota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_produto.SetFocus
    End If
End Sub
Private Sub txt_produto_LostFocus()
    Dim i As Integer
    If Val(txt_produto) > 0 And l_opcao > 0 Then
        tbl_produto.Seek "=", CLng(txt_produto)
        If Not tbl_produto.NoMatch Then
            If tbl_produto!Inativo = True Then
                MsgBox "O produto " & Trim(tbl_produto!Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            Else
                dbcbo_produto.BoundText = CLng(txt_produto)
                tbl_estoque.Seek "=", g_empresa, CLng(txt_produto)
                If Not tbl_estoque.NoMatch Then
                    If tbl_estoque![Preco de Venda] <> 0 Then
                        txt_valor_unitario = Format(tbl_estoque![Preco de Venda], "###,##0.0000")
                    Else
                        txt_valor_unitario = Format(tbl_produto![Preco de Venda], "###,##0.0000")
                    End If
                Else
                    MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                    txt_valor_unitario = ""
                    txt_valor_unitario.SetFocus
                    Exit Sub
                End If
            End If
            txt_quantidade.SetFocus
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_quantidade_GotFocus()
    txt_quantidade.SelStart = 0
    txt_quantidade.SelLength = Len(txt_quantidade.Text)
    g_string = ""
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 3 Then
        KeyAscii = 0
        ChamaCalcLitros
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_total.SetFocus
    End If
End Sub
Private Sub txt_quantidade_LostFocus()
    txt_quantidade.Text = Format(txt_quantidade.Text, "###,##0.00")
    If fValidaValor(txt_quantidade.Text) > 0 Then
        If lCalcLitro = False Then
            txt_valor_total.Text = Format(Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor2(txt_quantidade.Text), "###,##0.0"), "###,##0.00")
        Else
            lCalcLitro = False
        End If
    Else
        g_string = ""
    End If
End Sub
Private Sub txt_valor_total_GotFocus()
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total)
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total = Format(txt_valor_total, "###,##0.00")
End Sub
Private Sub txt_valor_unitario_GotFocus()
    txt_valor_unitario.SelStart = 0
    txt_valor_unitario.SelLength = Len(txt_valor_unitario.Text)
End Sub
Private Sub txt_valor_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub txt_valor_unitario_LostFocus()
    txt_valor_unitario = Format(txt_valor_unitario, "###,##0.0000")
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    If MovNotaAbastecimento.LocalizarUltimo(g_empresa) Then
        msk_data_abastecimento.Text = Format(MovNotaAbastecimento.DataAbastecimento, "dd/mm/yyyy")
        x_periodo = MovNotaAbastecimento.Periodo
        If MovNotaAbastecimento.Periodo >= l_qtd_periodo Then
            msk_data_abastecimento.Text = Format(MovNotaAbastecimento.DataAbastecimento + 1, "dd/mm/yyyy")
            x_periodo = 0
        End If
        cbo_periodo.ListIndex = x_periodo
        cbo_tipo_movimento.ListIndex = 0
        BuscaProximoCaixa = True
    Else
        msk_data_abastecimento.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
    End If
End Function
Private Sub ProcessaNotaAbastecimento()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovNotaAbastecimento.LocalizarPrimeiro() Then
        If MovNotaAbastecimento.Empresa = g_empresa And MovNotaAbastecimento.DataAbastecimento >= xData Then
            AtualTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovNotaAbastecimento.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If Not MovNotaAbastecimento.Alterar(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                    MsgBox "Erro ao alterar nota de abastecimento", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovNotaAbastecimento.LocalizarProximo = False
            If MovNotaAbastecimento.Empresa = g_empresa And MovNotaAbastecimento.DataAbastecimento >= xData Then
                AtualTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovNotaAbastecimento.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                    If Not MovNotaAbastecimento.Alterar(g_empresa, l_cliente, l_data_abastecimento, l_nota, lOrdem, l_codigo_produto, l_periodo) Then
                        MsgBox "Erro ao alterar nota de abastecimento", vbInformation, "Erro"
                    End If
                End If
            End If
        Loop
    
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    
    
'Baixa de Notas de Abastecimento
    
    '+Codigo do Cliente;+Data do Abastecimento;+Codigo do Produto2;+Numero da Nota;+Data do Pagamento;+Empresa;+Periodo
    tbl_baixa_nota_abastecimento.Seek ">", 0
    If Not tbl_baixa_nota_abastecimento.NoMatch Then
        Do Until tbl_baixa_nota_abastecimento.EOF
            If tbl_baixa_nota_abastecimento!Empresa = g_empresa And tbl_baixa_nota_abastecimento![Data do Abastecimento] >= xData Then
                If Not IncluiMovimentoCaixaBaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    tbl_baixa_nota_abastecimento.Edit
                    tbl_baixa_nota_abastecimento![Numero do Movimento do Caixa] = lNumeroMovimentoCaixa
                    tbl_baixa_nota_abastecimento.Update
                End If
            End If
            tbl_baixa_nota_abastecimento.MoveNext
        Loop
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    MsgBox "Processamento concluído!"
    Exit Sub
FileError:
    MsgBox "Erro ao processar Notas de Abastecimento Baixada", vbInformation, "ProcessaNotaAbastecimento"
End Sub
Function IncluiMovimentoCaixaBaixa() As Boolean
    Dim xComplemento As String
    Dim xNomeCliente As String
    IncluiMovimentoCaixaBaixa = False
    xNomeCliente = "Cliente Excluído"
    If Cliente.LocalizarCodigo(tbl_baixa_nota_abastecimento![Codigo do Cliente]) Then
        xNomeCliente = Cliente.RazaoSocial
    End If
    xComplemento = "TM:" & tbl_baixa_nota_abastecimento![Tipo do Movimento] & " P:" & tbl_baixa_nota_abastecimento!Periodo & " " & xNomeCliente
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = tbl_baixa_nota_abastecimento![Data do Abastecimento]
    MovCaixa.NumeroMovimento = 1
    MovCaixa.Valor = tbl_baixa_nota_abastecimento![Valor Total]
    MovCaixa.NumeroDocumento = tbl_baixa_nota_abastecimento![Numero da Nota]
    MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
    MovCaixa.Complemento = xComplemento
    MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
    MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = False
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixaBaixa = True
    End If
End Function

