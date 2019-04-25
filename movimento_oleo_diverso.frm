VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_oleo_diverso 
   Caption         =   "Movimento de Óleos/Filtros e Diversos"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   8850
   Icon            =   "movimento_oleo_diverso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_oleo_diverso.frx":030A
   ScaleHeight     =   6840
   ScaleWidth      =   8850
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_oleo_diverso.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cria um novo registro."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_oleo_diverso.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Altera o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_oleo_diverso.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_oleo_diverso.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_oleo_diverso.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5880
      Width           =   795
   End
   Begin MSGrid.Grid grid_oleo 
      Height          =   2535
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   4471
      _StockProps     =   77
      BackColor       =   16777215
      Cols            =   11
      FixedCols       =   0
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8655
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
         Height          =   315
         Left            =   7560
         TabIndex        =   17
         Top             =   1680
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoMovimento 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   3315
      End
      Begin VB.ComboBox cboIlha 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton btnTransfereVendaECF 
         Caption         =   "&Transfere do ECF"
         Height          =   795
         Left            =   6000
         Picture         =   "movimento_oleo_diverso.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Transfere a venda do ECF."
         Top             =   2220
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1320
         Width           =   555
      End
      Begin VB.ComboBox cboTipoSubEstoque 
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_unitario 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txt_valor_total 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   4140
         Top             =   1320
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "movimento_oleo_diverso.frx":8864
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   1320
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   4200
         Top             =   1680
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcProduto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "movimento_oleo_diverso.frx":8883
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   1680
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
      End
      Begin VB.Frame frmTotalNota 
         Caption         =   "Total de Vendas do Caixa"
         Height          =   675
         Left            =   3660
         TabIndex        =   38
         Top             =   2160
         Width           =   2295
         Begin VB.Label lblTotalVenda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   500
            TabIndex        =   39
            Top             =   265
            Width           =   1275
         End
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do Movimento"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Número da &Ilha"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do SubEstoque"
         Height          =   315
         Index           =   7
         Left            =   3780
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   315
         Index           =   6
         Left            =   3780
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço total"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do movimento"
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
      Left            =   6600
      TabIndex        =   32
      Top             =   5760
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_oleo_diverso.frx":889E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_oleo_diverso.frx":9D98
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_oleo_diverso.frx":B292
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_oleo_diverso.frx":C704
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7080
      Picture         =   "movimento_oleo_diverso.frx":DC86
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7980
      Picture         =   "movimento_oleo_diverso.frx":F290
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5880
      Width           =   795
   End
End
Attribute VB_Name = "movimento_oleo_diverso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_oleo_diverso As Integer
Dim lOpcao As String
Dim lEmpresa As Integer
Dim lData As Date
Dim lPeriodo As String
Dim lIlha As Integer
Dim lTipoSubEstoque As String
Dim lTipoMovimento As Integer
Dim lOrdem As Integer
Dim lCodigoProduto As Long
Dim lCodigoFuncionario As Integer
Dim lValorAnterior As Currency
Dim lNumeroMovimentoCaixa As Long
Dim lSQL As String
Dim lGravados As Long
Dim lPrecoCusto As Currency
Dim lQuantidade As Currency
'Dim l_vezes As Integer
Dim lQtdPeriodo As Integer
Dim lBaixaAutomaticaNoEstoque As Boolean
Dim lDataI As Date
Dim lDataF As Date
Dim lPeriodoI As Integer
Dim lPeriodoF As Integer
Dim lAlteraPrecoCadastro As Boolean
Dim lPriorizaSeguranca As Boolean
Dim lCaixaIndividual As Boolean

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxIlha As Integer
Dim lCxTipoMovimento As Integer
Dim lCxSubEstoque As Integer
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoFuncionario As Integer
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxOperacao As Integer
Dim lCxCodigoUsuario As Integer

Private lRS As adodb.Recordset

Private AberturaCaixa As New cAberturaCaixa
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private Estoque As New cEstoque
Private Funcionario As New cFuncionario
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovimentoLubrificante As New cMovimentoLubrificante
Private Produto As New cProduto
Private SubEstoque As New cSubEstoque
Private TransfInternaEstoque As New cTransfInternaEstoque

Private Sub AjustaCaixaPista()
    Dim xString As String
    Dim xOperacao As String
    
    xString = g_string
    xOperacao = RetiraString(2, xString)
    g_string = ""

    lCxData = CDate(RetiraString(3, xString))
    lCxPeriodo = RetiraString(4, xString)
    lCxTipoMovimento = Val(RetiraString(5, xString))
    lCxIlha = Val(RetiraString(6, xString))
    lCxSubEstoque = Val(RetiraString(7, xString))
    lCxCodigoFuncionario = Val(RetiraString(8, xString))
    lCxCodigoLancamentoPadrao = Val(RetiraString(9, xString))
    lDataI = lCxData
    lDataF = lCxData
    lPeriodoI = Val(lCxPeriodo)
    lPeriodoF = Val(lCxPeriodo)
    If xOperacao = "Incluir" Then
        lCxOperacao = 1
    ElseIf xOperacao = "Alterar" Then
        lCxOperacao = 2
    Else
        lCxOperacao = 3
    End If

    lData = lCxData
    lPeriodo = lCxPeriodo
    txtData.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    'cbo_tipo_movimento.Enabled = False
    If xOperacao = "Incluir" Then
        lCxCodigoUsuario = Val(RetiraString(10, xString))
        cmd_novo_Click
        AtualizaGrid
    ElseIf xOperacao = "Alterar" Then
        txtData.Text = Format(lCxData, "dd/mm/yyyy")
        cbo_periodo.ListIndex = lCxPeriodo - 1
        cboIlha.ListIndex = lCxIlha - 1
        cboTipoSubEstoque.ListIndex = lCxSubEstoque - 1
        cboTipoMovimento.ListIndex = lCxTipoMovimento - 1
        txt_funcionario.Text = lCxCodigoFuncionario
        dtcboFuncionario.BoundText = lCxCodigoFuncionario
        lCxCodigoUsuario = Val(RetiraString(10, xString))
        AtualizaGrid
        If grid_oleo.Rows = 2 Then
            MsgBox "Não existe venda a ser alterada!", vbInformation, "Erro de Integridade!"
            cmd_cancelar_Click
            Exit Sub
        End If
        grid_oleo.Row = grid_oleo.Rows - 2
        MarcaCelulaOleo
        AtivaBotoes
        grid_oleo.SetFocus
    ElseIf xOperacao = "Excluir" Then
        txtData.Text = Format(lCxData, "dd/mm/yyyy")
        cbo_periodo.ListIndex = lCxPeriodo - 1
        cboIlha.ListIndex = lCxIlha - 1
        cboTipoSubEstoque.ListIndex = lCxSubEstoque - 1
        cboTipoMovimento.ListIndex = lCxTipoMovimento - 1
        txt_funcionario.Text = lCxCodigoFuncionario
        dtcboFuncionario.BoundText = lCxCodigoFuncionario
        lCxCodigoUsuario = Val(RetiraString(10, xString))
        AtualizaGrid
        If grid_oleo.Rows = 2 Then
            MsgBox "Não existe venda a ser excluída!", vbInformation, "Erro de Integridade!"
            cmd_cancelar_Click
            Exit Sub
        End If
        grid_oleo.Row = grid_oleo.Rows - 2
        MarcaCelulaOleo
        AtivaBotoes
        cmd_excluir.Enabled = True
        cmd_excluir.SetFocus
    End If
End Sub
Private Sub AdcionaDadosGrid()
    Dim x_i As Integer
'        If grid_oleo.Rows > 2 Then
'            For x_i = 2 To grid_oleo.Rows
'                grid_oleo.Row = x_i - 1
'                grid_oleo.Col = 2
'                If Val(grid_oleo.Text) = !Codigo Then
'                    x_flag = False
'                    Exit For
'                End If
'            Next
'        End If

    grid_oleo.Row = grid_oleo.Rows - 1
    grid_oleo.Col = 0
    grid_oleo.Text = lRS("Data").Value
    grid_oleo.Col = 1
    grid_oleo.Text = lRS("Periodo").Value
    grid_oleo.Col = 2
    grid_oleo.Text = lRS("Numero da Ilha").Value
    grid_oleo.Col = 3
    grid_oleo.Text = lRS("Codigo do Tipo do SubEstoque").Value
    grid_oleo.Col = 4
    grid_oleo.Text = lRS("Codigo do Funcionario").Value
    grid_oleo.Col = 5
    grid_oleo.Text = lRS("NomeFuncionario").Value
    grid_oleo.Col = 6
    grid_oleo.Text = lRS("NomeProduto").Value
    grid_oleo.Col = 7
    grid_oleo.Text = Format(lRS("Valor Total").Value, "###,##0.00")
    grid_oleo.Col = 8
    grid_oleo.Text = lRS("Ordem da Digitacao").Value
    grid_oleo.Col = 9
    grid_oleo.Text = lRS("Codigo do Produto2").Value
    grid_oleo.Col = 10
    grid_oleo.Text = lRS("Tipo do Movimento").Value
    grid_oleo.Rows = grid_oleo.Rows + 1
End Sub
Private Sub AlteraPrecoVendaCadastro(ByVal pCodigo As Long)
    On Error GoTo FileError
    Exit Sub

    If Estoque.PrecoVenda <> MovimentoLubrificante.ValorVenda Then
        Produto.PrecoVenda = MovimentoLubrificante.ValorVenda
        Estoque.PrecoVenda = MovimentoLubrificante.ValorVenda
        If Produto.Alterar(pCodigo, g_empresa) Then
            If Estoque.Alterar(g_empresa, pCodigo) Then
            Else
                MsgBox "Não foi possível alterar preço de venda no Estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível alterar preço de venda no Produto!", vbInformation, "Erro de Integridade!"
        End If
    End If

    Exit Sub
FileError:
    Conectar.Conexao.RollbackTrans
    MsgBox "Erro desconhecido na atualização de registro!", vbInformation, "Erro de Integridade"
    Exit Sub
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    If g_nivel_acesso > 4 Then
        cmd_novo.Enabled = False
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    If lCxOperacao = 2 Then
        cmd_novo.Enabled = False
        cmd_pesquisa.Enabled = False
        frm_move.Visible = False
        cmd_excluir.Enabled = False
        If cmd_alterar.Enabled = False Then
            If AberturaCaixa.LocalizarCodigo(g_empresa, lCxData, "NF", lCxPeriodo, lCxIlha, lCxCodigoFuncionario, lCxTipoMovimento) Then
                If AberturaCaixa.DataFechamento = "00:00:00" Then
                    cmd_alterar.Enabled = True
                End If
            End If
        End If
    ElseIf lCxOperacao = 3 Then
        cmd_novo.Enabled = False
        cmd_alterar.Enabled = False
        cmd_pesquisa.Enabled = False
        frm_move.Visible = False
        If cmd_excluir.Enabled = False Then
            If AberturaCaixa.LocalizarCodigo(g_empresa, lCxData, "NF", lCxPeriodo, lCxIlha, lCxCodigoFuncionario, lCxTipoMovimento) Then
                If AberturaCaixa.DataFechamento = "00:00:00" Then
                    cmd_excluir.Enabled = True
                End If
            End If
        End If
    End If
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xValor As Currency
    
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    xValor = 0
    xComplemento = "VENDA DE LUBRIFICANTES"
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xComplemento = "LUBRIFICANTES Per:" & Val(cbo_periodo.Text) & " Ilha:" & Val(cboIlha.Text) & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & Val(cboTipoMovimento.Text)
        
        'Caso Exista Deleta e Guarda o Valor
        If lCaixaIndividual Then
            If MovimentoCaixaPista.LocalizarRegistroEspecialUsu(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cboIlha.Text), xComplemento, IntegracaoCaixa.ContaCredito, "C", lCxCodigoUsuario) Then
                xValor = MovimentoCaixaPista.Valor
                If Not MovimentoCaixaPista.Excluir(g_empresa, CDate(txtData.Text), MovimentoCaixaPista.NumeroMovimento) Then
                    MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                End If
            End If
        Else
            If MovimentoCaixaPista.LocalizarRegistroEspecial(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cboIlha.Text), xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                xValor = MovimentoCaixaPista.Valor
                If Not MovimentoCaixaPista.Excluir(g_empresa, CDate(txtData.Text), MovimentoCaixaPista.NumeroMovimento) Then
                    MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                End If
            End If
        End If
'        xValor = xValor + fValidaValor(txt_valor_total.Text)
'        If lOpcao = 2 Then
'            xValor = MovimentoLubrificante.TotalPeriodo(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cboTipoMovimento.Text))
'            xValor = xValor - lValorAnterior
'            xValor = xValor + fValidaValor(txt_valor_total.Text)
'        End If
'  Teoricamente a linha abaixo substitui todo o cálculo do total.
        If lOpcao = 1 Then
            xValor = fValidaValor(lblTotalVenda.Caption) + fValidaValor(txt_valor_total.Text)
        ElseIf lOpcao = 2 Then
            xValor = fValidaValor(lblTotalVenda.Caption) - lValorAnterior + fValidaValor(txt_valor_total.Text)
        End If
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = CDate(txtData.Text)
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Valor = xValor
        MovimentoCaixaPista.NumeroDocumento = ""
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovimentoCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovimentoCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovimentoCaixaPista.TipoMovimento = lCxTipoMovimento
        MovimentoCaixaPista.CodigoUsuario = lCxCodigoUsuario
        MovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        'MovimentoCaixaPista.DadosInterno = "LUBRI" & "|@|" & Val(cboTipoMovimento.Text) & "|@|" & Val(cboTipoSubEstoque.Text) & "|@|"
        MovimentoCaixaPista.DadosInterno = "LUBRI" & "|@|" & Val(cboTipoSubEstoque.Text) & "|@|"
        MovimentoCaixaPista.CodigoLancamentoPadrao = lCxCodigoLancamentoPadrao
        If lOpcao = 1 Then
            MovimentoCaixaPista.DataDigitacao = Format(Date, "dd/MM/yyyy")
            MovimentoCaixaPista.HoraDigitacao = Format(Time, "HH:mm:ss")
            MovimentoCaixaPista.DataAlteracao = "00:00:00"
            MovimentoCaixaPista.HoraAlteracao = "00:00:00"
        Else
            MovimentoCaixaPista.DataDigitacao = lCxDataDigitacao
            MovimentoCaixaPista.HoraDigitacao = lCxHoraDigitacao
            MovimentoCaixaPista.DataAlteracao = Format(Date, "dd/MM/yyyy")
            MovimentoCaixaPista.HoraAlteracao = Format(Time, "HH:mm:ss")
        End If
        If MovimentoCaixaPista.Incluir Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovimentoCaixaPista.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & txt_valor_total.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "VENDA DE LUBRIFICANTES" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub AtualizaConstantes()
    Dim xDados As String
    lBaixaAutomaticaNoEstoque = False
    xDados = ReadINI("CUPOM FISCAL", "Baixa Automatica no Estoque", gArquivoIni)
    If xDados = "SIM" Then
        lBaixaAutomaticaNoEstoque = True
    End If
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
        lAlteraPrecoCadastro = Configuracao.AlteraPrecoProdutoPelaVenda
    Else
        lQtdPeriodo = 1
        lAlteraPrecoCadastro = False
    End If
End Sub
Private Sub AtualTabe()
    MovimentoLubrificante.Empresa = g_empresa
    MovimentoLubrificante.Data = CDate(txtData.Text)
    MovimentoLubrificante.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovimentoLubrificante.NumeroIlha = Val(cboIlha.Text)
    MovimentoLubrificante.CodigoTipoSubEstoque = cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    MovimentoLubrificante.CodigoFuncionario = CLng(dtcboFuncionario.BoundText)
    MovimentoLubrificante.CodigoProduto = CLng(dtcboProduto.BoundText)
    MovimentoLubrificante.Quantidade = fValidaValor(txt_quantidade.Text)
    MovimentoLubrificante.ValorCusto = lPrecoCusto
    MovimentoLubrificante.ValorVenda = fValidaValor(txt_valor_unitario.Text)
    MovimentoLubrificante.ValorTotal = fValidaValor(txt_valor_total.Text)
    MovimentoLubrificante.OrdemDigitacao = lOrdem
    MovimentoLubrificante.TipoMovimento = cboTipoMovimento.ItemData(cboTipoMovimento.ListIndex)
End Sub
Private Sub AtualTabeTranferenciaInternaEstoque()
    TransfInternaEstoque.Empresa = MovimentoLubrificante.Empresa
    TransfInternaEstoque.Data = MovimentoLubrificante.Data
    TransfInternaEstoque.Periodo = MovimentoLubrificante.Periodo
    TransfInternaEstoque.NumeroIlha = MovimentoLubrificante.NumeroIlha
    TransfInternaEstoque.CodigoSubEstoqueEntrada = MovimentoLubrificante.CodigoTipoSubEstoque
    TransfInternaEstoque.CodigoProduto = MovimentoLubrificante.CodigoProduto
    TransfInternaEstoque.CodigoFuncionario = MovimentoLubrificante.CodigoFuncionario
    TransfInternaEstoque.CodigoSubEstoqueSaida = 1
    TransfInternaEstoque.Quantidade = MovimentoLubrificante.Quantidade
    TransfInternaEstoque.Transferido = False
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = MovimentoLubrificante.Data
    lPeriodo = MovimentoLubrificante.Periodo
    lIlha = MovimentoLubrificante.NumeroIlha
    lTipoSubEstoque = MovimentoLubrificante.CodigoTipoSubEstoque
    lTipoMovimento = MovimentoLubrificante.TipoMovimento
    lCodigoProduto = MovimentoLubrificante.CodigoProduto
    lCodigoFuncionario = MovimentoLubrificante.CodigoFuncionario
    lOrdem = MovimentoLubrificante.OrdemDigitacao
    lPrecoCusto = MovimentoLubrificante.ValorCusto
    lQuantidade = MovimentoLubrificante.Quantidade
    lValorAnterior = MovimentoLubrificante.ValorTotal
    
    txtData.Text = Format(MovimentoLubrificante.Data, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovimentoLubrificante.Periodo - 1
    cboIlha.ListIndex = MovimentoLubrificante.NumeroIlha - 1
    cboTipoSubEstoque.ListIndex = MovimentoLubrificante.CodigoTipoSubEstoque - 1
    cboTipoMovimento.ListIndex = MovimentoLubrificante.TipoMovimento - 1
    dtcboFuncionario.BoundText = ""
    If Funcionario.LocalizarCodigo(g_empresa, MovimentoLubrificante.CodigoFuncionario) Then
        txt_funcionario.Text = MovimentoLubrificante.CodigoFuncionario
        dtcboFuncionario.BoundText = MovimentoLubrificante.CodigoFuncionario
    End If
    dtcboProduto.BoundText = ""
    If Produto.LocalizarCodigo(MovimentoLubrificante.CodigoProduto) Then
        txt_produto.Text = MovimentoLubrificante.CodigoProduto
        dtcboProduto.BoundText = MovimentoLubrificante.CodigoProduto
    End If
    txt_valor_unitario.Text = Format(MovimentoLubrificante.ValorVenda, "###,##0.0000")
    txt_quantidade.Text = Format(MovimentoLubrificante.Quantidade, "###,##0.00")
    txt_valor_total.Text = Format(MovimentoLubrificante.ValorTotal, "###,##0.00")
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    
    If MovimentoLubrificante.LocalizarUltimo(g_empresa) Then
        txtData.Text = Format(MovimentoLubrificante.Data, "dd/mm/yyyy")
        x_periodo = MovimentoLubrificante.Periodo
        If MovimentoLubrificante.Periodo >= lQtdPeriodo Then
            txtData.Text = Format(MovimentoLubrificante.Data + 1, "dd/mm/yyyy")
            x_periodo = 0
        End If
        cbo_periodo.ListIndex = x_periodo
        cboIlha.ListIndex = 0
        cboTipoSubEstoque.ListIndex = 0
        cboTipoMovimento.ListIndex = 0
        BuscaProximoCaixa = True
        Exit Function
    End If
    txtData.Text = Format(g_data_def - 1, "dd/mm/yyyy")
    cbo_periodo.ListIndex = 0
    cboIlha.ListIndex = 0
    cboTipoSubEstoque.ListIndex = 0
    cboTipoMovimento.ListIndex = 0
End Function
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
Private Function ExcluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xValor As Currency
    Dim xLocalizouRegistro As Boolean
    
    ExcluiMovimentoCaixa = True
    lNumeroMovimentoCaixa = 0
    xValor = 0
    xComplemento = "VENDA DE LUBRIFICANTES"
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xComplemento = "LUBRIFICANTES Per:" & Val(cbo_periodo.Text) & " Ilha:" & Val(cboIlha.Text) & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & Val(cboTipoMovimento.Text)
        xLocalizouRegistro = False
        If lCaixaIndividual Then
            If MovimentoCaixaPista.LocalizarRegistroEspecialUsu(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cboIlha.Text), xComplemento, IntegracaoCaixa.ContaCredito, "C", lCxCodigoUsuario) Then
                xLocalizouRegistro = True
            End If
        Else
            If MovimentoCaixaPista.LocalizarRegistroEspecial(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cboIlha.Text), xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                xLocalizouRegistro = True
            End If
        End If
        If xLocalizouRegistro Then
            xValor = MovimentoCaixaPista.Valor
            lCxDataDigitacao = MovimentoCaixaPista.DataDigitacao
            lCxHoraDigitacao = MovimentoCaixaPista.HoraDigitacao
            lNumeroMovimentoCaixa = MovimentoCaixaPista.NumeroMovimento
            If lOpcao = 2 Then
                'Alterar
                MovimentoCaixaPista.Valor = MovimentoCaixaPista.Valor - lValorAnterior
                If Not MovimentoCaixaPista.Alterar(g_empresa, MovimentoCaixaPista.Data, lNumeroMovimentoCaixa) Then
                    MsgBox "Não foi possível alterar o movimento do caixa!", vbInformation, "Erro de Integridade."
                    ExcluiMovimentoCaixa = False
                End If
            Else
                'Excluir
                If MovimentoCaixaPista.Valor = lValorAnterior Then
                    If Not MovimentoCaixaPista.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
                        MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                        ExcluiMovimentoCaixa = False
                    End If
                Else
                    MovimentoCaixaPista.Valor = MovimentoCaixaPista.Valor - lValorAnterior
                    If Not MovimentoCaixaPista.Alterar(g_empresa, MovimentoCaixaPista.Data, lNumeroMovimentoCaixa) Then
                        MsgBox "Não foi possível alterar o movimento do caixa!", vbInformation, "Erro de Integridade."
                        ExcluiMovimentoCaixa = False
                    End If
                End If
            End If
        End If
    Else
        ExcluiMovimentoCaixa = False
    End If
    
    
'    If MovimentoCaixaPista.LocalizarCodigo(g_empresa, lData, lNumeroMovimentoCaixa) Then
'        lCxDataDigitacao = MovimentoCaixaPista.DataDigitacao
'        lCxHoraDigitacao = MovimentoCaixaPista.HoraDigitacao
'    Else
'        MsgBox "Não foi possível localizar o movimento do caixa!", vbInformation, "Erro de Integridade."
'    End If
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set AberturaCaixa = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set Estoque = Nothing
    Set Funcionario = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovimentoLubrificante = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
    Set TransfInternaEstoque = Nothing
End Sub
Private Sub AtualizaGrid()
    Dim xSQL As String
    Dim xTotal As Currency
    
    MontaGrid
    xTotal = 0
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Lubrificante.Data, Movimento_Lubrificante.Periodo, Movimento_Lubrificante.[Codigo do Funcionario], Movimento_Lubrificante.[Codigo do Produto2], Movimento_Lubrificante.Quantidade, Movimento_Lubrificante.[Valor Custo], Movimento_Lubrificante.[Valor Venda], Movimento_Lubrificante.[Valor Total], Movimento_Lubrificante.[Ordem da Digitacao], Movimento_Lubrificante.[Numero da Ilha], Movimento_Lubrificante.[Codigo do Tipo do SubEstoque], Funcionario.Nome as NomeFuncionario, Produto.Nome as NomeProduto, Movimento_Lubrificante.[Tipo do Movimento]"
    xSQL = xSQL & "  FROM Movimento_Lubrificante, Funcionario, Produto"
    xSQL = xSQL & " WHERE Movimento_Lubrificante.Empresa = " & g_empresa
    If IsDate(txtData.Text) Then
        xSQL = xSQL & "   AND Movimento_Lubrificante.Data = " & preparaData(CDate(txtData.Text))
        xSQL = xSQL & "   AND Movimento_Lubrificante.Periodo = " & preparaTexto(cbo_periodo.Text)
        xSQL = xSQL & "   AND Movimento_Lubrificante.[Numero da Ilha] = " & Val(cboIlha.Text)
        xSQL = xSQL & "   AND Movimento_Lubrificante.[Codigo do Tipo do SubEstoque] = " & cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
        xSQL = xSQL & "   AND Movimento_Lubrificante.[Tipo do Movimento] = " & cboTipoMovimento.ItemData(cboTipoMovimento.ListIndex)
        'xSQL = xSQL & "   AND Movimento_Lubrificante.[Codigo do Funcionario] = " & CLng(dtcboFuncionario.BoundText)
    End If
    If lCaixaIndividual Then
        xSQL = xSQL & "   AND Movimento_Lubrificante.[Codigo do Funcionario] = " & lCxCodigoFuncionario
    End If
    xSQL = xSQL & "   AND Funcionario.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Funcionario.Codigo = Movimento_Lubrificante.[Codigo do Funcionario]"
    xSQL = xSQL & "   AND Produto.Codigo = Movimento_Lubrificante.[Codigo do Produto2]"
    xSQL = xSQL & " ORDER BY Movimento_Lubrificante.[Codigo do Produto2] ASC"
    Set lRS = New adodb.Recordset
    Set lRS = Conectar.RsConexao(xSQL)
    'Set lRS = MovimentoLubrificante.MondaRS(xSQL)
    If lRS.RecordCount > 0 Then
        lRS.MoveFirst
        Do Until lRS.EOF
            AdcionaDadosGrid
            xTotal = xTotal + lRS("Valor Total").Value
            lRS.MoveNext
        Loop
    End If
    lblTotalVenda.Caption = Format(xTotal, "###,##0.00")
    'MontaGrid
    'If MovimentoLubrificante.LocalizarPrimeiroDataPer(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque) Then
    '    AdcionaDadosGrid
    '    Do Until MovimentoLubrificante.LocalizarProximo = False
    '        If MovimentoLubrificante.Data <> lData Or MovimentoLubrificante.Periodo <> lPeriodo Or MovimentoLubrificante.NumeroIlha <> lIlha Or MovimentoLubrificante.CodigoTipoSubEstoque <> lTipoSubEstoque Then
    '            Exit Do
    '        End If
    '        AdcionaDadosGrid
    '    Loop
    'End If
    'If MovimentoLubrificante.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
    '    AtualTela
    'End If
    grid_oleo.Row = grid_oleo.Rows - 1
    grid_oleo.Col = 0
End Sub
Private Sub PreencheCboIlha()
    Dim i As Integer
    
    cboIlha.Clear
    If Configuracao.LocalizarCodigo(g_empresa) Then
        For i = 1 To Configuracao.QuantidadeIlha
            cboIlha.AddItem i
            cboIlha.ItemData(cboIlha.NewIndex) = i
        Next
    End If
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
    Dim rstTipoMovimento As New adodb.Recordset
    
    cboTipoMovimento.Clear
    Set rstTipoMovimento = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoMovimentoCaixa ORDER BY Codigo")
    Do Until rstTipoMovimento.EOF
        cboTipoMovimento.AddItem rstTipoMovimento!Codigo & " " & rstTipoMovimento!Nome
        cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = rstTipoMovimento!Codigo
        rstTipoMovimento.MoveNext
    Loop
    rstTipoMovimento.Close
    Set rstTipoMovimento = Nothing
End Sub
Private Sub PreencheCboTipoSubEstoque()
    Dim rstTipoSubEstoque As New adodb.Recordset
    
    cboTipoSubEstoque.Clear
    Set rstTipoSubEstoque = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoSubEstoque ORDER BY Codigo")
    Do Until rstTipoSubEstoque.EOF
        cboTipoSubEstoque.AddItem rstTipoSubEstoque!Codigo & " " & rstTipoSubEstoque!Nome
        cboTipoSubEstoque.ItemData(cboTipoSubEstoque.NewIndex) = rstTipoSubEstoque!Codigo
        rstTipoSubEstoque.MoveNext
    Loop
    rstTipoSubEstoque.Close
    Set rstTipoSubEstoque = Nothing
End Sub
Private Sub btnTransfereVendaECF_Click()
    Dim xTipoVenda As String
    
    If IsDate(Me.txtData.Text) Then
        xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
        If xTipoVenda = "CONVENIENCIA" Then
            TransfereDadosConveniencia
        Else
            TransfereDadosECF
        End If
    
        If MovimentoLubrificante.LocalizarUltimo(g_empresa) Then
            AtualTela
        End If
        cmd_cancelar_Click
    End If
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlha.SetFocus
    End If
End Sub
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoSubEstoque.SetFocus
    End If
End Sub
Private Sub cboTipoSubEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoMovimento.SetFocus
    End If
End Sub
Private Sub cboTipoMovimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_quantidade.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovimentoLubrificante.LocalizarAnterior Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
    btnTransfereVendaECF.Enabled = False
    btnTransfereVendaECF.Visible = False
    LimpaTela
    If MovimentoLubrificante.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
        AtualTela
        AtivaBotoes
        AtualizaGrid
        If cmd_alterar.Enabled Then
            cmd_alterar.SetFocus
        Else
            cmd_novo.SetFocus
        End If
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Sub LimpaGrid()
    Do Until grid_oleo.Rows = 2
        grid_oleo.Row = grid_oleo.Rows - 1
        grid_oleo.RemoveItem grid_oleo.Row
    Loop
    grid_oleo.Row = 1
    grid_oleo.Col = 0
    grid_oleo.Text = ""
    grid_oleo.Col = 1
    grid_oleo.Text = ""
    grid_oleo.Col = 2
    grid_oleo.Text = ""
    grid_oleo.Col = 3
    grid_oleo.Text = ""
    grid_oleo.Col = 4
    grid_oleo.Text = ""
    grid_oleo.Col = 5
    grid_oleo.Text = ""
    grid_oleo.Col = 6
    grid_oleo.Text = ""
    grid_oleo.Col = 7
    grid_oleo.Text = ""
    grid_oleo.Col = 8
    grid_oleo.Text = ""
    grid_oleo.Col = 9
    grid_oleo.Text = ""
    grid_oleo.Col = 10
    grid_oleo.Text = ""
End Sub
Private Sub LimpaTela()
    If lGravados = 0 Then
        txtData.Text = ""
        cbo_periodo.ListIndex = -1
        cboIlha.ListIndex = -1
        cboTipoSubEstoque.ListIndex = -1
        cboTipoMovimento.ListIndex = -1
        txt_funcionario.Text = ""
        dtcboFuncionario.BoundText = ""
    End If
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    txt_valor_unitario.Text = ""
    txt_quantidade.Text = ""
    txt_valor_total.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    Dim xAtualizou As Boolean
    
    Call GravaAuditoria(1, Me.name, 4, "")
    If fValidaValor(txt_quantidade.Text) > 0 And fValidaValor(txt_valor_total.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & Val(cboIlha.Text) & " Vlr:" & txt_valor_total.Text & " Cod.Prod:" & txt_produto.Text)
            lOpcao = 3
            Conectar.IniciaTransacao
            xAtualizou = True
            
            'Caso a TransferencicInternaEstoque não exista
            'Será considerada como já transferida pois assim
            'Não gera erro ao tentar excluí-la
            If Not TransfInternaEstoque.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                TransfInternaEstoque.Transferido = True
            End If
            
            If MovimentoLubrificante.Excluir(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, True) Then
                    If ExcluiMovimentoCaixa Then
                        If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, lTipoSubEstoque, lQuantidade, True) Then
                            If TransfInternaEstoque.Transferido = False Then
                                If TransfInternaEstoque.Excluir(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                                    Conectar.ConfirmaTransacao
                                Else
                                    Conectar.CancelaTransacao
                                    MsgBox "Não foi possível excluir o registro de transferência!", vbInformation, "Erro de Integridade!"
                                End If
                            Else
                                Conectar.ConfirmaTransacao
                            End If
                        Else
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        MsgBox "Não foi possível excluir o movimento no caixa!", vbInformation, "Erro de Integridade!"
                        Conectar.CancelaTransacao
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                End If
            Else
                Conectar.CancelaTransacao
                MsgBox "Não foi possível excluir o registro!", vbInformation, "Erro de Integridade!"
            End If
            If lCxPeriodo > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
            If MovimentoLubrificante.LocalizarUltimo(g_empresa) Then
                AtualTela
            Else
                LimpaTela
                lGravados = 0
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
            lOpcao = 0
            AtualizaGrid
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
'
'    zzCalculaSaidaParaEstoque
'    ZZLeCupomGravarConveniencia
'    ZZLeConvenienciaGravarLubrificante
'    Exit Sub
'
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    btnTransfereVendaECF.Enabled = True
    btnTransfereVendaECF.Visible = True
    frmDados.Enabled = True
    If lGravados = 0 Then
        If lCxPeriodo > 0 Then
            txtData.Text = Format(lCxData, "dd/mm/yyyy")
            cbo_periodo.ListIndex = lCxPeriodo - 1
            cboIlha.ListIndex = lCxIlha - 1
            cboTipoSubEstoque.ListIndex = lCxSubEstoque - 1
            cboTipoMovimento.ListIndex = lCxTipoMovimento - 1
            txt_funcionario.Text = lCxCodigoFuncionario
            dtcboFuncionario.BoundText = lCxCodigoFuncionario
            txt_produto.SetFocus
            Exit Sub
        End If
        If BuscaProximoCaixa Then
            txt_funcionario.SetFocus
        Else
            txtData.SetFocus
        End If
    Else
        txt_produto.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    Dim xAtualizou As Boolean
    On Error GoTo FileError
    
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                AtualTabe
                Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & Val(cboIlha.Text) & " S.Est:" & Val(cboTipoSubEstoque.Text) & " Vlr:" & txt_valor_total.Text & " Cod.Prod:" & txt_produto.Text)
                Conectar.IniciaTransacao
                If IncluiMovimentoCaixa Then
                    If MovimentoLubrificante.Incluir Then
                        'If lAlteraPrecoCadastro Then
                        '    AlteraPrecoVendaCadastro (MovimentoLubrificante.CodigoProduto)
                        'End If
                        lData = MovimentoLubrificante.Data
                        lPeriodo = MovimentoLubrificante.Periodo
                        lIlha = MovimentoLubrificante.NumeroIlha
                        lTipoSubEstoque = MovimentoLubrificante.CodigoTipoSubEstoque
                        lTipoMovimento = MovimentoLubrificante.TipoMovimento
                        lCodigoProduto = MovimentoLubrificante.CodigoProduto
                        lCodigoFuncionario = MovimentoLubrificante.CodigoFuncionario
                        lQuantidade = MovimentoLubrificante.Quantidade
                        lOrdem = MovimentoLubrificante.OrdemDigitacao
                        If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, False) Then
                            lGravados = 1
                            AtualTabeTranferenciaInternaEstoque
                            If TransfInternaEstoque.Incluir Then
                                If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, lTipoSubEstoque, lQuantidade, False) Then
                                    Conectar.ConfirmaTransacao
                                Else
                                    Conectar.CancelaTransacao
                                    MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                                End If
                            Else
                                Conectar.CancelaTransacao
                                MsgBox "Não foi possível incluir o registro de transferência!", vbInformation, "Erro de Integridade!"
                            End If
                        Else
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível incluir o registro!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Erro ao incluir registro!" & Chr(10) & "Não foi possível integrar com o Caixa!", vbCritical, "Erro na Inclusão."
                End If
            ElseIf lOpcao = 2 Then
                xAtualizou = False
                Conectar.IniciaTransacao
                If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, True) Then
                    If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, lTipoSubEstoque, lQuantidade, True) Then
                        xAtualizou = True
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                End If
                If xAtualizou Then
                    AtualTabe
                    If ExcluiMovimentoCaixa Then
                        If IncluiMovimentoCaixa Then
                            Call GravaAuditoria(1, Me.name, 10, "De: Dt:" & lData & " Per:" & lPeriodo & " Ilha:" & lIlha & " Vlr:" & lValorAnterior & " Cod.Prod:" & lCodigoProduto)
                            Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & Val(cboIlha.Text) & " Vlr:" & txt_valor_total.Text & " Cod.Prod:" & txt_produto.Text)
                            If MovimentoLubrificante.Alterar(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                                If TransfInternaEstoque.Transferido = False Then
                                    If TransfInternaEstoque.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                                        AtualTabeTranferenciaInternaEstoque
                                        If Not TransfInternaEstoque.Alterar(g_empresa, lData, lPeriodo, lIlha, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                                            Conectar.CancelaTransacao
                                            xAtualizou = False
                                            MsgBox "Não foi possível alterar o registro de transferência!", vbInformation, "Erro de Integridade!"
                                        End If
                                    Else
                                        AtualTabeTranferenciaInternaEstoque
                                        If Not TransfInternaEstoque.Incluir Then
                                            Conectar.CancelaTransacao
                                            xAtualizou = False
                                            MsgBox "Não foi possível incluir o registro de transferência!", vbInformation, "Erro de Integridade!"
                                        End If
                                    End If
                                End If
                                If xAtualizou Then
                                    lData = MovimentoLubrificante.Data
                                    lPeriodo = MovimentoLubrificante.Periodo
                                    lIlha = MovimentoLubrificante.NumeroIlha
                                    lTipoSubEstoque = MovimentoLubrificante.CodigoTipoSubEstoque
                                    lTipoMovimento = MovimentoLubrificante.TipoMovimento
                                    lCodigoProduto = MovimentoLubrificante.CodigoProduto
                                    lCodigoFuncionario = MovimentoLubrificante.CodigoFuncionario
                                    lQuantidade = MovimentoLubrificante.Quantidade
                                    lOrdem = MovimentoLubrificante.OrdemDigitacao
                                    If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, False) Then
                                        If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, lTipoSubEstoque, lQuantidade, False) Then
                                            Conectar.ConfirmaTransacao
                                        Else
                                            Conectar.CancelaTransacao
                                            MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                                        End If
                                    Else
                                        Conectar.CancelaTransacao
                                        MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                                    End If
                                End If
                            Else
                                Conectar.CancelaTransacao
                                MsgBox "Não foi possível alterar o registro!", vbInformation, "Erro de Integridade!"
                            End If
                        Else
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                        End If
                    Else
                        Conectar.CancelaTransacao
                    End If
                End If
            End If
            'AtualizaGrid
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
                AtualTela
                AtualizaGrid
            Else
                MsgBox "Não foi possível localizar o registro!", vbInformation, "Erro de Integridade!"
            End If
            If lOpcao = 1 Then
                lOpcao = 0
                cmd_novo_Click
            Else
                lOpcao = 0
                cmd_alterar.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    Conectar.CancelaTransacao
    MsgBox "Erro desconhecido na atualização de registro!", vbInformation, "Erro de Integridade"
    Exit Sub
End Sub
Private Sub cmd_ok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 18 Then  'Crtl + R
        KeyAscii = 0
        'RecalculaMovimentoCaixa
    End If
End Sub
Private Sub ZZLeCupomGravarConveniencia()
    Dim xOrdem As Integer
    Dim xSQL As String
    Dim MovimentoVendaConveniencia As New cMovimentoVendaConveniencia
    
    On Error GoTo FileError
    
    Exit Sub
    xSQL = ""
    xSQL = xSQL & "SELECT Empresa, [Numero do Cupom], Ordem, Data, Hora, [Data do Cupom],"
    xSQL = xSQL & "       Periodo, [Codigo do Produto], [Valor Unitario], Quantidade,"
    xSQL = xSQL & "       [Valor Total], [Forma de Pagamento], [Valor Recebido], Operador,"
    xSQL = xSQL & "       [Cupom Cancelado], [Item Cancelado], [Codigo da Aliquota],"
    xSQL = xSQL & "       [Valor do Desconto], [Codigo do Cliente]"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & " WHERE [Codigo da ECF] = 3"
    xSQL = xSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & " ORDER BY Data"
    Set lRS = New adodb.Recordset
    Set lRS = Conectar.RsConexao(xSQL)
    
    If lRS.RecordCount > 0 Then
        lRS.MoveFirst
        Do Until lRS.EOF
            MovimentoVendaConveniencia.Empresa = g_empresa
            MovimentoVendaConveniencia.NumeroCupom = lRS("Numero do Cupom").Value
            MovimentoVendaConveniencia.Ordem = lRS("Ordem").Value
            MovimentoVendaConveniencia.Data = lRS("Data").Value
            MovimentoVendaConveniencia.Hora = lRS("Hora").Value
            MovimentoVendaConveniencia.DataCupom = lRS("Data do Cupom").Value
            MovimentoVendaConveniencia.Periodo = lRS("Periodo").Value
            MovimentoVendaConveniencia.TipoMovimento = 1
            MovimentoVendaConveniencia.CodigoProduto = lRS("Codigo do Produto").Value
            MovimentoVendaConveniencia.ValorUnitario = lRS("Valor Unitario").Value
            MovimentoVendaConveniencia.Quantidade = lRS("Quantidade").Value
            MovimentoVendaConveniencia.ValorTotal = lRS("Valor Total").Value
            MovimentoVendaConveniencia.FormaPagamento = lRS("Forma de Pagamento").Value
            MovimentoVendaConveniencia.ValorRecebido = lRS("Valor Recebido").Value
            MovimentoVendaConveniencia.operador = lRS("Operador").Value
            MovimentoVendaConveniencia.CupomCancelado = lRS("Cupom Cancelado").Value
            MovimentoVendaConveniencia.ItemCancelado = lRS("Item Cancelado").Value
            MovimentoVendaConveniencia.CodigoAliquota = lRS("Codigo da Aliquota").Value
            MovimentoVendaConveniencia.ValorDesconto = lRS("Valor do Desconto").Value
            MovimentoVendaConveniencia.NumeroJustificativa = 0
            MovimentoVendaConveniencia.CodigoCliente = lRS("Codigo do Cliente").Value
            If Produto.LocalizarCodigo(lRS("Codigo do produto").Value) Then
                MovimentoVendaConveniencia.CodigoGrupo = Produto.CodigoGrupo
            Else
                MovimentoVendaConveniencia.CodigoGrupo = 0
            End If
            If Not MovimentoVendaConveniencia.Incluir Then
                MsgBox "Não foi possível incluir o registro Venda de Conveniencia!", vbInformation, "Erro de Integridade!"
            End If
            lRS.MoveNext
        Loop
    End If
    MsgBox "Processamento Concluido!"
    Exit Sub
    
FileError:
    MsgBox "Erro: Venda Conveniencia", vbInformation, "Erro nao identificado!"
End Sub
Private Sub ZZLeConvenienciaGravarLubrificante()
    Dim xOrdem As Integer
    Dim xSQL As String
    Dim xData As Date
    
    On Error GoTo FileError
    
    Exit Sub
    xSQL = ""
    xSQL = xSQL & "SELECT Empresa, Data, Periodo, "
    xSQL = xSQL & "       Operador, [Codigo do Produto], "
    xSQL = xSQL & "       Quantidade, [Valor Unitario], "
    xSQL = xSQL & "       [Valor Total], [Tipo do Movimento]"
    xSQL = xSQL & "  FROM Movimento_Venda_Conveniencia"
    xSQL = xSQL & " WHERE [Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & " ORDER BY Data ASC, Periodo ASC, Operador ASC"
    Set lRS = New adodb.Recordset
    Set lRS = Conectar.RsConexao(xSQL)
    
    xOrdem = 50
    xData = CDate("01/01/1900")
    If lRS.RecordCount > 0 Then
        lRS.MoveFirst
        Do Until lRS.EOF
            If xData <> lRS("Data").Value Then
                xOrdem = 50
            End If
            xOrdem = xOrdem + 1
            MovimentoLubrificante.Empresa = lRS("Empresa").Value
            MovimentoLubrificante.Data = lRS("Data").Value
            MovimentoLubrificante.Periodo = lRS("Periodo").Value
            MovimentoLubrificante.NumeroIlha = 1
            MovimentoLubrificante.CodigoTipoSubEstoque = 2 'lRS("Tipo do Movimento").Value
            MovimentoLubrificante.CodigoFuncionario = lRS("Operador").Value
            MovimentoLubrificante.CodigoProduto = lRS("Codigo do Produto").Value
            MovimentoLubrificante.Quantidade = lRS("Quantidade").Value
            If Produto.LocalizarCodigo(lRS("Codigo do produto").Value) Then
                MovimentoLubrificante.ValorCusto = Produto.PrecoCusto
            Else
                MovimentoLubrificante.ValorCusto = lRS("Valor Unitario").Value
            End If

            MovimentoLubrificante.ValorVenda = lRS("Valor Unitario").Value
            MovimentoLubrificante.ValorTotal = lRS("Valor Total").Value
            MovimentoLubrificante.OrdemDigitacao = xOrdem
            MovimentoLubrificante.TipoMovimento = 1
            
            If MovimentoLubrificante.LocalizarCodigo(lRS("Empresa").Value, CDate(lRS("Data").Value), lRS("Periodo").Value, 1, 1, 2, CLng(lRS("Codigo do Produto").Value), Val(lRS("Operador").Value)) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + lRS("Quantidade").Value
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + lRS("Valor Total").Value
                If MovimentoLubrificante.Alterar(g_empresa, CDate(lRS("Data").Value), lRS("Periodo").Value, 1, 1, 2, CLng(lRS("Codigo do Produto").Value), Val(lRS("Operador").Value)) Then
                Else
                    MsgBox "Não foi possível alterar o registro!", vbInformation, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Incluir Then
                Else
                    MsgBox "Não foi possível incluir o registro!", vbInformation, "Erro de Integridade!"
                End If
            End If
            lRS.MoveNext
        Loop
    End If
    MsgBox "Processamento Concluido!"
    Exit Sub
    
FileError:
    MsgBox "Erro: Venda Lubrificante", vbInformation, "Erro nao identificado!"
End Sub
Private Sub TransfereDadosConveniencia()
    Dim xOrdem As Integer
    Dim xSQL As String
    
    On Error Resume Next
    'On Error GoTo FileError
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferida a venda de Conveniência na data " & txtData.Text & Chr(10) & "Do período " & cbo_periodo.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Venda de Conveniência!")) = vbNo Then
        Exit Sub
    End If
    
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Venda_Conveniencia.Data, Movimento_Venda_Conveniencia.Periodo, "
    xSQL = xSQL & "       Movimento_Venda_Conveniencia.Operador, Movimento_Venda_Conveniencia.[Codigo do Produto], "
    xSQL = xSQL & "       Movimento_Venda_Conveniencia.Quantidade, Movimento_Venda_Conveniencia.[Valor Unitario], "
    xSQL = xSQL & "       Movimento_Venda_Conveniencia.[Valor Total], Produto.[Preco de Custo], "
    xSQL = xSQL & "       Movimento_Venda_Conveniencia.[Tipo do Movimento]"
    xSQL = xSQL & "  FROM Movimento_Venda_Conveniencia, Produto"
    xSQL = xSQL & " WHERE Movimento_Venda_Conveniencia.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Movimento_Venda_Conveniencia.Data >= " & preparaData(CDate(txtData.Text))
    xSQL = xSQL & "   AND Movimento_Venda_Conveniencia.Periodo >= " & preparaTexto(cbo_periodo.Text)
    xSQL = xSQL & "   AND Movimento_Venda_Conveniencia.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Venda_Conveniencia.[Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Produto.Codigo = Movimento_Venda_Conveniencia.[Codigo do Produto]"
    xSQL = xSQL & " ORDER BY Movimento_Venda_Conveniencia.Operador ASC, Movimento_Venda_Conveniencia.[Codigo do Produto] ASC"
    Set lRS = New adodb.Recordset
    Set lRS = Conectar.RsConexao(xSQL)
    
    xOrdem = 0
    If lRS.RecordCount > 0 Then
        lRS.MoveFirst
        Do Until lRS.EOF
            xOrdem = xOrdem + 1
            MovimentoLubrificante.Empresa = g_empresa
            MovimentoLubrificante.Data = lRS("Data").Value
            MovimentoLubrificante.Periodo = lRS("Periodo").Value
            MovimentoLubrificante.NumeroIlha = 1
            MovimentoLubrificante.CodigoTipoSubEstoque = 2 'lRS("Tipo do Movimento").Value
            'If Funcionario.LocalizarCodigo(g_empresa, lRS!operador) Then
            '    If UCase(Funcionario.Cargo) Like "*TROCADOR*" Then
            '        MovimentoLubrificante.CodigoTipoSubEstoque = 3
            '    End If
            'End If
            MovimentoLubrificante.CodigoFuncionario = lRS("Operador").Value
            MovimentoLubrificante.CodigoProduto = lRS("Codigo do Produto").Value
            MovimentoLubrificante.Quantidade = lRS("Quantidade").Value
            MovimentoLubrificante.ValorCusto = lRS("Preco de Custo").Value
            MovimentoLubrificante.ValorVenda = lRS("Valor Unitario").Value
            MovimentoLubrificante.ValorTotal = lRS("Valor Total").Value
            MovimentoLubrificante.OrdemDigitacao = xOrdem
            
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, CDate(lRS("Data").Value), lRS("Periodo").Value, 1, 1, 2, CLng(lRS("Codigo do Produto").Value), Val(lRS("Operador").Value)) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + lRS("Quantidade").Value
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + lRS("Valor Total").Value
                If MovimentoLubrificante.Alterar(g_empresa, CDate(lRS("Data").Value), lRS("Periodo").Value, 1, 1, 2, CLng(lRS("Codigo do Produto").Value), Val(lRS("Operador").Value)) Then
                Else
                    MsgBox "Não foi possível alterar o registro!", vbInformation, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Incluir Then
                Else
                    MsgBox "Não foi possível incluir o registro!", vbInformation, "Erro de Integridade!"
                End If
            End If
            lRS.MoveNext
        Loop
    End If
    MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a venda de Conveniência transferida!", vbInformation, "Transferência Concluida!"
    Exit Sub
    
FileError:
    MsgBox "Erro: TransfereDadosConveniencia", vbInformation, "Erro de Integridade!"
End Sub
Private Sub TransfereDadosECF()
    Dim xOrdem As Integer
    Dim xSQL As String
    
    On Error GoTo FileError
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferida a venda do ECF na data " & txtData.Text & Chr(10) & "Do período " & cbo_periodo.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Venda do ECF!")) = vbNo Then
        Exit Sub
    End If
    
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Cupom_Fiscal.Data, Movimento_Cupom_Fiscal.Periodo, "
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Operador, Movimento_Cupom_Fiscal.[Codigo do Produto], "
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Quantidade, Movimento_Cupom_Fiscal.[Valor Unitario], "
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Total], Produto.[Preco de Custo], "
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo do Movimento]"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal, Produto"
    xSQL = xSQL & " WHERE Movimento_Cupom_Fiscal.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Data = " & preparaData(CDate(txtData.Text))
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Periodo = " & preparaTexto(cbo_periodo.Text)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Produto.Codigo = Movimento_Cupom_Fiscal.[Codigo do Produto]"
    xSQL = xSQL & "   AND Produto.[Codigo do Grupo] <> " & 4
    xSQL = xSQL & " ORDER BY Movimento_Cupom_Fiscal.Operador ASC, Movimento_Cupom_Fiscal.[Codigo do Produto] ASC"
    Set lRS = New adodb.Recordset
    Set lRS = Conectar.RsConexao(xSQL)
    
    xOrdem = 0
    If lRS.RecordCount > 0 Then
        lRS.MoveFirst
        Do Until lRS.EOF
            xOrdem = xOrdem + 1
            MovimentoLubrificante.Empresa = g_empresa
            MovimentoLubrificante.Data = lRS("Data").Value
            MovimentoLubrificante.Periodo = lRS("Periodo").Value
            MovimentoLubrificante.NumeroIlha = 1
            MovimentoLubrificante.CodigoTipoSubEstoque = lRS("Tipo do Movimento").Value
            'If Funcionario.LocalizarCodigo(g_empresa, lRS!operador) Then
            '    If UCase(Funcionario.Cargo) Like "*TROCADOR*" Then
            '        MovimentoLubrificante.CodigoTipoSubEstoque = 3
            '    End If
            'End If
            MovimentoLubrificante.CodigoFuncionario = lRS("Operador").Value
            MovimentoLubrificante.CodigoProduto = lRS("Codigo do Produto").Value
            MovimentoLubrificante.Quantidade = lRS("Quantidade").Value
            MovimentoLubrificante.ValorCusto = lRS("Preco de Custo").Value
            MovimentoLubrificante.ValorVenda = lRS("Valor Unitario").Value
            MovimentoLubrificante.ValorTotal = lRS("Valor Total").Value
            MovimentoLubrificante.OrdemDigitacao = xOrdem
            
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, lRS("Data").Value, lRS("Periodo").Value, 1, 2, lRS("Tipo do Movimento").Value, lRS("Codigo do Produto").Value, lRS("Operador").Value) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + lRS("Quantidade").Value
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + lRS("Valor Total").Value
                If MovimentoLubrificante.Alterar(g_empresa, lRS("Data").Value, lRS("Periodo").Value, 1, 2, lRS("Tipo do Movimento").Value, lRS("Codigo do Produto").Value, lRS("Operador").Value) Then
                    'If lBaixaAutomaticaNoEstoque = False Then
                    '    If Not Estoque.AlterarQuantidade(g_empresa, lRS("Codigo do Produto").Value, lRS("Quantidade").Value, False) Then
                    '        MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                    '    End If
                    'End If
                    If TransfInternaEstoque.LocalizarCodigo(g_empresa, lRS("Data").Value, lRS("Periodo").Value, 1, lRS("Tipo do Movimento").Value, lRS("Codigo do Produto").Value, lRS("Operador").Value) Then
                        TransfInternaEstoque.Quantidade = TransfInternaEstoque.Quantidade + lRS("Quantidade").Value
                        If TransfInternaEstoque.Alterar(g_empresa, lRS("Data").Value, lRS("Periodo").Value, 1, lRS("Tipo do Movimento").Value, lRS("Codigo do Produto").Value, lRS("Operador").Value) Then
                            'If Not SubEstoque.AlterarQuantidade(g_empresa, lRS("Codigo do Produto").Value, lRS("Tipo do Movimento").Value, lRS("Quantidade").Value, False) Then
                            '    MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                            'End If
                        Else
                            MsgBox "Não foi possível alterar o registro de transferência!", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        MsgBox "Não foi localizar o registro de transferência!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    MsgBox "Não foi possível alterar o registro!", vbInformation, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Incluir Then
                    'If lBaixaAutomaticaNoEstoque = False Then
                    '    If Not Estoque.AlterarQuantidade(g_empresa, MovimentoLubrificante.CodigoProduto, MovimentoLubrificante.Quantidade, False) Then
                    '        MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                    '    End If
                    'End If
                    AtualTabeTranferenciaInternaEstoque
                    If TransfInternaEstoque.Incluir Then
                        'If Not SubEstoque.AlterarQuantidade(g_empresa, MovimentoLubrificante.CodigoProduto, lRS("Tipo do Movimento").Value, MovimentoLubrificante.Quantidade, False) Then
                        '    MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                        'End If
                    Else
                        MsgBox "Não foi possível incluir o registro de transferência!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    MsgBox "Não foi possível incluir o registro!", vbInformation, "Erro de Integridade!"
                End If
            End If
            lRS.MoveNext
        Loop
    End If
    MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a venda de ECF transferida!", vbInformation, "Transferência Concluida!"
    Exit Sub
    
FileError:
    MsgBox "Erro: TransfereDadosECF", vbInformation, "Erro de Integridade!"
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a data.", vbInformation, "Atenção!"
        txtData.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Selecione uma Ilha.", vbInformation, "Atenção!"
        cboIlha.SetFocus
    ElseIf cboTipoSubEstoque.ListIndex = -1 Then
        MsgBox "Escolha um tipo de sub-estoque.", vbInformation, "Atenção!"
        cboTipoSubEstoque.SetFocus
    ElseIf cboTipoMovimento.ListIndex = -1 Then
        MsgBox "Escolha um tipo de movimento.", vbInformation, "Atenção!"
        cboTipoMovimento.SetFocus
    ElseIf dtcboFuncionario.BoundText = "" Then
        MsgBox "Escolha o funcionario.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf Not fValidaValor(txt_valor_unitario.Text) > 0 Then
        MsgBox "Informe o valor unitário do produto.", vbInformation, "Atenção!"
        txt_valor_unitario.SetFocus
    ElseIf Not fValidaValor(txt_quantidade.Text) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf fValidaValor(txt_quantidade.Text) > 1000 And g_nome_empresa <> "*LANCHONETE BOM SUCESSO*" Then
        MsgBox "Quantidade acima de 1.000 não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor(txt_valor_total.Text) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
    ElseIf ValidaInclusao = False Then
        txt_produto.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Function ValidaInclusao() As Boolean
    ValidaInclusao = True
    If lOpcao = 1 Then
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, CDate(txtData.Text), cbo_periodo.Text, Val(cboIlha.Text), Val(cboTipoMovimento.ItemData(cboTipoMovimento.ListIndex)), Val(cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)), CLng(txt_produto.Text), Val(dtcboFuncionario.BoundText)) Then
            MsgBox "O produto " & Produto.Nome & Chr(10) & "Já tem venda neste caixa." & Chr(10) & "Por este motivo esta alteração não será aceita.", vbInformation, "Atenção! Procedimento não aceito."
            ValidaInclusao = False
        End If
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovimentoLubrificante.Empresa < g_cfg_empresa_i Or MovimentoLubrificante.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovimentoLubrificante.Data < g_cfg_data_i Or MovimentoLubrificante.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovimentoLubrificante.Periodo < g_cfg_periodo_i Or MovimentoLubrificante.Periodo > g_cfg_periodo_f Then
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
    If lCxOperacao = 2 Then
        cmd_novo.Enabled = False
        cmd_pesquisa.Enabled = False
        frm_move.Visible = False
        cmd_excluir.Enabled = False
        If cmd_alterar.Enabled = False Then
            If AberturaCaixa.LocalizarCodigo(g_empresa, lCxData, "NF", lCxPeriodo, lCxIlha, lCxCodigoFuncionario, lCxTipoMovimento) Then
                If AberturaCaixa.DataFechamento = "00:00:00" Then
                    cmd_alterar.Enabled = True
                End If
            End If
        End If
    ElseIf lCxOperacao = 3 Then
        cmd_novo.Enabled = False
        cmd_alterar.Enabled = False
        cmd_pesquisa.Enabled = False
        frm_move.Visible = False
        If cmd_excluir.Enabled = False Then
            If AberturaCaixa.LocalizarCodigo(g_empresa, lCxData, "NF", lCxPeriodo, lCxIlha, lCxCodigoFuncionario, lCxTipoMovimento) Then
                If AberturaCaixa.DataFechamento = "00:00:00" Then
                    cmd_excluir.Enabled = True
                End If
            End If
        End If
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If CDate(txtData.Text) < lDataI Or CDate(txtData.Text) > lDataF Then
        MsgBox "A data do movimento deve estar entre " & Format(lDataI, "dd/mm/yyyy") & " a " & Format(lDataF, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        If txtData.Enabled Then
            txtData.SetFocus
        Else
            cmd_cancelar.SetFocus
        End If
    ElseIf Val(cbo_periodo.Text) < lPeriodoI Or Val(cbo_periodo.Text) > lPeriodoF Then
        MsgBox "O período deve estar entre " & lPeriodoI & " ao " & lPeriodoF & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_oleo.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lIlha = RetiraGString(3)
        lTipoSubEstoque = RetiraGString(4)
        lCodigoProduto = RetiraGString(5)
        lOrdem = RetiraGString(6)
        lCodigoFuncionario = RetiraGString(7)
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar o registro!", vbInformation, "Erro de Integridade!"
        End If
        AtualizaGrid
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovimentoLubrificante.LocalizarPrimeiro() Then
        AtualTela
        AtualizaGrid
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovimentoLubrificante.LocalizarProximo Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovimentoLubrificante.LocalizarUltimo(g_empresa) Then
        AtualTela
        AtualizaGrid
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmdPesquisa_Click()
    Dim xString As String
    Dim xCodigo As Long
    
    xString = g_string
    g_string = ""
    consulta_produto.Show 1
    If Len(g_string) > 0 Then
        xCodigo = RetiraGString(1)
        If Produto.LocalizarCodigo(xCodigo) Then
            txt_produto.Text = xCodigo
            dtcboProduto.BoundText = CLng(txt_produto.Text)
            txt_produto_LostFocus
        End If
    End If
    g_string = xString
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" And lOpcao > 0 Then
        txt_produto.Text = dtcboProduto.BoundText
        If lOpcao = 2 Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, CDate(txtData.Text), cbo_periodo.Text, Val(cboIlha.Text), Val(cboTipoMovimento.ItemData(cboTipoMovimento.ListIndex)), Val(cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)), CLng(txt_produto.Text), Val(dtcboFuncionario.BoundText)) Then
                MsgBox "O produto " & Produto.Nome & Chr(10) & "Já tem venda neste caixa." & Chr(10) & "Por este motivo esta alteração não será aceita." & Chr(10) & "Se o produto estiver errado, deverá ser excluído.", vbInformation, "Atenção! Procedimento não aceito."
                dtcboProduto.BoundText = ""
                txt_produto.Text = ""
                cmd_cancelar_Click
                Exit Sub
            End If
        End If
        If lOpcao = 1 Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, CDate(txtData.Text), cbo_periodo.Text, Val(cboIlha.Text), Val(cboTipoMovimento.Text), Val(cboTipoSubEstoque.Text), CLng(txt_produto.Text), Val(txt_funcionario.Text)) Then
                MsgBox "Já existe movimento com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", vbInformation, "Duplicidade de Registro!"
                txt_produto.Text = ""
                dtcboProduto.BoundText = ""
                dtcboProduto.SetFocus
                Exit Sub
            End If
        End If
        txt_produto_LostFocus
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" And lOpcao > 0 Then
        txt_funcionario.Text = dtcboFuncionario.BoundText
        txt_funcionario_LostFocus
        txt_produto.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        flag_movimento_oleo_diverso = 0
    End If
    If flag_movimento_oleo_diverso = 0 Then
        AtualizaConstantes
        lGravados = 0
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        lDataI = g_cfg_data_i
        lDataF = g_cfg_data_f
        lPeriodoI = Val(g_cfg_periodo_i)
        lPeriodoF = Val(g_cfg_periodo_f)
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        Else
            If MovimentoLubrificante.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtivaBotoes
                AtualizaGrid
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                AtualizaGrid
            End If
            If cmd_novo.Enabled Then
                cmd_novo.SetFocus
            End If
        End If
    Else
        flag_movimento_oleo_diverso = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_oleo_diverso = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Dim xString As String
    
    Call GravaAuditoria(1, Me.name, 1, "")
    CentraForm Me
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    PreencheCboIlha
    PreencheCboPeriodo
    PreencheCboTipoSubEstoque
    PreencheCboTipoMovimento
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " AND [Periodo] < 5 ORDER BY [Nome]")
    
    xString = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xString <> "CUPOM FISCAL/CONVENIENCIA" Then
        If xString = "CONVENIENCIA" Then
            xString = " AND [Exclusivo Loja] = " & preparaBooleano(True)
        Else
            xString = " AND [Exclusivo Posto] = " & preparaBooleano(True)
        End If
    Else
        xString = ""
    End If
    xString = "SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & xString & " ORDER BY Nome"
    Set adodcProduto.Recordset = Conectar.RsConexao(xString)
    
    lCxPeriodo = 0
    lCxCodigoLancamentoPadrao = 1
    lCxOperacao = 0
    lCxTipoMovimento = 2
    lPriorizaSeguranca = True
    lCxCodigoUsuario = g_usuario
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Prioriza Segurança") Then
        lPriorizaSeguranca = ConfiguracaoDiversa.Verdadeiro
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub grid_oleo_DblClick()
    MarcaCelulaOleo
End Sub
Private Sub grid_oleo_GotFocus()
'    grid_oleo.Row = 1
'    grid_oleo.Col = 0
End Sub
Private Sub grid_oleo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MarcaCelulaOleo
    End If
End Sub
Private Sub MarcaCelulaOleo()
    grid_oleo.Col = 0
    If grid_oleo.Text <> "" Then
        grid_oleo.Col = 0
        lData = grid_oleo.Text
        grid_oleo.Col = 1
        lPeriodo = grid_oleo.Text
        grid_oleo.Col = 2
        lIlha = grid_oleo.Text
        grid_oleo.Col = 3
        lTipoSubEstoque = grid_oleo.Text
        grid_oleo.Col = 8
        lOrdem = grid_oleo.Text
        grid_oleo.Col = 9
        lCodigoProduto = grid_oleo.Text
        grid_oleo.Col = 4
        lCodigoFuncionario = grid_oleo.Text
        grid_oleo.Col = 10
        lTipoMovimento = grid_oleo.Text
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, lData, lPeriodo, lIlha, lTipoMovimento, lTipoSubEstoque, lCodigoProduto, lCodigoFuncionario) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar o registro!", vbInformation, "Erro de Integridade!"
        End If
        If cmd_alterar.Enabled Then
            cmd_alterar.SetFocus
        End If
    End If
End Sub
Private Sub MontaGrid()
    LimpaGrid
    grid_oleo.Row = 0
    grid_oleo.Col = 0
    grid_oleo.Text = "Data"
    grid_oleo.ColWidth(0) = TextWidth(String$(11, "9"))
    grid_oleo.ColAlignment(0) = 2
   'obs: o "9"equivale ao tab
    '0 = left, 1 = right ,2 =  center
    grid_oleo.Col = 1
    grid_oleo.Text = "Per."
    grid_oleo.ColWidth(1) = TextWidth(String$(4, "9"))
    grid_oleo.ColAlignment(1) = 2
    grid_oleo.Col = 2
    grid_oleo.Text = "Ilha"
    grid_oleo.ColWidth(2) = TextWidth(String$(4, "9"))
    grid_oleo.ColAlignment(2) = 2
    grid_oleo.Col = 3
    grid_oleo.Text = "Sub."
    grid_oleo.ColWidth(3) = TextWidth(String$(5, "9"))
    grid_oleo.ColAlignment(3) = 2
    grid_oleo.Col = 4
    grid_oleo.Text = "Func."
    grid_oleo.ColWidth(4) = TextWidth(String$(5, "9"))
    grid_oleo.ColAlignment(4) = 1
    grid_oleo.Col = 5
    grid_oleo.Text = "Nome"
    grid_oleo.ColWidth(5) = TextWidth(String$(25, "9"))
    grid_oleo.ColAlignment(5) = 0
    grid_oleo.Col = 6
    grid_oleo.Text = "Produto"
    grid_oleo.ColWidth(6) = TextWidth(String$(22, "9"))
    grid_oleo.ColAlignment(6) = 0
    grid_oleo.Col = 7
    grid_oleo.Text = "Valor"
    grid_oleo.ColWidth(7) = TextWidth(String$(8, "9"))
    grid_oleo.ColAlignment(7) = 1
    grid_oleo.Col = 8
    grid_oleo.Text = "Ordem Dig."
    grid_oleo.ColWidth(8) = TextWidth(String$(6, "9"))
    grid_oleo.ColAlignment(8) = 1
    grid_oleo.Col = 9
    grid_oleo.Text = "Codigo Prod"
    grid_oleo.ColWidth(9) = TextWidth(String$(6, "9"))
    grid_oleo.ColAlignment(9) = 1
    grid_oleo.Col = 10
    grid_oleo.Text = "Tipo Mov."
    grid_oleo.ColWidth(9) = TextWidth(String$(4, "9"))
    grid_oleo.ColAlignment(9) = 1
End Sub
Private Sub txt_funcionario_GotFocus()
    If lOpcao = 1 Then
        txt_funcionario.Text = ""
    End If
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario.Text) > 0 And lOpcao > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text)) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dtcboFuncionario.BoundText = Funcionario.Codigo
                txt_produto.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    End If
End Sub
Private Sub txt_produto_LostFocus()
    If fContemLetra(txt_produto.Text) Then
        SendKeys txt_produto.Text
        txt_produto.Text = ""
    End If
    If Val(txt_produto.Text) > 0 And lOpcao > 0 Then
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If Produto.Inativo = True Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            Else
                If lOpcao = 2 Then
                    If MovimentoLubrificante.LocalizarCodigo(g_empresa, CDate(txtData.Text), cbo_periodo.Text, Val(cboIlha.Text), Val(cboTipoMovimento.ItemData(cboTipoMovimento.ListIndex)), Val(cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)), CLng(txt_produto.Text), Val(dtcboFuncionario.BoundText)) Then
                        MsgBox "O produto " & Produto.Nome & Chr(10) & "Já tem venda neste caixa." & Chr(10) & "Por este motivo esta alteração não será aceita." & Chr(10) & "Se o produto estiver errado, deverá ser excluído.", vbInformation, "Atenção! Procedimento não aceito."
                        dtcboProduto.BoundText = ""
                        txt_produto.Text = ""
                        cmd_cancelar_Click
                        Exit Sub
                    End If
                End If
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                lPrecoCusto = Produto.PrecoCusto
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    If Estoque.PrecoVenda <> 0 Then
                        txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                    Else
                        txt_valor_unitario.Text = Format(Produto.PrecoVenda, "###,##0.0000")
                    End If
                Else
                    MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                    txt_valor_unitario.Text = ""
                    txt_valor_unitario.SetFocus
                    Exit Sub
                End If
                txt_quantidade.SetFocus
            End If
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
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_quantidade_LostFocus()
    txt_quantidade.Text = Format(txt_quantidade.Text, "###,##0.00")
    txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
End Sub
Private Sub txt_valor_total_GotFocus()
    If lPriorizaSeguranca = True And g_nivel_acesso >= 3 Then
        cmd_ok.SetFocus
        Exit Sub
    End If
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total.Text)
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
End Sub
Private Sub txt_valor_unitario_GotFocus()
    If lPriorizaSeguranca = True And g_nivel_acesso >= 3 Then
        txt_quantidade.SetFocus
        Exit Sub
    End If
    txt_valor_unitario.SelStart = 0
    txt_valor_unitario.SelLength = Len(txt_valor_unitario.Text)
End Sub
Private Sub txt_valor_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_unitario_LostFocus()
    txt_valor_unitario.Enabled = Format(txt_valor_unitario.Text, "###,##0.0000")
End Sub
Private Sub txtData_GotFocus()
    txtData.Text = fDesmascaraData(txtData.Text)
    txtData.SelStart = 0
    txtData.SelLength = 4
    txtData.MaxLength = 8
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtData_LostFocus()
    txtData.MaxLength = 10
    txtData.Text = fMascaraData(txtData.Text)
End Sub
