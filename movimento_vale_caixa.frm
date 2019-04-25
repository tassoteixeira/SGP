VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form movimento_vale_caixa 
   Caption         =   "Movimento de Vales do Caixa"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   11175
   Icon            =   "movimento_vale_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_vale_caixa.frx":030A
   ScaleHeight     =   6450
   ScaleWidth      =   11175
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_vale_caixa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cria um novo registro."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1080
      Picture         =   "movimento_vale_caixa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Altera o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2040
      Picture         =   "movimento_vale_caixa.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   3000
      Picture         =   "movimento_vale_caixa.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3960
      Picture         =   "movimento_vale_caixa.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5520
      Width           =   795
   End
   Begin MSGrid.Grid grid_vale_caixa 
      Height          =   2835
      Left            =   120
      TabIndex        =   22
      Top             =   2580
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   5001
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   11
      FixedCols       =   0
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   10935
      Begin VB.TextBox txt_observacao 
         Height          =   300
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2160
         Width           =   7275
      End
      Begin VB.ComboBox cbo_tipo_documento 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1380
         Width           =   3015
      End
      Begin VB.ComboBox cbo_tipo_vale 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1380
         Width           =   2175
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   300
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_numero_ilha 
         Height          =   300
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   300
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1020
         Width           =   555
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
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
         RecordSource    =   "Funcionario"
         Top             =   1020
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_valor 
         Height          =   300
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
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
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "movimento_vale_caixa.frx":7472
         Height          =   315
         Left            =   3120
         TabIndex        =   11
         Top             =   1020
         Width           =   5355
         _ExtentX        =   9446
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
         Caption         =   "Observação"
         Height          =   300
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10900
         Y1              =   980
         Y2              =   980
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do documento"
         Height          =   300
         Index           =   5
         Left            =   6240
         TabIndex        =   14
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do vale"
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Número do documento"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Número da &ilha"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário responsável"
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor do vale"
         Height          =   300
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do movimento"
         Height          =   300
         Index           =   7
         Left            =   6240
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   300
         Index           =   6
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do vale"
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   8880
      TabIndex        =   30
      Top             =   5400
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_vale_caixa.frx":7490
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_vale_caixa.frx":898A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_vale_caixa.frx":9E84
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_vale_caixa.frx":B2F6
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   9300
      Picture         =   "movimento_vale_caixa.frx":C878
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   10260
      Picture         =   "movimento_vale_caixa.frx":DE82
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5520
      Width           =   795
   End
End
Attribute VB_Name = "movimento_vale_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_vale_caixa As Integer
Dim lOpcao As String
Dim l_empresa As Integer
Dim l_data As Date
Dim l_periodo As String
Dim l_ilha As Integer
Dim l_tipo_movimento As String
Dim l_numero_documento As String * 10
Dim l_sql As String
Dim l_gravados As Long
'Dim l_vezes As Integer
Dim l_qtd_periodo As Integer
Dim tbl_configuracao As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_vale_caixa As Table
Private Sub AdcionaDadosGridValeCaixa()
    Dim x_i As Integer
    With tbl_movimento_vale_caixa
        grid_vale_caixa.Row = grid_vale_caixa.Rows - 1
        grid_vale_caixa.Col = 0
        grid_vale_caixa.Text = !Data
        grid_vale_caixa.Col = 1
        grid_vale_caixa.Text = !Periodo
        grid_vale_caixa.Col = 2
        grid_vale_caixa.Text = ![Numero da Ilha]
        grid_vale_caixa.Col = 3
        grid_vale_caixa.Text = ![Tipo do Movimento]
        grid_vale_caixa.Col = 4
        grid_vale_caixa.Text = ![Codigo do Funcionario]
        tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
        grid_vale_caixa.Col = 5
        If Not tbl_funcionario.NoMatch Then
            grid_vale_caixa.Text = tbl_funcionario!Nome
        Else
            grid_vale_caixa.Text = "** Não Cadastrado **"
        End If
        grid_vale_caixa.Col = 6
        If ![Tipo do Vale] = 1 Then
            grid_vale_caixa.Text = "Vale Emitido"
        Else
            grid_vale_caixa.Text = "Vale Recebido"
        End If
        grid_vale_caixa.Col = 7
        If ![Tipo do Documento] = 1 Then
            grid_vale_caixa.Text = "Nota de Abastecimento"
        ElseIf ![Tipo do Documento] = 2 Then
            grid_vale_caixa.Text = "Cheque"
        ElseIf ![Tipo do Documento] = 3 Then
            grid_vale_caixa.Text = "Despesas"
        ElseIf ![Tipo do Documento] = 4 Then
            grid_vale_caixa.Text = "Vale Funcionario"
        Else
            grid_vale_caixa.Text = "Outros"
        End If
        grid_vale_caixa.Col = 8
        grid_vale_caixa.Text = ![Numero do Documento]
        grid_vale_caixa.Col = 9
        grid_vale_caixa.Text = Format(!valor, "###,##0.00")
        grid_vale_caixa.Col = 10
        grid_vale_caixa.Text = !Observacao
        grid_vale_caixa.Rows = grid_vale_caixa.Rows + 1
    End With
End Sub
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
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub AtualizaConstantes()
    tbl_configuracao.Index = "id_codigo"
    tbl_configuracao.Seek "=", g_empresa
    If Not tbl_configuracao.NoMatch Then
        l_qtd_periodo = tbl_configuracao![Quantidade de Periodos]
    Else
        l_qtd_periodo = 1
    End If
End Sub
Private Sub AtualTabe()
    l_data = msk_data
    l_periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    l_ilha = Val(txt_numero_ilha)
    l_tipo_movimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    l_numero_documento = txt_numero_documento
    With tbl_movimento_vale_caixa
        !Empresa = g_empresa
        !Data = msk_data
        !Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
        ![Numero da Ilha] = Val(txt_numero_ilha)
        ![Tipo do Movimento] = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
        ![Codigo do Funcionario] = CLng(dbcbo_funcionario.BoundText)
        ![Tipo do Vale] = cbo_tipo_vale.ItemData(cbo_tipo_vale.ListIndex)
        ![Tipo do Documento] = cbo_tipo_documento.ItemData(cbo_tipo_documento.ListIndex)
        ![Numero do Documento] = txt_numero_documento
        !valor = fValidaValor(txt_valor)
        !Observacao = txt_observacao
    End With
End Sub
Private Sub AtualTela()
    Dim i As Integer
    With tbl_movimento_vale_caixa
        l_data = !Data
        l_periodo = !Periodo
        l_ilha = ![Numero da Ilha]
        l_tipo_movimento = ![Tipo do Movimento]
        l_numero_documento = ![Numero do Documento]
        msk_data = Format(!Data, "dd/mm/yyyy")
        cbo_periodo.ListIndex = !Periodo - 1
        txt_numero_ilha = ![Numero da Ilha]
        cbo_tipo_movimento.ListIndex = ![Tipo do Movimento] - 1
        dbcbo_funcionario.BoundText = ""
        tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
        If Not tbl_funcionario.NoMatch Then
            txt_funcionario = ![Codigo do Funcionario]
            dbcbo_funcionario.BoundText = ![Codigo do Funcionario]
        End If
        cbo_tipo_vale.ListIndex = ![Tipo do Vale] - 1
        cbo_tipo_documento.ListIndex = ![Tipo do Documento] - 1
        txt_numero_documento = ![Numero do Documento]
        txt_valor = Format(!valor, "###,###,##0.00")
        txt_observacao = !Observacao
    End With
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Function BuscaRegistro(x_data As Date, x_periodo As String, x_ilha As Integer, x_tipo_movimento As String, x_numero_documento As String) As Boolean
    BuscaRegistro = False
    If tbl_movimento_vale_caixa.RecordCount > 0 Then
        tbl_movimento_vale_caixa.Seek "=", g_empresa, x_data, x_periodo, x_ilha, x_tipo_movimento, x_numero_documento
        If Not tbl_movimento_vale_caixa.NoMatch Then
            AtualTela
            BuscaRegistro = True
        End If
    End If
End Function
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_movimento_vale_caixa
        If .RecordCount > 0 Then
            If lOpcao = 3 Then
                If Not .EOF Then
                    .MoveNext
                    If Not .EOF Then
                        If !Empresa = g_empresa Then
                            AtualTela
                            BuscaDados = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            .Seek "<", g_empresa, CDate("31/12/2500"), "9", 9, "9"
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    BuscaDados = True
                    Exit Function
                End If
            End If
        End If
        l_gravados = 0
        LimpaTela
    End With
End Function
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    With tbl_movimento_vale_caixa
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate("31/12/2500"), 9, 9, 9, "ZZZZZZZZZZ"
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    msk_data = !Data
                    x_periodo = !Periodo
                    txt_numero_ilha = 1
                    If !Periodo >= l_qtd_periodo Then
                        msk_data = !Data + 1
                        x_periodo = 0
                    End If
                    cbo_periodo.ListIndex = x_periodo
                    cbo_tipo_movimento.ListIndex = 0
                    BuscaProximoCaixa = True
                    Exit Function
                End If
            End If
        End If
        msk_data = g_data_def - 1
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
        txt_numero_ilha = 1
    End With
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_configuracao.Close
    tbl_funcionario.Close
    tbl_movimento_vale_caixa.Close
End Sub
Private Sub PesquisaValeCaixa()
    MontaGridValeCaixa
    With tbl_movimento_vale_caixa
        .Seek ">=", g_empresa, l_data, l_periodo, l_ilha, l_tipo_movimento, "        "
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> g_empresa Or !Data <> l_data Or !Periodo <> l_periodo Or ![Tipo do Movimento] <> l_tipo_movimento Or ![Numero da Ilha] <> l_ilha Then
                    Exit Do
                End If
                AdcionaDadosGridValeCaixa
                .MoveNext
            Loop
        End If
        Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento, l_numero_documento)
        grid_vale_caixa.Row = grid_vale_caixa.Rows - 1
        grid_vale_caixa.Col = 0
    End With
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
Private Sub PreencheCboTipoDocumento()
    cbo_tipo_documento.Clear
    cbo_tipo_documento.AddItem "1 - Nota de Abastecimento"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 1
    cbo_tipo_documento.AddItem "2 - Cheque"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 2
    cbo_tipo_documento.AddItem "3 - Despesas"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 3
    cbo_tipo_documento.AddItem "4 - Vale Funcionario"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 4
    cbo_tipo_documento.AddItem "5 - Outros"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 5
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 - Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Caixa da Borracharia/Lavagem"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub PreencheCboTipoVale()
    cbo_tipo_vale.Clear
    cbo_tipo_vale.AddItem "1 - Vale Emitido"
    cbo_tipo_vale.ItemData(cbo_tipo_vale.NewIndex) = 1
    cbo_tipo_vale.AddItem "2 - Vale Recebido"
    cbo_tipo_vale.ItemData(cbo_tipo_vale.NewIndex) = 2
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_ilha.SetFocus
    End If
End Sub
Private Sub cbo_tipo_documento_GotFocus()
    SendMessageLong cbo_tipo_documento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_documento.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub cbo_tipo_vale_GotFocus()
    SendMessageLong cbo_tipo_vale.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_vale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_documento.SetFocus
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
    txt_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If tbl_movimento_vale_caixa.RecordCount > 0 Then
        tbl_movimento_vale_caixa.MovePrevious
        If Not tbl_movimento_vale_caixa.BOF Then
            If tbl_movimento_vale_caixa!Empresa = g_empresa Then
                AtualTela
                PesquisaValeCaixa
                Exit Sub
            End If
        End If
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        tbl_movimento_vale_caixa.MoveNext
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento, l_numero_documento) Then
        AtivaBotoes
        PesquisaValeCaixa
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
Sub LimpaGridValeCaixa()
    Do Until grid_vale_caixa.Rows = 2
        grid_vale_caixa.Row = grid_vale_caixa.Rows - 1
        grid_vale_caixa.RemoveItem grid_vale_caixa.Row
    Loop
    grid_vale_caixa.Row = 1
    grid_vale_caixa.Col = 0
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 1
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 2
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 3
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 4
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 5
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 6
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 7
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 8
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 9
    grid_vale_caixa.Text = ""
    grid_vale_caixa.Col = 10
    grid_vale_caixa.Text = ""
End Sub
Private Sub LimpaTela()
    If l_gravados = 0 Then
        msk_data = "__/__/____"
        cbo_periodo.ListIndex = -1
        txt_numero_ilha = ""
        cbo_tipo_movimento.ListIndex = -1
    End If
    txt_funcionario = ""
    dbcbo_funcionario.BoundText = ""
    cbo_tipo_vale.ListIndex = -1
    cbo_tipo_documento.ListIndex = -1
    txt_numero_documento = ""
    txt_valor = ""
    txt_observacao = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If fValidaValor(txt_valor) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
'            Call GravaAuditoria(1, Me.name, 10, "Data:" & lData & " Per:" & l_periodo & " Vlr:" & txt_valor.Text)
            lOpcao = 3
            tbl_movimento_vale_caixa.Edit
            tbl_movimento_vale_caixa.Delete
            If Not BuscaDados Then
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
            lOpcao = 0
            PesquisaValeCaixa
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If l_gravados = 0 Then
        If BuscaProximoCaixa Then
            txt_funcionario.SetFocus
        Else
            msk_data.SetFocus
        End If
    Else
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                tbl_movimento_vale_caixa.AddNew
                AtualTabe
                tbl_movimento_vale_caixa.Update
                l_gravados = 1
            ElseIf lOpcao = 2 Then
                tbl_movimento_vale_caixa.Edit
                AtualTabe
                tbl_movimento_vale_caixa.Update
            End If
            PesquisaValeCaixa
            Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento, l_numero_documento)
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_movimento_vale_caixa.name, "Movimento Vale Caixao"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data do vale.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", 64, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf Not Val(txt_numero_ilha) > 0 Then
        MsgBox "O número da ilha deve ser maior que 0.", 64, "Atenção!"
        txt_numero_ilha.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf dbcbo_funcionario.BoundText = "" Then
        MsgBox "Escolha o funcionario.", 64, "Atenção!"
        dbcbo_funcionario.SetFocus
    ElseIf cbo_tipo_vale.ListIndex = -1 Then
        MsgBox "Escolha o tipo do vale.", 64, "Atenção!"
        cbo_tipo_vale.SetFocus
    ElseIf cbo_tipo_documento.ListIndex = -1 Then
        MsgBox "Escolha o tipo do documento.", 64, "Atenção!"
        cbo_tipo_documento.SetFocus
    ElseIf txt_numero_documento = "" Then
        MsgBox "Informe o número do documento.", 64, "Atenção!"
        txt_numero_documento.SetFocus
    ElseIf Not fValidaValor(txt_valor) > 0 Then
        MsgBox "Informe o valor do vale.", 64, "Atenção!"
        txt_valor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    With tbl_movimento_vale_caixa
        If g_nivel_acesso > 4 Then
            If !Empresa < g_cfg_empresa_i Or !Empresa > g_cfg_empresa_f Then
                x_flag = False
            ElseIf !Data < g_cfg_data_i Or !Data > g_cfg_data_f Then
                x_flag = False
            ElseIf !Periodo < g_cfg_periodo_i Or !Periodo > g_cfg_periodo_f Then
                x_flag = False
            End If
        End If
    End With
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
    If msk_data < g_cfg_data_i Or msk_data > g_cfg_data_f Then
        MsgBox "A data do movimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", 64, "Digitação Não Autorizada!"
        msk_data.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", 64, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_vale_caixa.Show 1
    If Len(g_string) > 0 Then
        l_data = RetiraGString(1)
        l_periodo = RetiraGString(2)
        l_ilha = RetiraGString(3)
        l_tipo_movimento = RetiraGString(4)
        l_numero_documento = RetiraGString(5)
        Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento, l_numero_documento)
        PesquisaValeCaixa
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If tbl_movimento_vale_caixa.RecordCount > 0 Then
        tbl_movimento_vale_caixa.Seek ">", g_empresa, CDate("01/01/1900"), 0, 0, 0
        If Not tbl_movimento_vale_caixa.NoMatch Then
            If tbl_movimento_vale_caixa!Empresa = g_empresa Then
                AtualTela
                PesquisaValeCaixa
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If tbl_movimento_vale_caixa.RecordCount > 0 Then
        tbl_movimento_vale_caixa.MoveNext
        If Not tbl_movimento_vale_caixa.EOF Then
            If tbl_movimento_vale_caixa!Empresa = g_empresa Then
                AtualTela
                PesquisaValeCaixa
                Exit Sub
            End If
        End If
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        tbl_movimento_vale_caixa.MovePrevious
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If tbl_movimento_vale_caixa.RecordCount > 0 Then
        tbl_movimento_vale_caixa.Seek "<", g_empresa, CDate("31/12/2500"), 9, 9, 9
        If Not tbl_movimento_vale_caixa.NoMatch Then
            If tbl_movimento_vale_caixa!Empresa = g_empresa Then
                AtualTela
                PesquisaValeCaixa
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cbo_tipo_vale.SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If dbcbo_funcionario.BoundText <> "" And lOpcao > 0 Then
        txt_funcionario = dbcbo_funcionario.BoundText
        txt_funcionario_LostFocus
        cbo_tipo_vale.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> l_empresa Then
        flag_movimento_vale_caixa = 0
    End If
    If flag_movimento_vale_caixa = 0 Then
        AtualizaConstantes
        l_gravados = 0
        lOpcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
            PesquisaValeCaixa
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
            PesquisaValeCaixa
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        flag_movimento_vale_caixa = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_vale_caixa = 1
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
    CentraForm Me
    Set tbl_configuracao = bd_sgp.OpenTable("configuracao")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_vale_caixa = bd_sgp.OpenTable("Movimento_Vale_Caixa")
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_vale_caixa.Index = "id_data"
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    PreencheCboTipoVale
    PreencheCboTipoDocumento
    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = " & Chr(34) & "A" & Chr(34) & " And [Periodo] < 5 Order By [Nome]"
    dta_funcionario.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub grid_vale_caixa_DblClick()
    MarcaCelulaValeCaixa
End Sub
Private Sub grid_vale_caixa_GotFocus()
'    grid_vale_caixa.Row = 1
'    grid_vale_caixa.Col = 0
End Sub
Private Sub grid_vale_caixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MarcaCelulaValeCaixa
    End If
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub MarcaCelulaValeCaixa()
    'grid_vale_caixa.Col = 0
    'If grid_vale_caixa.Text <> "" Then
    '    grid_vale_caixa.Col = 0
    '    l_data = grid_vale_caixa.Text
    '    grid_vale_caixa.Col = 1
    '    l_periodo = grid_vale_caixa.Text
    '    grid_vale_caixa.Col = 3
    '    l_tipo_movimento = grid_vale_caixa.Text
    '    grid_vale_caixa.Col = 2
    '    l_ilha = grid_vale_caixa.Text
    '    grid_vale_caixa.Col = 4
    '    l_codigo_funcionario = grid_vale_caixa.Text
    '    grid_vale_caixa.Col = 8
    '    l_numero_documento = grid_vale_caixa.Text
    '    Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento, l_numero_documento)
    '    cmd_alterar.SetFocus
    'End If
End Sub
Private Sub MontaGridValeCaixa()
    LimpaGridValeCaixa
    grid_vale_caixa.Row = 0
    grid_vale_caixa.Col = 0
    grid_vale_caixa.Text = "Data"
    grid_vale_caixa.ColWidth(0) = TextWidth(String$(11, "9"))
    grid_vale_caixa.ColAlignment(0) = 2
   'obs: o "9"equivale ao tab
    '0 = left, 1 = right ,2 =  center
    grid_vale_caixa.Col = 1
    grid_vale_caixa.Text = "Per."
    grid_vale_caixa.ColWidth(1) = TextWidth(String$(4, "9"))
    grid_vale_caixa.ColAlignment(1) = 2
    grid_vale_caixa.Col = 2
    grid_vale_caixa.Text = "Ilha"
    grid_vale_caixa.ColWidth(2) = TextWidth(String$(4, "9"))
    grid_vale_caixa.ColAlignment(2) = 2
    grid_vale_caixa.Col = 3
    grid_vale_caixa.Text = "Mov."
    grid_vale_caixa.ColWidth(3) = TextWidth(String$(5, "9"))
    grid_vale_caixa.ColAlignment(3) = 2
    grid_vale_caixa.Col = 4
    grid_vale_caixa.Text = "Func."
    grid_vale_caixa.ColWidth(4) = TextWidth(String$(5, "9"))
    grid_vale_caixa.ColAlignment(4) = 1
    grid_vale_caixa.Col = 5
    grid_vale_caixa.Text = "Nome do Funcionário"
    grid_vale_caixa.ColWidth(5) = TextWidth(String$(25, "9"))
    grid_vale_caixa.ColAlignment(5) = 0
    grid_vale_caixa.Col = 6
    grid_vale_caixa.Text = "Tipo do Vale"
    grid_vale_caixa.ColWidth(6) = TextWidth(String$(15, "9"))
    grid_vale_caixa.ColAlignment(6) = 0
    grid_vale_caixa.Col = 7
    grid_vale_caixa.Text = "Tipo do Documento"
    grid_vale_caixa.ColWidth(7) = TextWidth(String$(20, "9"))
    grid_vale_caixa.ColAlignment(7) = 0
    grid_vale_caixa.Col = 8
    grid_vale_caixa.Text = "Número do Documento"
    grid_vale_caixa.ColWidth(8) = TextWidth(String$(11, "9"))
    grid_vale_caixa.ColAlignment(8) = 0
    grid_vale_caixa.Col = 9
    grid_vale_caixa.Text = "Valor"
    grid_vale_caixa.ColWidth(9) = TextWidth(String$(9, "9"))
    grid_vale_caixa.ColAlignment(9) = 1
    grid_vale_caixa.Col = 10
    grid_vale_caixa.Text = "Observação"
    grid_vale_caixa.ColWidth(10) = TextWidth(String$(40, "9"))
    grid_vale_caixa.ColAlignment(10) = 0
End Sub
Private Sub txt_funcionario_GotFocus()
    If lOpcao = 1 Then
        txt_funcionario = ""
    End If
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario) > 0 And lOpcao > 0 Then
        tbl_funcionario.Seek "=", g_empresa, Val(txt_funcionario)
        If Not tbl_funcionario.NoMatch Then
            If tbl_funcionario!Situacao = "I" Then
                MsgBox "O funcionário " & Trim(tbl_funcionario!Nome) & " está inativo.", 64, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dbcbo_funcionario.BoundText = tbl_funcionario!Codigo
                cbo_tipo_vale.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", 64, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_numero_ilha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Function ExisteRegistro() As Boolean
    'ExisteRegistro = False
    'With tbl_movimento_vale_caixa
    '    If .RecordCount > 0 Then
    '        .Index = "id_data"
    '        .Seek "=", g_empresa, CDate(msk_data), Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)), Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)), CLng(txt_produto), Val(txt_funcionario)
    '        If Not .NoMatch Then
    '           MsgBox "Já existe movimento com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", 64, "Duplicidade de Registro!"
    '           ExisteRegistro = True
    '        End If
    '        .Index = "id_digitacao"
    '    End If
    'End With
End Function
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,###,##0.00")
End Sub
