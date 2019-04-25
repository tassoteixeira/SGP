VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_despesa_caixa 
   Caption         =   "Movimento de Despesas do Caixa"
   ClientHeight    =   4320
   ClientLeft      =   2250
   ClientTop       =   1530
   ClientWidth     =   7035
   Icon            =   "movimento_despesa_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_despesa_caixa.frx":030A
   ScaleHeight     =   4320
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_despesa_caixa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_despesa_caixa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_despesa_caixa.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_despesa_caixa.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Altera o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_despesa_caixa.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cria um novo registro."
      Top             =   3360
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.ComboBox cboIlha 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_e 
         Height          =   315
         Left            =   2880
         Picture         =   "movimento_despesa_caixa.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_valor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   23
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   285
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2040
         Width           =   1035
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   315
         ItemData        =   "movimento_despesa_caixa.frx":874C
         Left            =   1740
         List            =   "movimento_despesa_caixa.frx":874E
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txt_registro 
         Height          =   285
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
      Begin MSMask.MaskEdBox msk_data_movimento 
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodc_fornecedor 
         Height          =   330
         Left            =   3180
         Top             =   1320
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "adodc_fornecedor"
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
      Begin MSDataListLib.DataCombo dtcbo_fornecedor 
         Bindings        =   "movimento_despesa_caixa.frx":8750
         DataSource      =   "adodc_fornecedor"
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_fornecedor"
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5580
         TabIndex        =   37
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   255
         Left            =   4980
         TabIndex        =   36
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Número da &Ilha"
         Height          =   300
         Left            =   4440
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   315
         Index           =   6
         Left            =   4260
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl_data_digitacao 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5580
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data da Digitação"
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "&Valor da Despesa"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "&Data do Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Fornecedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Co&mplemento"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "&Número Documento"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Registro"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Cont&a"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   31
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_despesa_caixa.frx":876F
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_despesa_caixa.frx":9C69
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_despesa_caixa.frx":B163
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_despesa_caixa.frx":C5D5
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "movimento_despesa_caixa.frx":DB57
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "movimento_despesa_caixa.frx":F051
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3360
      Width           =   795
   End
End
Attribute VB_Name = "movimento_despesa_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lEmpresa As Integer
Dim lOpcao As Integer
Dim lData As Date
Dim lRegistro As Long
Dim lGravados As Integer
Dim lNumeroMovimentoCaixa As Long
Dim lDataAnterior As Date
Dim lPeriodoAnterior As Integer
Dim lValorAnterior As Currency
Dim lFornecedorAnterior As Currency

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxTipoMovimento As Integer
Dim lCxRegistro As Integer
Dim lCxTipoMov As Integer
Dim lCxIlha As Integer
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer

Private rsTabela As New adodb.Recordset

Private BaixaPagar As New cBaixaPagar
Private Configuracao As New cConfiguracao
Private Contas As New cContas
Private IntegracaoCaixa As New cIntegracaoCaixa
Private Fornecedor As New cFornecedor
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovContaPagar As New cMovimentoContaPagar
Private MovDespesaCaixa As New cMovimentoDespesaCaixa
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
    lCxTipoMov = Val(RetiraString(7, xString))
    lCxCodigoLancamentoPadrao = Val(RetiraString(8, xString))

    lData = lCxData
    'lPeriodo = lCxPeriodo
    'lTipoMovimento = lCxTipoMov
    'AtualizaDataGrid()

    txt_registro.Enabled = False
    lbl_data_digitacao.Enabled = False
    msk_data_movimento.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    cbo_tipo_movimento.Enabled = False
    If xOperacao = "Incluir" Then
        cmd_novo_Click
    ElseIf xOperacao = "Alterar" Then
        lCxRegistro = CLng(RetiraString(9, xString))
        If MovDespesaCaixa.LocalizarCodigo(g_empresa, lCxRegistro) Then
            AtualTela
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lCxRegistro = CLng(RetiraString(9, xString))
        If MovDespesaCaixa.LocalizarCodigo(g_empresa, lCxRegistro) Then
            AtualTela
        End If
        cmd_excluir_Click
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_registro.Enabled = True
End Sub
Private Sub AtualTabe()
    MovDespesaCaixa.Empresa = g_empresa
    MovDespesaCaixa.Registro = Val(txt_registro.Text)
    MovDespesaCaixa.DataDigitacao = CDate(lbl_data_digitacao.Caption)
    MovDespesaCaixa.DataMovimento = CDate(msk_data_movimento.Text)
    MovDespesaCaixa.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovDespesaCaixa.NumeroIlha = Val(cboIlha.Text)
    MovDespesaCaixa.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovDespesaCaixa.CodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
    MovDespesaCaixa.Valor = fValidaValor2(txt_valor.Text)
    MovDespesaCaixa.NumeroDocumento = txt_numero_documento.Text
    MovDespesaCaixa.CodigoConta = cbo_conta.ItemData(cbo_conta.ListIndex)
    MovDespesaCaixa.Complemento = Trim(txt_complemento.Text)
    MovDespesaCaixa.NumeroContaContabil = Fornecedor.ContaContabil
    'If lOpcao = 1 Then
        MovDespesaCaixa.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
        MovDespesaCaixa.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
    'End If
    'Baixa Pagar
    BaixaPagar.Empresa = MovDespesaCaixa.Empresa
    BaixaPagar.Registro = MovDespesaCaixa.Registro
    BaixaPagar.CodigoFornecedor = MovDespesaCaixa.CodigoFornecedor
    BaixaPagar.NomeFornecedor = dtcbo_fornecedor.Text
    BaixaPagar.DataEmissao = MovDespesaCaixa.DataMovimento
    BaixaPagar.DataVencimento = MovDespesaCaixa.DataMovimento
    BaixaPagar.Valor = MovDespesaCaixa.Valor
    BaixaPagar.NumeroDocumento = MovDespesaCaixa.NumeroDocumento
    BaixaPagar.LocalCobranca = 1
    BaixaPagar.CodigoConta = MovDespesaCaixa.CodigoConta
    BaixaPagar.Complemento = MovDespesaCaixa.Complemento
    BaixaPagar.DataDigitacao = MovDespesaCaixa.DataDigitacao
    BaixaPagar.DataPagamento = MovDespesaCaixa.DataMovimento
    BaixaPagar.ValorPagamento = MovDespesaCaixa.Valor
    BaixaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    If IsNull(BaixaPagar.NumeroMovimentoCaixa) Then
        BaixaPagar.NumeroMovimentoCaixa = 0
    End If
    If IsNull(BaixaPagar.NumeroMovimentoCaixaBaixa) Then
        BaixaPagar.NumeroMovimentoCaixaBaixa = 0
    End If
    BaixaPagar.TipoBaixa = 2
End Sub
Private Sub BuscaUltimoLancamento()
    Dim i As Integer
    Dim x_data As Date
    x_data = "00:00:00"
    If MovContaPagar.LocalizarUltimoFornecedor(g_empresa, Val(dtcbo_fornecedor.BoundText)) Then
        x_data = MovContaPagar.DataEmissao
        txt_complemento.Text = MovContaPagar.Complemento
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(i) = MovContaPagar.CodigoConta Then
                Exit For
            End If
        Next
    End If
    If BaixaPagar.LocalizarUltimoFornecedor(g_empresa, Val(dtcbo_fornecedor.BoundText)) Then
        If x_data < BaixaPagar.DataEmissao Then
            x_data = BaixaPagar.DataEmissao
            txt_complemento.Text = BaixaPagar.Complemento
            For i = 0 To cbo_conta.ListCount - 1
                cbo_conta.ListIndex = i
                If cbo_conta.ItemData(i) = BaixaPagar.CodigoConta Then
                    Exit For
                End If
            Next
        End If
    End If
    If MovDespesaCaixa.LocalizarUltimoFornecedor(g_empresa, Val(dtcbo_fornecedor.BoundText)) Then
        If x_data < MovDespesaCaixa.DataMovimento Then
            x_data = MovDespesaCaixa.DataMovimento
            txt_complemento.Text = MovDespesaCaixa.Complemento
            For i = 0 To cbo_conta.ListCount - 1
                cbo_conta.ListIndex = i
                If cbo_conta.ItemData(i) = MovDespesaCaixa.CodigoConta Then
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = MovDespesaCaixa.DataMovimento
    lRegistro = MovDespesaCaixa.Registro
    lDataAnterior = MovDespesaCaixa.DataMovimento
    lPeriodoAnterior = MovDespesaCaixa.Periodo
    lValorAnterior = MovDespesaCaixa.Valor
    lFornecedorAnterior = MovDespesaCaixa.CodigoFornecedor
    
    txt_registro.Text = MovDespesaCaixa.Registro
    lbl_data_digitacao = Format(MovDespesaCaixa.DataDigitacao, "dd/mm/yyyy")
    msk_data_movimento = Format(MovDespesaCaixa.DataMovimento, "dd/mm/yyyy")
    cbo_periodo.ListIndex = -1
    For i = 0 To cbo_periodo.ListCount - 1
        If cbo_periodo.ItemData(i) = MovDespesaCaixa.Periodo Then
            cbo_periodo.ListIndex = i
            Exit For
        End If
    Next
    cboIlha.ListIndex = MovDespesaCaixa.NumeroIlha - 1
    cbo_tipo_movimento.ListIndex = -1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        If cbo_tipo_movimento.ItemData(i) = MovDespesaCaixa.TipoMovimento Then
            cbo_tipo_movimento.ListIndex = i
            Exit For
        End If
    Next
    dtcbo_fornecedor.BoundText = MovDespesaCaixa.CodigoFornecedor
    txt_valor.Text = Format(MovDespesaCaixa.Valor, "###,##0.00")
    txt_numero_documento.Text = MovDespesaCaixa.NumeroDocumento
    If Contas.LocalizarCodigo(MovDespesaCaixa.CodigoConta) Then
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(i) = Contas.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_conta.ListIndex = -1
    End If
    txt_complemento.Text = MovDespesaCaixa.Complemento
    lNumeroMovimentoCaixa = MovDespesaCaixa.NumeroMovimentoCaixaCredito
    lbl_total = Format(MovDespesaCaixa.TotalPeriodo(g_empresa, CDate(msk_data_movimento), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text)), "###,##0.00")
    frm_dados.Enabled = False
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
Private Sub ExcluiMovimentoCaixa()
    If MovimentoCaixaPista.LocalizarCodigo(g_empresa, lData, lNumeroMovimentoCaixa) Then
        lCxDataDigitacao = MovimentoCaixaPista.DataDigitacao
        lCxHoraDigitacao = MovimentoCaixaPista.HoraDigitacao
        If Not MovimentoCaixaPista.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
            MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    Else
        MsgBox "Não foi possível localizar o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set BaixaPagar = Nothing
    Set Configuracao = Nothing
    Set Contas = Nothing
    Set Fornecedor = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovContaPagar = Nothing
    Set MovDespesaCaixa = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_registro.Text = MovDespesaCaixa.ProximoRegistro(g_empresa)
    lbl_data_digitacao = Format(g_data_def, "dd/mm/yyyy")
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_complemento.SetFocus
    End If
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlha.SetFocus
    End If
End Sub
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_fornecedor.SetFocus
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
    frm_dados.Enabled = True
    txt_registro.Enabled = False
    dtcbo_fornecedor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovDespesaCaixa.LocalizarAnterior Then
        AtualTela
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
    LimpaTela
    If MovDespesaCaixa.LocalizarCodigo(g_empresa, lRegistro) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    If lGravados = 0 Then
        txt_registro.Text = ""
        lbl_data_digitacao.Caption = ""
        msk_data_movimento.Text = "__/__/____"
        cbo_periodo.ListIndex = -1
        cboIlha.ListIndex = 0
        cbo_tipo_movimento.ListIndex = -1
    End If
    dtcbo_fornecedor.BoundText = ""
    txt_valor.Text = ""
    txt_numero_documento.Text = ""
    cbo_conta.ListIndex = -1
    txt_complemento.Text = ""
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "DESPESA CAIXA") Then
        xComplemento = Trim(txt_complemento.Text) & " TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
        
        
        
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = CDate(msk_data_movimento.Text)
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Valor = fValidaValor(txt_valor.Text)
        MovimentoCaixaPista.NumeroDocumento = txt_numero_documento.Text
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.Complemento = xComplemento
        MovimentoCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        If Len(Fornecedor.ContaContabil) = 9 Then
            MovimentoCaixaPista.NumeroContaDebito = Fornecedor.ContaContabil
        End If
        MovimentoCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovimentoCaixaPista.TipoMovimento = lCxTipoMovimento
        MovimentoCaixaPista.CodigoUsuario = g_usuario
        MovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        MovimentoCaixaPista.DadosInterno = "DESPC" & "|@|" & txt_registro.Text & "|@|"
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
            MsgBox "Não foi integrado no caixa o valor=" & txt_valor.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "DESPESA CAIXA" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub cmd_data_e_Click()
    g_string = msk_data_movimento
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_movimento = RetiraGString(1)
        cbo_periodo.SetFocus
    Else
        msk_data_movimento = RetiraGString(1)
        cbo_periodo.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If Val(txt_registro.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Not BaixaPagar.Excluir(g_empresa, CLng(txt_registro.Text)) Then
                MsgBox "Não foi possível excluir registro de baixa!", vbInformation, "Erro de Integridade!"
            End If
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & msk_data_movimento.Text & " Per:" & cbo_periodo.Text & " N.Reg:" & txt_registro.Text & " Vlr:" & txt_valor.Text & " Cod.Forn:" & dtcbo_fornecedor.BoundColumn)
            Call ExcluiMovimentoCaixa
            If MovDespesaCaixa.Excluir(g_empresa, Val(txt_registro.Text)) Then
                LimpaTela
                If MovDespesaCaixa.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtivaBotoes
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Registro não excluido!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
End Sub
Private Sub cmd_novo_Click()
    '
    'bd_sgp.Execute "UpDate Contas_Pagar Set [Data da Digitacao] = data_emissao"
    '
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    'atribui o valor de g_empresa para lEmpresa
    'caso o usuário cria novo registro e muda de
    'empresa sem cancelar o sistema o avisará
    lEmpresa = g_empresa
    If lGravados = 0 Then
        If lCxPeriodo > 0 Then
            msk_data_movimento.Text = Format(lCxData, "dd/mm/yyyy")
            cbo_periodo.ListIndex = lCxPeriodo - 1
            cboIlha.ListIndex = lCxIlha - 1
            cbo_tipo_movimento.ListIndex = lCxTipoMov - 1
            dtcbo_fornecedor.SetFocus
            Exit Sub
        End If
        msk_data_movimento.SetFocus
    Else
        dtcbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaDespesaDeCaixa
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & msk_data_movimento.Text & " Per:" & cbo_periodo.Text & " N.Reg:" & txt_registro.Text & " Vlr:" & txt_valor.Text & " Cod.Forn:" & dtcbo_fornecedor.BoundColumn)
            If MovDespesaCaixa.Incluir Then
                lGravados = lGravados + 1
                lRegistro = Val(txt_registro.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
            If Not BaixaPagar.Incluir Then
                MsgBox "Não foi possível incluir o registro de baixa!", vbInformation, "Erro de Integridade!"
            End If
        ElseIf lOpcao = 2 Then
            Call ExcluiMovimentoCaixa
            lData = CDate(msk_data_movimento.Text)
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            If Not BaixaPagar.Excluir(g_empresa, CLng(txt_registro.Text)) Then
                MsgBox "Não foi possível excluir registro de baixa!", vbInformation, "Erro de Integridade!"
            End If
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, "De: Dt:" & lDataAnterior & " Per:" & lPeriodoAnterior & " N.Reg:" & txt_registro.Text & " Vlr:" & lValorAnterior & " Cod.Forn:" & lFornecedorAnterior)
            Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & msk_data_movimento.Text & " Per:" & cbo_periodo.Text & " N.Reg:" & txt_registro.Text & " Vlr:" & txt_valor.Text & " Cod.Forn:" & dtcbo_fornecedor.BoundColumn)
            If Not MovDespesaCaixa.Alterar(g_empresa, Val(txt_registro.Text)) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
            If Not BaixaPagar.Incluir Then
                MsgBox "Não foi possível incluir o registro de baixa!", vbInformation, "Erro de Integridade!"
            End If
            If lCxPeriodo > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
        End If
        Call MovDespesaCaixa.LocalizarCodigo(g_empresa, lRegistro)
        AtualTela
        If lCxPeriodo > 0 And lOpcao = 1 Then
            lOpcao = 0
            cmd_novo_Click
            Exit Sub
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_movimento_despesa_caixa.Name, "Despesa de Caixaa"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_registro.Text) > 0 Then
        MsgBox "Informe o número do registro.", vbInformation, "Atenção!"
        txt_registro.SetFocus
    ElseIf Not IsDate(msk_data_movimento.Text) Then
        MsgBox "Informe a data do movimento.", vbInformation, "Atenção!"
        msk_data_movimento.SetFocus
    ElseIf Not ValidaPeriodo Then
        cbo_periodo.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Selecione uma ilha.", vbInformation, "Atenção!"
        cboIlha.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf Not Val(dtcbo_fornecedor.BoundText) > 0 Then
        MsgBox "Selecione o fornecedor.", vbInformation, "Atenção!"
        dtcbo_fornecedor.SetFocus
    ElseIf Not fValidaValor2(txt_valor.Text) > 0 Then
        MsgBox "Informe o valor do vencimento.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    ElseIf Not ValidaConta Then
        cbo_conta.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaConta() As Integer
    ValidaConta = False
    If Not Fornecedor.LocalizarCodigo(g_empresa, Val(dtcbo_fornecedor.BoundText)) Then
        MsgBox "Fornecedor não cadastrado, selecione outro fornecedor.", vbInformation, "Mensagem do Sistema!"
        Exit Function
    End If
    If Not Contas.LocalizarCodigo(Fornecedor.CodigoConta) Then
        MsgBox "Conta não cadastrada.", vbInformation, "Mensagem do Sistema!"
        Exit Function
    End If
    If cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", vbInformation, "Mensagem do Sistema!"
    ElseIf cbo_conta.ItemData(cbo_conta.ListIndex) <> Fornecedor.CodigoConta Then
        If (MsgBox("A conta correta para este fornecedor é.: " & Chr(34) & Contas.Nome & Chr(34) & "." & Chr(10) & Chr(13) & "Porém, a conta selecionada foi.: " & Chr(34) & cbo_conta & Chr(34) & "." & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "Deseja mudar para a conta correta?", vbYesNo + vbDefaultButton1 + vbQuestion, "Mensagem do Sistema!")) = 7 Then
            If (MsgBox("Caso tenha alguma dúvida deste lançamento consulte a diretoria." & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "Tem certeza que este lançamento é da conta.: " & Chr(34) & cbo_conta & Chr(34) & "." & Chr(10) & Chr(13), vbYesNo + vbDefaultButton2 + vbQuestion, "Mensagem do Sistema!")) = 6 Then
                ValidaConta = True
            End If
        End If
    Else
        ValidaConta = True
    End If
End Function
Function ValidaPeriodo() As Boolean
    ValidaPeriodo = True
    If cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        ValidaPeriodo = False
    ElseIf cbo_periodo.ListIndex = 3 And cbo_tipo_movimento.ListIndex <> 1 Then
        If (MsgBox("Tipo de movimento não aceito para este período." & Chr(10) & Chr(10) & "Deseja cadastrar mesmo assim?", vbYesNo + vbDefaultButton2, "Atenção!")) = 7 Then
            ValidaPeriodo = False
        End If
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_despesa_caixa.Show 1
    If Len(g_string) > 0 Then
        lRegistro = RetiraGString(1)
        Call MovDespesaCaixa.LocalizarCodigo(g_empresa, lRegistro)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovDespesaCaixa.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovDespesaCaixa.LocalizarProximo Then
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
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovDespesaCaixa.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_valor.SetFocus
    End If
End Sub
Private Sub dtcbo_fornecedor_LostFocus()
    If Val(dtcbo_fornecedor.BoundText) > 0 And lOpcao = 1 Then
        BuscaUltimoLancamento
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    'msk_data_movimento.ZOrder 1
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaFornecedor
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        Else
            lCxPeriodo = 0
            If MovDespesaCaixa.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtivaBotoes
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub AtualizaFornecedor()
    Set adodc_fornecedor.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
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
    Screen.MousePointer = 1
    CentraForm Me
    
    PreencheCboConta
    PreencheCboIlha
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    lGravados = 0
    lCxPeriodo = 0
    lCxCodigoLancamentoPadrao = 0
    lCxTipoMovimento = 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Caixa da Conveniência"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub PreencheCboConta()
    Dim xSQL As String
    
    cbo_conta.Clear
    'Prepara SQL
    xSQL = "SELECT Nome, Codigo"
    xSQL = xSQL & "  FROM Contas"
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica tabela
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_conta.AddItem rsTabela("Nome").Value
            cbo_conta.ItemData(cbo_conta.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub msk_data_movimento_GotFocus()
    If lOpcao = 1 And Not IsDate(msk_data_movimento) Then
        msk_data_movimento = Format(g_data_def, "dd/mm/yyyy")
    End If
    msk_data_movimento.SelStart = 0
    msk_data_movimento.SelLength = 2
End Sub
Private Sub msk_data_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub txt_complemento_GotFocus()
    txt_complemento.SelStart = 0
    txt_complemento.SelLength = Len(txt_complemento.Text)
End Sub
Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_numero_documento_GotFocus()
    txt_numero_documento.SelStart = 0
    txt_numero_documento.SelLength = Len(txt_numero_documento.Text)
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_complemento.SetFocus
    End If
End Sub
Private Sub txt_registro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_fornecedor.SetFocus
    End If
End Sub
Private Sub txt_registro_LostFocus()
    If lOpcao = 1 And Val(txt_registro.Text) > 0 Then
        If MovDespesaCaixa.LocalizarCodigo(g_empresa, Val(txt_registro.Text)) Then
            MsgBox "Já existe registro cadastrado com este número." & Chr(10) & Chr(10) & "Mude o número informado.", vbInformation, "Duplicidade de Registro!"
            txt_registro.SetFocus
            Exit Sub
        End If
        If MovContaPagar.LocalizarCodigo(g_empresa, Val(txt_registro.Text)) Then
            MsgBox "Já existe registro de contas à pagar cadastrado com este número." & Chr(10) & Chr(10) & "Mude o número informado.", vbInformation, "Duplicidade de Registro!"
            txt_registro.SetFocus
            Exit Sub
        End If
        If BaixaPagar.LocalizarCodigo(g_empresa, Val(txt_registro.Text)) Then
            MsgBox "Já existe registro baixado cadastrado com este número." & Chr(10) & Chr(10) & "Mude o número informado.", vbInformation, "Duplicidade de Registro!"
            txt_registro.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor.Text)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_documento.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor.Text = Format(txt_valor.Text, "###,##0.00")
End Sub
Private Sub ProcessaDespesaDeCaixa()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovDespesaCaixa.LocalizarPrimeiro Then
        If MovDespesaCaixa.DataMovimento >= xData Then
            AtualTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovDespesaCaixa.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
                MovDespesaCaixa.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
                If Not MovDespesaCaixa.Alterar(g_empresa, Val(txt_registro.Text)) Then
                    MsgBox "Erro ao alterar despesa de caixa", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovDespesaCaixa.LocalizarProximo = False
            If MovDespesaCaixa.DataMovimento >= xData Then
                AtualTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovDespesaCaixa.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
                    MovDespesaCaixa.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
                    If Not MovDespesaCaixa.Alterar(g_empresa, Val(txt_registro.Text)) Then
                        MsgBox "Erro ao alterar despesa de caixa", vbInformation, "Erro"
                    End If
                End If
            End If
        Loop
    
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    MsgBox "Processamento Concluído!"
    Exit Sub
FileError:
    MsgBox "Erro ao processar Despesa de Caixa", vbInformation, "ProcessaDespesaDeCaixa"
End Sub


