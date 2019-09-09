VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form mov_contas_pagar 
   Caption         =   "Movimento de Contas à Pagar"
   ClientHeight    =   4695
   ClientLeft      =   2250
   ClientTop       =   1530
   ClientWidth     =   7035
   Icon            =   "MOV_CONTAS_PAGAR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "MOV_CONTAS_PAGAR.frx":030A
   ScaleHeight     =   4695
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "MOV_CONTAS_PAGAR.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "MOV_CONTAS_PAGAR.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "MOV_CONTAS_PAGAR.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "MOV_CONTAS_PAGAR.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Altera o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "MOV_CONTAS_PAGAR.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cria um novo registro."
      Top             =   3720
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.CommandButton cmd_data_e 
         Height          =   315
         Left            =   2880
         Picture         =   "MOV_CONTAS_PAGAR.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_v 
         Height          =   315
         Left            =   2880
         Picture         =   "MOV_CONTAS_PAGAR.frx":874C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbo_local 
         Height          =   315
         ItemData        =   "MOV_CONTAS_PAGAR.frx":9A26
         Left            =   1740
         List            =   "MOV_CONTAS_PAGAR.frx":9A28
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   4935
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
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   22
         Top             =   3120
         Width           =   4935
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   285
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2040
         Width           =   1035
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   315
         ItemData        =   "MOV_CONTAS_PAGAR.frx":9A2A
         Left            =   1740
         List            =   "MOV_CONTAS_PAGAR.frx":9A2C
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
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
      Begin MSMask.MaskEdBox msk_data_vencimento 
         Height          =   300
         Left            =   1740
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   255
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
      Begin MSMask.MaskEdBox msk_data_emissao 
         Height          =   300
         Left            =   1740
         TabIndex        =   8
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcFornecedor 
         Height          =   330
         Left            =   3120
         Top             =   600
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
      Begin MSDataListLib.DataCombo dtcboFornecedor 
         Bindings        =   "MOV_CONTAS_PAGAR.frx":9A2E
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFornecedor"
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
      Begin VB.Label Label3 
         Caption         =   "Local de Cobrança"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Valor do Vencimento"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Data do Vencimento"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fornecedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Complemento"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Número Documento"
         Height          =   315
         Left            =   120
         TabIndex        =   15
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
         Caption         =   "Conta"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   30
      Top             =   3600
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "MOV_CONTAS_PAGAR.frx":9A4C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "MOV_CONTAS_PAGAR.frx":AF46
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "MOV_CONTAS_PAGAR.frx":C440
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "MOV_CONTAS_PAGAR.frx":D8B2
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
      Left            =   5220
      Picture         =   "MOV_CONTAS_PAGAR.frx":EE34
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "MOV_CONTAS_PAGAR.frx":1043E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3720
      Width           =   795
   End
End
Attribute VB_Name = "mov_contas_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lEmpresa As Integer
Dim lOpcao As Integer
Dim lRegistro As Long
Dim lData As Date
Dim lNumeroMovimentoCaixa As Long
Dim lIntegraFinanceiroDiario As Boolean
Dim lAlterarContaFornecedor As Boolean

Private rsTabela As New ADODB.Recordset

Private Contas As New cContas
Private IntegracaoCaixa As New cIntegracaoCaixa
Private Fornecedor As New cFornecedor
Private LocalCobranca As New cLocalCobranca
Private MovCaixa As New cMovimentoCaixa
Private MovContaPagar As New cMovimentoContaPagar
Private ConfigDiversa As New cConfiguracaoDiversa

Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    lRegistro = Val(txt_registro.Text)
    MovContaPagar.Empresa = g_empresa
    MovContaPagar.Registro = lRegistro
    MovContaPagar.CodigoFornecedor = Val(dtcboFornecedor.BoundText)
    MovContaPagar.NomeFornecedor = dtcboFornecedor
    MovContaPagar.DataEmissao = msk_data_emissao.Text
    MovContaPagar.DataVencimento = msk_data_vencimento.Text
    MovContaPagar.Valor = fValidaValor2(txt_valor.Text)
    MovContaPagar.NumeroDocumento = txt_numero_documento.Text
    MovContaPagar.LocalCobranca = cbo_local.ItemData(cbo_local.ListIndex)
    MovContaPagar.CodigoConta = cbo_conta.ItemData(cbo_conta.ListIndex)
    MovContaPagar.Complemento = txt_complemento.Text
    MovContaPagar.DataDigitacao = CDate(lbl_data_digitacao.Caption)
    MovContaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
End Sub
Private Sub BuscaUltimoLancamento()
    Dim i As Integer
    Dim xLocalCobranca As Integer
    Dim xCodigoConta As Integer
    g_string = MovContaPagar.LocalizaCobrancaConta(g_empresa, Val(dtcboFornecedor.BoundText))
    If g_string <> "" Then
        xLocalCobranca = RetiraGString(1)
        xCodigoConta = RetiraGString(2)
        For i = 0 To cbo_local.ListCount - 1
            cbo_local.ListIndex = i
            If cbo_local.ItemData(i) = xLocalCobranca Then
                Exit For
            End If
        Next
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(i) = xCodigoConta Then
                Exit For
            End If
        Next
    End If
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lRegistro = MovContaPagar.Registro
    lData = MovContaPagar.DataEmissao
    lNumeroMovimentoCaixa = MovContaPagar.NumeroMovimentoCaixa
    txt_registro.Text = MovContaPagar.Registro
    lbl_data_digitacao.Caption = Format(MovContaPagar.DataDigitacao, "dd/mm/yyyy")
    dtcboFornecedor.BoundText = MovContaPagar.CodigoFornecedor
    msk_data_emissao.Text = Format(MovContaPagar.DataEmissao, "dd/mm/yyyy")
    msk_data_vencimento.Text = Format(MovContaPagar.DataVencimento, "dd/mm/yyyy")
    txt_valor.Text = Format(MovContaPagar.Valor, "###,##0.00")
    txt_numero_documento.Text = MovContaPagar.NumeroDocumento
    If LocalCobranca.LocalizarCodigo(MovContaPagar.LocalCobranca, MovContaPagar.Empresa) Then
        For i = 0 To cbo_local.ListCount - 1
            cbo_local.ListIndex = i
            If cbo_local.ItemData(i) = LocalCobranca.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_local.ListIndex = -1
    End If
    If Contas.LocalizarCodigo(MovContaPagar.CodigoConta, MovContaPagar.Empresa) Then
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(i) = Contas.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_conta.ListIndex = -1
    End If
    txt_complemento.Text = MovContaPagar.Complemento
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
    If Not MovCaixa.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
        MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub Finaliza()
    Set Contas = Nothing
    Set IntegracaoCaixa = Nothing
    Set Fornecedor = Nothing
    Set LocalCobranca = Nothing
    Set MovCaixa = Nothing
    Set MovContaPagar = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_registro.Text = Format(MovContaPagar.ProximoRegistro(g_empresa), "###,##0")
    lbl_data_digitacao = Format(g_data_def, "dd/mm/yyyy")
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xNome As String
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xCodigoHistorico As Integer
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    xNome = ""
    If Fornecedor.CodigoConta = 1 Then
        xNome = "-ESTOQUE"
    End If
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "CONTAS A PAGAR" & xNome) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        xCodigoHistorico = IntegracaoCaixa.HistoricoPadrao
    Else
        xContaDebito = "421030028"
        xContaCredito = "221050003"
        xCodigoHistorico = 18
    End If
    xComplemento = Trim(txt_complemento.Text) '& " TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    
    
    
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(msk_data_emissao.Text)
    MovCaixa.NumeroMovimento = 1
    MovCaixa.Valor = fValidaValor(txt_valor.Text)
    MovCaixa.NumeroDocumento = txt_numero_documento.Text
    MovCaixa.CodigoHistorico = xCodigoHistorico
    MovCaixa.Complemento = xComplemento
    MovCaixa.NumeroContaDebito = xContaDebito
    MovCaixa.NumeroContaCredito = xContaCredito
    If Len(Fornecedor.ContaContabil) = 9 Then
        MovCaixa.NumeroContaCredito = Fornecedor.ContaContabil
    End If
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = True
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixa = True
        lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
    Else
        MsgBox "Não foi integrado no caixa o valor=" & txt_valor.Text, vbInformation, "Erro de Integridade"
    End If
End Function
Private Sub cbo_conta_GotFocus()
    SendMessageLong cbo_conta.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_complemento.SetFocus
    End If
End Sub
Private Sub cbo_local_GotFocus()
    SendMessageLong cbo_local.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_local_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_conta.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    dtcboFornecedor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovContaPagar.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If MovContaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
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
    txt_registro = ""
    lbl_data_digitacao = ""
    dtcboFornecedor.BoundText = ""
    msk_data_emissao = "__/__/____"
    msk_data_vencimento = "__/__/____"
    txt_valor = ""
    txt_numero_documento = ""
    cbo_local.ListIndex = -1
    cbo_conta.ListIndex = -1
    txt_complemento = ""
End Sub
Private Sub cmd_data_e_Click()
    g_string = msk_data_emissao
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_emissao = RetiraGString(1)
        msk_data_vencimento = RetiraGString(2)
        txt_valor.SetFocus
    Else
        msk_data_emissao = RetiraGString(1)
        msk_data_vencimento.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_v_Click()
    g_string = msk_data_emissao
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_emissao = RetiraGString(1)
        msk_data_vencimento = RetiraGString(2)
    Else
        msk_data_vencimento = RetiraGString(1)
    End If
    g_string = " "
    txt_valor.SetFocus
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_registro.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call ExcluiMovimentoCaixa
            If MovContaPagar.Excluir(g_empresa, CLng(txt_registro.Text)) Then
                LimpaTela
                If MovContaPagar.LocalizarUltimo(g_empresa) Then
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
End Sub
Private Sub cmd_novo_Click()
    '
    'bd_sgp.Execute "UpDate Contas_Pagar Set [Data da Digitacao] = data_emissao"
    '
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    'atribui o valor de g_empresa para lEmpresa
    'caso o usuário cria novo registro e muda de
    'empresa sem cancelar o sistema o avisará
    lEmpresa = g_empresa
    dtcboFornecedor.SetFocus
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaContasPagar
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
            If MovContaPagar.Incluir Then
                lRegistro = MovContaPagar.Registro
                If lAlterarContaFornecedor Then
                    If Fornecedor.LocalizarCodigo(g_empresa, MovContaPagar.CodigoFornecedor) Then
                        Fornecedor.CodigoConta = MovContaPagar.CodigoConta
                        If Not Fornecedor.Alterar(g_empresa, Fornecedor.Codigo) Then
                            MsgBox "Não foi possível alterar conta do Fornecedor!", vbInformation, "Erro de Integridade."
                        End If
                    End If
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            Call ExcluiMovimentoCaixa
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            If MovContaPagar.Alterar(g_empresa, lRegistro) Then
                lRegistro = MovContaPagar.Registro
                
                If lAlterarContaFornecedor Then
                    If Fornecedor.LocalizarCodigo(g_empresa, MovContaPagar.CodigoFornecedor) Then
                        Fornecedor.CodigoConta = MovContaPagar.CodigoConta
                        If Not Fornecedor.Alterar(g_empresa, Fornecedor.Codigo) Then
                            MsgBox "Não foi possível alterar conta do Fornecedor!", vbInformation, "Erro de Integridade."
                        End If
                    End If
                End If
                
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If MovContaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Private Sub ProcessaContasPagar()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovContaPagar.LocalizarPrimeiro Then
        If MovContaPagar.DataEmissao >= xData Then
            AtualTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovContaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If Not MovContaPagar.Alterar(g_empresa, lRegistro) Then
                    MsgBox "Erro ao alterar contas a pagar", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovContaPagar.LocalizarProximo = False
            If MovContaPagar.DataEmissao >= xData Then
                AtualTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovContaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                    If Not MovContaPagar.Alterar(g_empresa, lRegistro) Then
                        MsgBox "Erro ao alterar contas a pagar", vbInformation, "Erro"
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
    MsgBox "Erro ao processar Contas à Pagar", vbInformation, "ProcessaContasPagar"
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_registro) > 0 Then
        MsgBox "Informe o número do registro.", 64, "Atenção!"
        txt_registro.SetFocus
    ElseIf Not Val(dtcboFornecedor.BoundText) > 0 Then
        MsgBox "Selecione o fornecedor.", 64, "Atenção!"
        dtcboFornecedor.SetFocus
    ElseIf Not IsDate(msk_data_emissao) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data_emissao.SetFocus
    ElseIf Not IsDate(msk_data_vencimento) Then
        MsgBox "Informe a data de vencimento.", 64, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf Not fValidaValor2(txt_valor) > 0 Then
        MsgBox "Informe o valor do vencimento.", 64, "Atenção!"
        txt_valor.SetFocus
    ElseIf cbo_local.ListIndex = -1 Then
        MsgBox "Selecione o local de cobrança.", 64, "Atenção!"
        cbo_local.SetFocus
    ElseIf Not ValidaConta Then
        cbo_conta.SetFocus
    ElseIf CDate(msk_data_vencimento) < CDate(msk_data_emissao) Then
        MsgBox "Data de vencimento deve ser maior ou igual a " & msk_data_emissao.Text & ".", 64, "Atenção!"
        msk_data_vencimento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaConta() As Integer
    ValidaConta = False
    If Not Fornecedor.LocalizarCodigo(g_empresa, Val(dtcboFornecedor.BoundText)) Then
        MsgBox "Fornecedor não cadastrado, selecione outro fornecedor.", vbInformation, "Mensagem do Sistema!"
        Exit Function
    End If
    If Not Contas.LocalizarCodigo(Fornecedor.CodigoConta, Fornecedor.Empresa) Then
        MsgBox "Conta não cadastrada.", vbInformation, "Mensagem do Sistema!"
        Exit Function
    End If
    If cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", vbInformation, "Mensagem do Sistema!"
    ElseIf cbo_conta.ItemData(cbo_conta.ListIndex) <> Fornecedor.CodigoConta Then
        If (MsgBox("A conta correta para este fornecedor é.: " & Chr(34) & Contas.Nome & Chr(34) & "." & Chr(10) & Chr(13) & "Porém, a conta selecionada foi.: " & Chr(34) & cbo_conta & Chr(34) & "." & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "Deseja mudar para a conta correta?", vbYesNo + vbDefaultButton1 + vbQuestion, "Mensagem do Sistema!")) = 7 Then
            If lIntegraFinanceiroDiario Then
                ValidaConta = True
                If (MsgBox("Deseja alterar a conta fornecedor de " & Contas.Nome & " para " & cbo_conta & "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Mensagem do Sistema!")) = 6 Then
                    lAlterarContaFornecedor = True
                Else
                    lAlterarContaFornecedor = False
                End If
            Else
                If (MsgBox("Caso tenha alguma dúvida deste lançamento consulte a diretoria." & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "Tem certeza que este lançamento é da conta.: " & Chr(34) & cbo_conta & Chr(34) & "." & Chr(10) & Chr(13), vbYesNo + vbDefaultButton2 + vbQuestion, "Mensagem do Sistema!")) = 6 Then
                    ValidaConta = True
                End If
            End If
        End If
    Else
        ValidaConta = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_contas_pagar.Show 1
    If Len(g_string) > 0 Then
        lRegistro = RetiraGString(1)
        If MovContaPagar.LocalizarCodigo(g_empresa, lRegistro) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovContaPagar.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovContaPagar.LocalizarProximo Then
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
    If MovContaPagar.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbofornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        msk_data_emissao.SetFocus
    End If
End Sub
Private Sub dtcbofornecedor_LostFocus()
    If Val(dtcboFornecedor.BoundText) > 0 And lOpcao = 1 Then
        BuscaUltimoLancamento
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_emissao.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If Val(dtcboFornecedor.BoundText) > 0 And lOpcao = 1 Then
        BuscaUltimoLancamento
    End If
End Sub
Private Sub Form_Activate()
    msk_data_emissao.ZOrder 1
    msk_data_vencimento.ZOrder 1
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaFornecedor
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If MovContaPagar.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub AtualizaFornecedor()
    Set adodcFornecedor.Recordset = Conectar.RsConexao("Select * From Fornecedor Where Empresa = " & g_empresa & " Order By Nome")
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
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
    Screen.MousePointer = 1
    CentraForm Me
    
    lIntegraFinanceiroDiario = False
    lAlterarContaFornecedor = False
    
    If ConfigDiversa.LocalizarCodigo(1, "CONTAS A PAGAR:INTEGRADO COM FINANCEIRO") Then
       lIntegraFinanceiroDiario = ConfigDiversa.Verdadeiro
    End If
    
    
    PreencheCboLocal
    PreencheCboConta
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub PreencheCboLocal()
    Dim xSQL As String
    
    cbo_local.Clear
    'Prepara SQL
    xSQL = "SELECT Nome, Codigo"
    xSQL = xSQL & "  FROM Local_Cobrancas WHERE Empresa = " & g_empresa
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica tabela
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_local.AddItem rsTabela("Nome").Value
            cbo_local.ItemData(cbo_local.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub PreencheCboConta()
    Dim xSQL As String
    
    cbo_conta.Clear
    'Prepara SQL
    xSQL = "SELECT Nome, Codigo"
    xSQL = xSQL & "  FROM Contas"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
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
Private Sub msk_data_emissao_GotFocus()
    If lOpcao = 1 And Not IsDate(msk_data_emissao) Then
        msk_data_emissao = "__/__/" & Format(g_data_def, "yyyy")
    End If
End Sub
Private Sub msk_data_emissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_vencimento.SetFocus
    End If
End Sub
Private Sub msk_data_vencimento_GotFocus()
    If lOpcao = 1 And Not IsDate(msk_data_vencimento) Then
        msk_data_vencimento = "__/__/" & Format(g_data_def, "yyyy")
    End If
End Sub
Private Sub msk_data_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_local.SetFocus
    End If
End Sub
Private Sub txt_registro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFornecedor.SetFocus
    End If
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
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
