VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_cheque_devolvido 
   Caption         =   "Movimentação de Cheques Devolvidos"
   ClientHeight    =   6780
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "movimento_cheque_devolvido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cheque_devolvido.frx":030A
   ScaleHeight     =   6780
   ScaleWidth      =   6975
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   5715
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtCnpjCpf 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   22
         Top             =   4020
         Width           =   2655
      End
      Begin VB.TextBox txt_motivo_devolucao 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   26
         Top             =   4860
         Width           =   2655
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txt_emitente 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   20
         Top             =   3600
         Width           =   4935
      End
      Begin VB.TextBox txt_cheque 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   18
         Top             =   3180
         Width           =   735
      End
      Begin VB.TextBox txt_conta 
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox msk_valor 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1500
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_vencimento 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_emissao 
         Height          =   315
         Left            =   5520
         TabIndex        =   4
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
      Begin MSMask.MaskEdBox msk_data_devolucao 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   4440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodc_banco 
         Height          =   330
         Left            =   3600
         Top             =   2340
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
         Caption         =   "adodc_banco"
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
      Begin MSDataListLib.DataCombo dtcbo_banco 
         Bindings        =   "movimento_cheque_devolvido.frx":0750
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   2340
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_banco"
      End
      Begin MSAdodcLib.Adodc adodcSituacao 
         Height          =   330
         Left            =   2700
         Top             =   5280
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcSituacao"
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
      Begin MSDataListLib.DataCombo dtcboSituacao 
         Bindings        =   "movimento_cheque_devolvido.frx":076A
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   5280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboSituacao"
      End
      Begin MSMask.MaskEdBox msk_data_digitacao 
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
      Begin VB.Label Label15 
         Caption         =   "CNP&J / CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "&Data de Digitação"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Situação"
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "&Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "&Data da Devolução"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Motivo Devolução"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "No&me do Emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "D&ata do Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Nú&mero do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da Conta"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "&Tipo de Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Período"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   38
      Top             =   5700
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_cheque_devolvido.frx":0786
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cheque_devolvido.frx":1C80
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cheque_devolvido.frx":317A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cheque_devolvido.frx":45EC
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cheque_devolvido.frx":5B6E
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_cheque_devolvido.frx":7178
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cheque_devolvido.frx":8672
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_cheque_devolvido.frx":9D04
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_cheque_devolvido.frx":B176
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_cheque_devolvido.frx":C808
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Altera o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cheque_devolvido.frx":DD02
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Cria um novo registro."
      Top             =   5820
      Width           =   795
   End
End
Attribute VB_Name = "movimento_cheque_devolvido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lDataDigitacao As Date
Dim lDataEmissao As Date
Dim lPeriodo As String
Dim lTipoMovimento As String
Dim lOrdem As Integer
Dim lConta As String
Dim lCheque As String
Dim lTotal As Currency
Dim lGravados As Integer
Dim lDados As String
Dim lCodigoBarra1 As String
Dim lCodigoBarra2 As String
Dim lCodigoBarra3 As String
Dim lQtdPeriodo As Integer
Private MovBaixaCheque As New cMovimentoBaixaCheque
Private MovChequeDevolvido As New cMovimentoChequeDevolvido
Private Sub AtualTabe()
    MovChequeDevolvido.Empresa = g_empresa
    MovChequeDevolvido.DataDigitacao = msk_data_digitacao.Text
    MovChequeDevolvido.DataEmissao = msk_data_emissao.Text
    MovChequeDevolvido.CodigoBanco = Val(dtcbo_banco.BoundText)
    MovChequeDevolvido.NumeroConta = Val(txt_conta.Text)
    MovChequeDevolvido.NumeroCheque = Val(txt_cheque.Text)
    MovChequeDevolvido.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    MovChequeDevolvido.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovChequeDevolvido.valor = fValidaValor2(msk_valor.Text)
    MovChequeDevolvido.DataVencimento = msk_data_vencimento.Text
    MovChequeDevolvido.Emitente = txt_emitente.Text
    MovChequeDevolvido.OrdemDigitacao = lOrdem
    MovChequeDevolvido.CodigoBarra1 = lCodigoBarra1
    MovChequeDevolvido.CodigoBarra2 = lCodigoBarra2
    MovChequeDevolvido.CodigoBarra3 = lCodigoBarra3
    MovChequeDevolvido.BancoAgencia = Mid(lCodigoBarra3, 1, 7)
    MovChequeDevolvido.DataDevolucao = msk_data_devolucao.Text
    MovChequeDevolvido.MotivoDevolucao = txt_motivo_devolucao.Text
    MovChequeDevolvido.Situacao = Val(dtcboSituacao.BoundText)
    MovChequeDevolvido.CnpjCpf = txtCnpjCpf.Text
End Sub
Private Sub MostraDadosInicial()
    Dim i As Integer
    msk_data_digitacao.Text = lDataDigitacao
    msk_data_emissao.Text = lDataEmissao
    cbo_periodo = lPeriodo
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = lTipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
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
    cbo_tipo_movimento.AddItem "1 - Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de Óleos/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Cheque Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lDataDigitacao = MovChequeDevolvido.DataDigitacao
    lDataEmissao = MovChequeDevolvido.DataEmissao
    lPeriodo = MovChequeDevolvido.Periodo
    lTipoMovimento = MovChequeDevolvido.TipoMovimento
    lOrdem = MovChequeDevolvido.OrdemDigitacao
    lConta = MovChequeDevolvido.NumeroConta
    lCheque = MovChequeDevolvido.NumeroCheque
    lCodigoBarra1 = MovChequeDevolvido.CodigoBarra1
    lCodigoBarra2 = MovChequeDevolvido.CodigoBarra2
    lCodigoBarra3 = MovChequeDevolvido.CodigoBarra3
    msk_data_digitacao.Text = Format(MovChequeDevolvido.DataDigitacao, "dd/mm/yyyy")
    msk_data_emissao.Text = Format(MovChequeDevolvido.DataEmissao, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovChequeDevolvido.Periodo - 1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovChequeDevolvido.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    dtcbo_banco.BoundText = ""
    dtcbo_banco.BoundText = MovChequeDevolvido.CodigoBanco
    txt_conta.Text = MovChequeDevolvido.NumeroConta
    txt_cheque.Text = MovChequeDevolvido.NumeroCheque
    msk_valor.Text = Format(MovChequeDevolvido.valor, "###,##0.00")
    msk_data_vencimento.Text = Format(MovChequeDevolvido.DataVencimento, "dd/mm/yyyy")
    txt_emitente.Text = MovChequeDevolvido.Emitente
    txtCnpjCpf.Text = MovChequeDevolvido.CnpjCpf
    msk_data_devolucao.Text = Format(MovChequeDevolvido.DataDevolucao, "dd/mm/yyyy")
    txt_motivo_devolucao.Text = MovChequeDevolvido.MotivoDevolucao
    lbl_total.Caption = Format(MovChequeDevolvido.TotalPeriodo(g_empresa, CDate(msk_data_emissao.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text)), "###,##0.00")
    dtcboSituacao.BoundText = ""
    dtcboSituacao.BoundText = MovChequeDevolvido.Situacao
    frm_dados.Enabled = False
End Sub
Private Sub AtualTelaBaixa()
    Dim i As Integer
    msk_data_digitacao.Text = Format(MovBaixaCheque.DataEmissao, "dd/mm/yyyy")
    msk_data_emissao.Text = Format(MovBaixaCheque.DataEmissao, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovBaixaCheque.Periodo - 1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovBaixaCheque.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    txt_conta.Text = MovBaixaCheque.NumeroConta
    txt_cheque.Text = MovBaixaCheque.NumeroCheque
    msk_valor.Text = Format(MovBaixaCheque.valor, "###,##0.00")
    msk_data_vencimento.Text = Format(MovBaixaCheque.DataVencimento, "dd/mm/yyyy")
    txt_emitente.Text = MovBaixaCheque.Emitente
    'msk_data_devolucao.Text = Format(MovBaixaCheque.DataDevolucao, "dd/mm/yyyy")
    'txt_motivo_devolucao.Text = MovBaixaCheque.MotivoDevolucao
    'lbl_total.Caption = Format(MovBaixaCheque.TotalPeriodo(g_empresa, CDate(msk_data_emissao.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text)), "###,##0.00")
    'frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set MovBaixaCheque = Nothing
    Set MovChequeDevolvido = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
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
    msk_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovChequeDevolvido.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If MovChequeDevolvido.LocalizarRegistro(g_empresa, lDataEmissao, lConta, lCheque) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LeituraCheque()
Dim x As String
    abre_porta
    x = DRCarrega
    If x = 4 Then
        MsgBox "Cheque Não Inserido!"
    ElseIf x = 1 Then
        Open "\VB5\SGP\DATA\DR10.RET" For Input As #1
        Line Input #1, lDados
        Close #1
        txt_conta = Mid(lDados, 25, 8)
        txt_cheque = Mid(lDados, 14, 6)
        lCodigoBarra1 = Mid(lDados, 2, 8)
        lCodigoBarra2 = Mid(lDados, 11, 10)
        lCodigoBarra3 = Mid(lDados, 22, 12)
    Else
        MsgBox "Erro não identificado! " & x
    End If
    fechar_porta
End Sub
Private Sub LimpaTela()
    If lGravados = 0 Then
        msk_data_digitacao.Text = "__/__/____"
        msk_data_emissao.Text = "__/__/____"
        cbo_periodo.ListIndex = -1
        cbo_tipo_movimento.ListIndex = -1
    End If
    dtcbo_banco.BoundText = ""
    txt_conta.Text = ""
    txt_cheque.Text = ""
    msk_valor.Text = ""
    msk_data_vencimento.Text = "__/__/____"
    txt_emitente.Text = ""
    txtCnpjCpf.Text = ""
    msk_data_devolucao.Text = "__/__/____"
    txt_motivo_devolucao.Text = ""
    dtcboSituacao.BoundText = ""
    lCodigoBarra1 = "00000000"
    lCodigoBarra2 = "0000000000"
    lCodigoBarra3 = "000000000000"
End Sub
Private Sub cmd_excluir_Click()
    If lConta <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If MovChequeDevolvido.Excluir(g_empresa, lDataEmissao, lConta, lCheque) Then
                LimpaTela
                If MovChequeDevolvido.LocalizarUltimo(g_empresa) Then
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
    'ZZGambiArrumaOrdemDigitacao
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    'If lGravados = 0 Then
    '    If BuscaProximoCaixa Then
    '        msk_valor.SetFocus
    '    Else
    '        msk_data_emissao.SetFocus
    '    End If
    'Else
    '    msk_valor.SetFocus
    'End If
    consulta_cheque_baixado.Show 1
    If Len(g_string) > 0 Then
        lDataEmissao = RetiraGString(1)
        lConta = RetiraGString(2)
        lCheque = RetiraGString(3)
        If MovBaixaCheque.LocalizarRegistro(g_empresa, lDataEmissao, lConta, lCheque) Then
            AtualTelaBaixa
            msk_data_devolucao.SetFocus
            Exit Sub
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
    msk_data_digitacao.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                lOrdem = MovChequeDevolvido.LocalizarOrdemDigitacao(g_empresa, CDate(msk_data_emissao), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text)) + 1
                lGravados = 1
                AtualTabe
                If MovChequeDevolvido.Incluir Then
                    lDataEmissao = msk_data_emissao.Text
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lOrdem = lOrdem
                    lConta = Val(txt_conta.Text)
                    lCheque = Val(txt_cheque.Text)
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
            ElseIf lOpcao = 2 Then
                AtualTabe
                If MovChequeDevolvido.Alterar(g_empresa, lDataEmissao, lConta, lCheque) Then
                    lDataEmissao = msk_data_emissao.Text
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lOrdem = lOrdem
                    lConta = Val(txt_conta.Text)
                    lCheque = Val(txt_cheque.Text)
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
            End If
            If MovChequeDevolvido.LocalizarRegistro(g_empresa, lDataEmissao, lConta, lCheque) Then
                AtualTela
            Else
                LimpaTela
                MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
            End If
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    Dim dias As Integer
    ValidaCampos = False
    If IsDate(msk_data_emissao.Text) And IsDate(msk_data_vencimento.Text) Then
        dias = DateDiff("d", msk_data_emissao.Text, msk_data_vencimento.Text)
    End If
    If Not IsDate(msk_data_digitacao.Text) Then
        MsgBox "Informe a data de digitação.", vbInformation, "Atenção!"
        msk_data_digitacao.SetFocus
    ElseIf Not IsDate(msk_data_emissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data_emissao.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Informe o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf Not Val(dtcbo_banco.BoundText) > 0 Then
        MsgBox "Selecione um banco.", vbInformation, "Atenção!"
        dtcbo_banco.SetFocus
    ElseIf Not Val(txt_conta.Text) > 0 Then
        MsgBox "Informe o número da conta.", vbInformation, "Atenção!"
        txt_conta.SetFocus
    ElseIf Not Val(txt_cheque.Text) > 0 Then
        MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
        txt_cheque.SetFocus
    ElseIf Not fValidaValor2(msk_valor.Text) > 0 Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Atenção!"
        msk_valor.SetFocus
    ElseIf Not IsDate(msk_data_vencimento.Text) Then
        MsgBox "Informe a data de vencimento.", vbInformation, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf CDate(msk_data_vencimento.Text) < CDate(msk_data_emissao.Text) Then
        MsgBox "Data de vencimento deve ser maior ou igual a " & msk_data_emissao.Text & ".", vbInformation, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf Not txt_emitente.Text <> "" Then
        MsgBox "Informe o nome do emitente.", vbInformation, "Atenção!"
        txt_emitente.SetFocus
    ElseIf Not (Len(txtCnpjCpf.Text) = 11 Or Len(txtCnpjCpf.Text) = 14) Then
        MsgBox "Informe corretamente o CNPJ ou CPF do emitente.", vbInformation, "Atenção!"
        txtCnpjCpf.SetFocus
    ElseIf Not IsDate(msk_data_devolucao.Text) Then
        MsgBox "Informe a data da devolução.", vbInformation, "Atenção!"
        msk_data_devolucao.SetFocus
    ElseIf Not txt_motivo_devolucao.Text <> "" Then
        MsgBox "Informe o motivo da devolução.", vbInformation, "Atenção!"
        txt_motivo_devolucao.SetFocus
    ElseIf Not Val(dtcboSituacao.BoundText) > 0 Then
        MsgBox "Selecione uma Situação.", vbInformation, "Atenção!"
        dtcboSituacao.SetFocus
    ElseIf dias > 60 Then
        If MsgBox("Este cheque está com mais de 60 dias de prazo." & Chr(13) & "Cheque com " & dias & " dia(s)." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Confirma mesmo assim?", 292, "Prazo de Cheque Incorreto!") = 7 Then
            msk_data_vencimento.SetFocus
        Else
            ValidaCampos = True
        End If
    'ElseIf dias < 5 And dias > 0 Then
    '    If MsgBox("Este cheque está com menos 5 dias de prazo." & Chr(13) & "Cheque com " & dias & " dia(s)." & Chr(13) & Chr(13) & Chr(13) & "Mude para: " & CDate(msk_data_emissao) + 5 & Chr(13) & Chr(13) & Chr(13) & "Confirma mesmo assim?", 292, "Prazo de Cheque Incorreto!") = 7 Then
    '        msk_data_vencimento.SetFocus
    '    Else
    '        ValidaCampos = True
    '    End If
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaCodigoBarra() As Boolean
    Dim i As Integer
    ValidaCodigoBarra = True
    If Len(lCodigoBarra1) <> 8 Or Len(lCodigoBarra2) <> 10 Or Len(lCodigoBarra3) <> 12 Then
        ValidaCodigoBarra = False
        Exit Function
    End If
    For i = 1 To 8
        If Asc(Mid(lCodigoBarra1, i, 1)) < 48 Or Asc(Mid(lCodigoBarra1, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
    For i = 1 To 10
        If Asc(Mid(lCodigoBarra2, i, 1)) < 48 Or Asc(Mid(lCodigoBarra2, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
    For i = 1 To 12
        If Asc(Mid(lCodigoBarra3, i, 1)) < 48 Or Asc(Mid(lCodigoBarra3, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
End Function
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data_emissao.Text < g_cfg_data_i Or msk_data_emissao.Text > g_cfg_data_f Then
        MsgBox "A data de emissão deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data_emissao.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cheque_devolvido.Show 1
    If Len(g_string) > 0 Then
        lDataEmissao = RetiraGString(1)
        lConta = RetiraGString(2)
        lCheque = RetiraGString(3)
        If MovChequeDevolvido.LocalizarRegistro(g_empresa, lDataEmissao, lConta, lCheque) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovChequeDevolvido.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovChequeDevolvido.LocalizarProximo Then
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
    If MovChequeDevolvido.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_conta.SetFocus
    End If
End Sub
Private Sub dtcboSituacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        lGravados = 0
        DesativaBotoes
        If MovChequeDevolvido.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
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
    CentraForm Me
    adodc_banco.ConnectionString = gConnectionString
    adodc_banco.RecordSource = "SELECT Codigo, Nome FROM Bancos ORDER BY Nome"
    adodc_banco.Refresh
    
    adodcSituacao.ConnectionString = gConnectionString
    adodcSituacao.RecordSource = "SELECT Codigo, Nome FROM Situacao_Cheque_Devolvido ORDER BY Nome"
    adodcSituacao.Refresh
    
    PreencheCboPeriodo
    PreencheCboTipoMovimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub msk_data_devolucao_GotFocus()
    msk_data_devolucao.SelStart = 0
    msk_data_devolucao.SelLength = 5
End Sub
Private Sub msk_data_devolucao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo_devolucao.SetFocus
    End If
End Sub
Private Sub msk_data_digitacao_GotFocus()
    If Not IsDate(msk_data_digitacao.Text) Then
        msk_data_digitacao.Text = Format(CDate(g_data_def), "dd/mm/yyyy")
    End If
    msk_data_digitacao.SelStart = 0
    msk_data_digitacao.SelLength = 5
End Sub
Private Sub msk_data_digitacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_emissao.SetFocus
    End If
End Sub
Private Sub msk_data_emissao_GotFocus()
    If Not IsDate(msk_data_emissao.Text) Then
        msk_data_emissao.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
    End If
    msk_data_emissao.SelStart = 0
    msk_data_emissao.SelLength = 5
End Sub
Private Sub msk_data_emissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub msk_data_emissao_LostFocus()
    If IsDate(msk_data_emissao.Text) And lOpcao = 1 Then
    End If
End Sub
Private Sub msk_data_vencimento_GotFocus()
    msk_data_vencimento.SelStart = 0
    msk_data_vencimento.SelLength = 5
End Sub
Private Sub msk_data_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_banco.SetFocus
    End If
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_vencimento.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    If Val(msk_valor.Text) > 0 Then
        msk_valor.Text = Format(msk_valor.Text, "###,##0.00")
    End If
End Sub
Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_emitente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cheque_LostFocus()
    If lOpcao = 1 Then
        If MovChequeDevolvido.ExisteRegistro(g_empresa, msk_data_emissao, txt_conta, txt_cheque) Then
            MsgBox "Cheque já cadastrado.", vbInformation, "Atenção!"
            txt_cheque.SetFocus
        End If
    End If
End Sub
Private Sub txt_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cheque.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_conta_LostFocus()
    Dim xString As String
    If lOpcao = 1 And txt_emitente.Text = "" Then
        xString = MovChequeDevolvido.LocalizarConta(txt_conta.Text)
        If xString <> "" Then
            txt_emitente.Text = xString
        End If
    End If
End Sub
Private Sub txt_emitente_GotFocus()
    txt_emitente.SelStart = 0
    txt_emitente.SelLength = Len(txt_emitente.Text)
End Sub
Private Sub txt_emitente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCnpjCpf.SetFocus
    End If
End Sub
Private Sub txt_motivo_devolucao_GotFocus()
    txt_motivo_devolucao.SelStart = 0
    txt_motivo_devolucao.SelLength = Len(txt_motivo_devolucao.Text)
End Sub
Private Sub txt_motivo_devolucao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboSituacao.SetFocus
    End If
End Sub
Private Sub txtCnpjCpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_devolucao.SetFocus
    End If
End Sub
