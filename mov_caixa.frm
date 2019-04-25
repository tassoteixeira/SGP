VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_caixa 
   Caption         =   "Movimento do Caixa"
   ClientHeight    =   5205
   ClientLeft      =   1170
   ClientTop       =   1065
   ClientWidth     =   9885
   Icon            =   "mov_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   37
      Top             =   3480
      Width           =   9735
      Begin VB.Label lblTotalCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   7800
         TabIndex        =   41
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lblTotalDebito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2580
         TabIndex        =   40
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total de crédito do dia"
         Height          =   315
         Left            =   5340
         TabIndex        =   39
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Total de débito do dia"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "mov_caixa.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4260
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "mov_caixa.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4260
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "mov_caixa.frx":2E0E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4260
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "mov_caixa.frx":44A0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Altera o registro atual."
      Top             =   4260
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "mov_caixa.frx":599A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cria um novo registro."
      Top             =   4260
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin MSAdodcLib.Adodc adodcContaCredito 
         Height          =   330
         Left            =   4560
         Top             =   1440
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
         Caption         =   "adodcContaCredito"
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
      Begin VB.TextBox txtContaReduzidaCredito 
         Height          =   285
         Left            =   2580
         MaxLength       =   5
         TabIndex        =   34
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtHistoricoPadrao 
         Height          =   285
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   16
         Top             =   2220
         Width           =   555
      End
      Begin VB.CheckBox chkFluxoCaixa 
         Height          =   315
         Left            =   8100
         TabIndex        =   9
         Top             =   660
         Width           =   795
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8100
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtContaReduzidaDebito 
         Height          =   285
         Left            =   2580
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3660
         Picture         =   "mov_caixa.frx":702C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txt_documento 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2580
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txt_complemento 
         Height          =   315
         Left            =   2580
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2640
         Width           =   5535
      End
      Begin VB.TextBox txt_valor 
         Height          =   315
         Left            =   2580
         MaxLength       =   14
         TabIndex        =   14
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox txt_numero_movimento 
         Height          =   315
         Left            =   2580
         MaxLength       =   9
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   2580
         TabIndex        =   6
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
      Begin MSAdodcLib.Adodc adodc_historico 
         Height          =   330
         Left            =   4560
         Top             =   2280
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
         Caption         =   "adodc_historico"
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
      Begin MSDataListLib.DataCombo dtcbo_historico 
         Bindings        =   "mov_caixa.frx":8306
         Height          =   315
         Left            =   3240
         TabIndex        =   17
         Top             =   2220
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_historico"
      End
      Begin MSAdodcLib.Adodc adodcContaDebito 
         Height          =   330
         Left            =   4530
         Top             =   1140
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
         Caption         =   "adodcContaDebito"
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
      Begin MSDataListLib.DataCombo dtcboContaDebito 
         Bindings        =   "mov_caixa.frx":8324
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Conta Reduzida"
         Text            =   "dtcboContaDebito"
      End
      Begin MSDataListLib.DataCombo dtcboContaCredito 
         Bindings        =   "mov_caixa.frx":8343
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Top             =   1440
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Conta Reduzida"
         Text            =   "dtcboContaCredito"
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "&Conta Reduzida (Crédito)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   2400
      End
      Begin VB.Label Label10 
         Caption         =   "Fluxo de caixa"
         Height          =   255
         Left            =   6600
         TabIndex        =   8
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de movimento"
         Height          =   255
         Left            =   6600
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "&Conta Reduzida (Débito)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "&Histórico"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2220
         Width           =   2400
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Número do doc&umento"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   3060
         Width           =   2400
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   2400
      End
      Begin VB.Label Label1 
         Caption         =   "&Número do movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label5 
         Caption         =   "&Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2400
      End
      Begin VB.Label Label6 
         Caption         =   "C&omplemento"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   2400
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   7620
      TabIndex        =   29
      Top             =   4140
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_caixa.frx":8363
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_caixa.frx":98E5
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_caixa.frx":AD57
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_caixa.frx":C251
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   8100
      Picture         =   "mov_caixa.frx":D74B
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4260
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   9000
      Picture         =   "mov_caixa.frx":ED55
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4260
      Width           =   795
   End
End
Attribute VB_Name = "movimento_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lData As Date
Dim lNumeroMovimento As Long
Dim l_valor As Currency
Private HistoricoPadrao As New cHistoricoPadrao
Private PlanoConta As New cPlanoConta
Private MovCaixa As New cMovimentoCaixa
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Function AtualTabe() As Boolean
    AtualTabe = False
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(msk_data.Text)
    If lOpcao = 1 Then
        MovCaixa.NumeroMovimento = 1
    End If
    MovCaixa.valor = fValidaValor2(txt_valor.Text)
    'MovCaixa.DebitoouCredito = cbo_credito_debito.Text
    MovCaixa.NumeroDocumento = txt_documento.Text
    MovCaixa.CodigoHistorico = Val(dtcbo_historico.BoundText)
    MovCaixa.Complemento = txt_complemento.Text
    MovCaixa.NumeroContaDebito = ""
    If dtcboContaDebito.BoundText <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, dtcboContaDebito.BoundText) Then
            MovCaixa.NumeroContaDebito = PlanoConta.Codigo
        Else
            MsgBox "Conta não cadastrada=" & dtcboContaDebito.BoundText, vbCritical, "Erro de Integridade"
            Exit Function
        End If
    End If
    MovCaixa.NumeroContaCredito = ""
    If dtcboContaCredito.BoundText <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, dtcboContaCredito.BoundText) Then
            MovCaixa.NumeroContaCredito = PlanoConta.Codigo
        Else
            MsgBox "Conta não cadastrada=" & dtcboContaCredito.BoundText, vbCritical, "Erro de Integridade"
            Exit Function
        End If
    End If
    MovCaixa.TipoMovimento = cbo_tipo_movimento.ListIndex + 1
    If chkFluxoCaixa.Value = 0 Then
        MovCaixa.FluxoCaixa = False
    Else
        MovCaixa.FluxoCaixa = True
    End If
    AtualTabe = True
End Function
Private Sub AtualTela()
    lData = MovCaixa.Data
    lNumeroMovimento = MovCaixa.NumeroMovimento
    msk_data.Text = MovCaixa.Data
    txt_numero_movimento.Text = Format(MovCaixa.NumeroMovimento, "#,###,##0")
    cbo_tipo_movimento.ListIndex = MovCaixa.TipoMovimento - 1
    'If MovCaixa.DebitoouCredito = "C" Then
    '    cbo_credito_debito.ListIndex = 0
    'Else
    '    cbo_credito_debito.ListIndex = 1
    'End If
    'cbo_credito_debito.Text = MovCaixa.DebitoouCredito
    txt_valor.Text = Format(MovCaixa.valor, "###,###,##0.00")
    txt_documento.Text = MovCaixa.NumeroDocumento
    txtHistoricoPadrao.Text = MovCaixa.CodigoHistorico
    dtcbo_historico.BoundText = MovCaixa.CodigoHistorico
    txt_complemento.Text = MovCaixa.Complemento
    txtContaReduzidaDebito.Text = ""
    dtcboContaDebito.BoundText = ""
    If MovCaixa.NumeroContaDebito <> "" Then
        If PlanoConta.LocalizarCodigo(g_empresa, MovCaixa.NumeroContaDebito) Then
            txtContaReduzidaDebito.Text = PlanoConta.ContaReduzida
            dtcboContaDebito.BoundText = PlanoConta.ContaReduzida
        Else
            MsgBox "Conta não cadastrada=" & MovCaixa.NumeroContaDebito, vbCritical, "Erro de Integridade"
        End If
    End If
    txtContaReduzidaCredito.Text = ""
    dtcboContaCredito.BoundText = ""
    If MovCaixa.NumeroContaCredito <> "" Then
        If PlanoConta.LocalizarCodigo(g_empresa, MovCaixa.NumeroContaCredito) Then
            txtContaReduzidaCredito.Text = PlanoConta.ContaReduzida
            dtcboContaCredito.BoundText = PlanoConta.ContaReduzida
        Else
            MsgBox "Conta não cadastrada=" & MovCaixa.NumeroContaCredito, vbCritical, "Erro de Integridade"
        End If
    End If
    If MovCaixa.TipoMovimento = 1 Then
        cmd_alterar.Enabled = True
        cmd_excluir.Enabled = True
    Else
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
    If MovCaixa.FluxoCaixa = False Then
        chkFluxoCaixa.Value = 0
    Else
        chkFluxoCaixa.Value = 1
    End If
    lblTotalDebito.Caption = Format(MovCaixa.TotalData(g_empresa, lData, lData, "D"), "###,###,##0.00")
    lblTotalCredito.Caption = Format(MovCaixa.TotalData(g_empresa, lData, lData, "C"), "###,###,##0.00")
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_alterar.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set HistoricoPadrao = Nothing
    Set PlanoConta = Nothing
    Set MovCaixa = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_numero_movimento.Text = Format(MovCaixa.ProximoRegistro(g_empresa), "###,##0")
End Sub
Private Sub IncluiLancamentosAutomatico()
    Dim x_dinheiro As Currency
    Dim x_cheque_a_vista As Currency
    Dim x_numero_movimento As Integer
    'x_dinheiro = 0
    'x_cheque_a_vista = 0
    'x_numero_movimento = 0
    'With tbl_movimento_Complemento
    '    .Seek ">=", g_empresa, CDate(msk_data), 0, 0, 0
    '    If Not .NoMatch Then
    '        Do Until .EOF
    '            If !Empresa <> g_empresa Or !Data <> CDate(msk_data) Then
    '                Exit Do
    '            End If
    '            x_dinheiro = x_dinheiro + !Dinheiro
    '            x_cheque_a_vista = x_cheque_a_vista + ![Cheque A Vista]
    '            .MoveNext
    '        Loop
    '    End If
    'End With
    'With tbl_movimento_caixa
    '    If x_dinheiro > 0 Then
    '        x_numero_movimento = x_numero_movimento + 1
    '        .AddNew
    '        !Empresa = g_empresa
    '        ![Numero do Movimento] = x_numero_movimento
    '        !Data = msk_data
    '        !Complemento = "Vendas em Dinheiro"
    '        ![Debito ou Credito] = "C"
    '        !valor = x_dinheiro
    '        Call LoopAdcionaSaldo(!Data, !valor, ![Debito ou Credito])
    '        .Update
    '    End If
    '    If x_cheque_a_vista > 0 Then
    '        x_numero_movimento = x_numero_movimento + 1
    '        .AddNew
    '        !Empresa = g_empresa
    '        ![Numero do Movimento] = x_numero_movimento
    '        !Data = msk_data
    '        !Complemento = "Vendas em Cheque a Vista"
    '        ![Debito ou Credito] = "C"
    '        !valor = x_cheque_a_vista
    '        Call LoopAdcionaSaldo(!Data, !valor, ![Debito ou Credito])
    '        .Update
    '    End If
    'End With
End Sub
Private Sub cbo_credito_debito_GotFocus()
    'SendMessageLong cbo_credito_debito.hwnd, CB_SHOWDROPDOWN, True, 0
    'If lOpcao = 1 And txt_valor = "" Then
    '    cbo_credito_debito.ListIndex = 0
    'End If
End Sub
Private Sub cbo_credito_debito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_complemento.SetFocus
    ElseIf KeyAscii = 49 Then
        KeyAscii = 68
    ElseIf KeyAscii = 50 Then
        KeyAscii = 67
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
    txt_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovCaixa.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    lOpcao = 0
    LimpaTela
    If MovCaixa.LocalizarRegistro(g_empresa, lData, lNumeroMovimento) Then
        AtivaBotoes
        AtualTela
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
End Sub
Private Sub LimpaTela()
    txt_numero_movimento.Text = ""
    msk_data.Text = "__/__/____"
    chkFluxoCaixa.Value = 1
    txt_valor.Text = ""
    txtContaReduzidaDebito.Text = ""
    dtcboContaDebito.BoundText = ""
    txtContaReduzidaCredito.Text = ""
    dtcboContaCredito.BoundText = ""
    dtcbo_historico.BoundText = ""
    cbo_tipo_movimento.ListIndex = 0
    txt_complemento.Text = ""
    txt_documento.Text = ""
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    txt_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_excluir_Click()
    If msk_data.Text <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If MovCaixa.Excluir(g_empresa, CDate(msk_data.Text), lNumeroMovimento) Then
                LimpaTela
                If MovCaixa.LocalizarUltimo(g_empresa) Then
                    AtivaBotoes
                    AtualTela
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
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    msk_data.SetFocus
End Sub

Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    'Crtl + R
    If KeyAscii = 18 Then
        If (MsgBox("Deseja realmente recalcular o saldo das contas?", vbYesNo + vbQuestion + vbDefaultButton2, "Recalcula Saldo")) = vbYes Then
            MovCaixa.RecalculaSaldo (g_empresa)
            MsgBox "Saldo das contas recalculado com sucesso!", vbInformation, "Operação Concluída"
        End If
    End If
End Sub

Private Sub cmd_ok_Click()

On Error GoTo FileError
    
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            If AtualTabe Then
                If MovCaixa.Incluir > 0 Then
                    lData = MovCaixa.Data
                    lNumeroMovimento = MovCaixa.NumeroMovimento
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
            End If
        ElseIf lOpcao = 2 Then
            If AtualTabe Then
                If MovCaixa.Alterar(g_empresa, lData, lNumeroMovimento) Then
                    lData = MovCaixa.Data
                    lNumeroMovimento = MovCaixa.NumeroMovimento
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
            End If
        End If
        lOpcao = 0
        If MovCaixa.LocalizarRegistro(g_empresa, lData, lNumeroMovimento) Then
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
    'ErroArquivo tbl_movimento_caixa.Name, "Movimento de Caixao"
    Exit Sub
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 - Manual"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Automático"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Fechamento"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_numero_movimento.Text = "" Then
        MsgBox "Informe o número do movimento de caixa.", vbInformation, "Atenção!"
        txt_numero_movimento.SetFocus
    ElseIf Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data do movimento de caixa.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not fValidaValor(txt_valor.Text) > 0 Then
        MsgBox "Informe o valor.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    ElseIf dtcbo_historico.BoundText = "" Then
        MsgBox "Selecione um histórico.", vbInformation, "Atenção!"
        dtcbo_historico.SetFocus
    ElseIf dtcboContaDebito.BoundText = "" And dtcboContaCredito.BoundText = "" Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        dtcboContaDebito.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_movimento_caixa.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lNumeroMovimento = RetiraGString(2)
        If MovCaixa.LocalizarRegistro(g_empresa, lData, lNumeroMovimento) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovCaixa.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovCaixa.LocalizarProximo Then
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
    If MovCaixa.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_historico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_complemento.SetFocus
    End If
End Sub
Private Sub dtcbo_historico_LostFocus()
    If dtcbo_historico.BoundText <> "" Then
        If HistoricoPadrao.LocalizarCodigo(CInt(dtcbo_historico.BoundText)) Then
            txtHistoricoPadrao.Text = dtcbo_historico.BoundText
        Else
            MsgBox "Historico não existe!", vbInformation, "Erro de Integridade."
            dtcbo_historico.SetFocus
        End If
    End If
End Sub
Private Sub dtcboContaCredito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub dtcboContaCredito_LostFocus()
    If dtcboContaCredito.BoundText <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, CInt(dtcboContaCredito.BoundText)) Then
            txtContaReduzidaCredito.Text = dtcboContaCredito.BoundText
        Else
            MsgBox "Conta selecionada não existe!", vbInformation, "Erro de Integridade."
            dtcboContaCredito.SetFocus
        End If
    End If
End Sub
Private Sub dtcboContaDebito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub dtcboContaDebito_LostFocus()
    If dtcboContaDebito.BoundText <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, CInt(dtcboContaDebito.BoundText)) Then
            txtContaReduzidaDebito.Text = dtcboContaDebito.BoundText
        Else
            MsgBox "Conta selecionada não existe!", vbInformation, "Erro de Integridade."
            dtcboContaDebito.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    If lFlagMovimento = 0 Then
        DesativaBotoes
        If MovCaixa.LocalizarUltimo(g_empresa) Then
            AtivaBotoes
            AtualTela
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagMovimento = 0
    End If
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
'    adodc_historico.ConnectionString = gConnectionString
'    adodc_historico.RecordSource = "SELECT Codigo, Nome FROM HistoricoPadrao ORDER BY Nome"
'    adodc_historico.Refresh
    Set adodc_historico.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM HistoricoPadrao ORDER BY Nome")
'    adodcContaDebito.ConnectionString = gConnectionString
'    adodcContaDebito.RecordSource = "SELECT [Conta Reduzida], Nome FROM Plano_Conta WHERE Empresa = " & g_empresa & " AND [Tipo da Conta] = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome"
'    adodcContaDebito.Refresh
    Set adodcContaDebito.Recordset = Conectar.RsConexao("SELECT [Conta Reduzida], Nome FROM Plano_Conta WHERE Empresa = " & g_empresa & " AND [Tipo da Conta] = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome")
'    adodcContaCredito.ConnectionString = gConnectionString
'    adodcContaCredito.RecordSource = "SELECT [Conta Reduzida], Nome FROM Plano_Conta WHERE Empresa = " & g_empresa & " AND [Tipo da Conta] = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome"
'    adodcContaCredito.Refresh
    Set adodcContaCredito.Recordset = Conectar.RsConexao("SELECT [Conta Reduzida], Nome FROM Plano_Conta WHERE Empresa = " & g_empresa & " AND [Tipo da Conta] = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome")
    PreencheCboTipoMovimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data) Then
        If lData <> "00:00:00" Then
            msk_data = lData
        Else
            msk_data = g_data_def
        End If
    End If
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboContaDebito.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If lOpcao = 1 And IsDate(msk_data) Then
        'Call CriaNovoMovimento(msk_data)
        'If (MsgBox("Nesta data não existe lancamentos." & Chr(10) & Chr(13) & "Deseja realmente fazer os lançamentos de forma automática?", vbYesNo + vbDefaultButton1, "Inclusão Automática de Registro!")) = 6 Then
        '    If txt_numero_movimento = 1 Then
        '        IncluiLancamentosAutomatico
        '        Call CriaNovoMovimento(msk_data)
        '    End If
        'End If
    End If
End Sub
Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_documento.SetFocus
    End If
End Sub
Private Sub txt_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_numero_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_movimento_LostFocus()
    txt_numero_movimento = Format(txt_numero_movimento, "#,###,##0")
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_historico.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_LostFocus()
    If txt_valor.Text = "" Then
        txt_valor.Text = l_valor
    End If
    txt_valor.Text = Format(txt_valor.Text, "###,##0.00")
End Sub
Private Sub txtContaReduzidaCredito_GotFocus()
    txtContaReduzidaCredito.SelStart = 0
    txtContaReduzidaCredito.SelLength = Len(txtContaReduzidaCredito.Text)
End Sub
Private Sub txtContaReduzidaCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboContaCredito.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtContaReduzidaCredito_LostFocus()
    If lOpcao <> 0 And txtContaReduzidaCredito.Text <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, CInt(txtContaReduzidaCredito.Text)) Then
            dtcboContaCredito.BoundText = CInt(txtContaReduzidaCredito.Text)
            dtcboContaCredito_LostFocus
        Else
            MsgBox "Conta não cadastrada!", vbInformation, "Erro de Verificação!"
            dtcboContaCredito.BoundText = ""
            txtContaReduzidaCredito.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtContaReduzidaDebito_GotFocus()
    txtContaReduzidaDebito.SelStart = 0
    txtContaReduzidaDebito.SelLength = Len(txtContaReduzidaDebito.Text)
End Sub
Private Sub txtContaReduzidaDebito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboContaDebito.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtContaReduzidaDebito_LostFocus()
    If lOpcao <> 0 And txtContaReduzidaDebito.Text <> "" Then
        If PlanoConta.LocalizarContaReduzida(g_empresa, CInt(txtContaReduzidaDebito.Text)) Then
            dtcboContaDebito.BoundText = CInt(txtContaReduzidaDebito.Text)
            dtcboContaDebito_LostFocus
        Else
            MsgBox "Conta não cadastrada!", vbInformation, "Erro de Verificação!"
            dtcboContaDebito.BoundText = ""
            txtContaReduzidaDebito.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtHistoricoPadrao_GotFocus()
    txtHistoricoPadrao.SelStart = 0
    txtHistoricoPadrao.SelLength = Len(txtHistoricoPadrao.Text)
End Sub
Private Sub txtHistoricoPadrao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_historico.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtHistoricoPadrao_LostFocus()
    If lOpcao <> 0 And txtHistoricoPadrao.Text <> "" Then
        If HistoricoPadrao.LocalizarCodigo(CInt(txtHistoricoPadrao.Text)) Then
            dtcbo_historico.BoundText = CInt(txtHistoricoPadrao.Text)
            dtcbo_historico_LostFocus
        Else
            MsgBox "Histórico não cadastrado!", vbInformation, "Erro de Verificação!"
            dtcbo_historico.BoundText = ""
            txtHistoricoPadrao.SetFocus
            Exit Sub
        End If
    End If
End Sub
