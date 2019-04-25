VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form movimento_historico 
   Caption         =   "Movimento do Histórico do Caixa"
   ClientHeight    =   5415
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   7875
   Icon            =   "movimento_historico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_historico.frx":030A
   ScaleHeight     =   5415
   ScaleWidth      =   7875
   Begin VB.CommandButton Command2 
      Caption         =   "Conversao"
      Height          =   735
      Left            =   4320
      TabIndex        =   52
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_historico.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Cria um novo registro."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_historico.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Altera o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_historico.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_historico.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_historico.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4440
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.TextBox txt_despesa 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   50
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_numero_ilha 
         Height          =   300
         Left            =   6420
         MaxLength       =   1
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txt_hipercheque 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_transferencia 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   28
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_afericao 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_assalto 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   30
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4500
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_amex 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txt_dinners 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_visa 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_nota 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_dinheiro 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_cheque_avista 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_cheque_predatado 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_observacao 
         Height          =   285
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   49
         Top             =   4980
         Width           =   4215
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   34
         Top             =   3840
         Width           =   795
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "movimento_historico.frx":7472
         Height          =   315
         Left            =   2880
         TabIndex        =   35
         Top             =   3840
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "nome"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin VB.Label Label15 
         Caption         =   "Despesas do Caixa"
         Height          =   315
         Left            =   4500
         TabIndex        =   51
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Número da &Ilha"
         Height          =   300
         Left            =   4500
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Outros Cartões"
         Height          =   315
         Left            =   4500
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Transferência"
         Height          =   315
         Left            =   4500
         TabIndex        =   27
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Aferição/Devolução"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Sinistro"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Sollo/Amex"
         Height          =   315
         Left            =   4500
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "CredCard/Dinners"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Visa"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Nota firma"
         Height          =   315
         Left            =   4500
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Dinheiro"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cheque a vista"
         Height          =   315
         Left            =   4500
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Cheque pré-datado"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Observação"
         Height          =   315
         Index           =   11
         Left            =   180
         TabIndex        =   48
         Top             =   4980
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   3840
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
      Left            =   5580
      TabIndex        =   43
      Top             =   4320
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_historico.frx":7490
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_historico.frx":898A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_historico.frx":9E84
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_historico.frx":B2F6
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_historico.frx":C878
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "movimento_historico.frx":DE82
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4440
      Width           =   795
   End
End
Attribute VB_Name = "movimento_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_historico As Integer
Dim lOpcao As String
Dim l_empresa As Integer
Dim l_data As Date
Dim l_periodo As String
Dim l_ilha As Integer
Dim l_qtd_periodo As Integer
Dim l_tipo_movimento As String
Dim tbl_configuracao As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_afericao As Table
Dim tbl_movimento_cartao_credito As Table
Dim tbl_movimento_cheque As Table
Dim tbl_movimento_cheque_avista As Table
Dim tbl_movimento_despesa_caixa As Table
Dim tbl_movimento_historico As Table
Dim tbl_movimento_nota As Table
Private MovimentoComposicaoCaixa As New cMovimentoComposicaoCaixa
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
Function BuscaDespesaCaixa(x_data As Date, x_periodo As String, x_tipo_movimento As String) As Currency
    BuscaDespesaCaixa = 0
    '+Empresa;+Data do Movimento;+Periodo;+Tipo do Movimento;+Registro
    With tbl_movimento_despesa_caixa
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, x_data, x_periodo, x_tipo_movimento, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Data do Movimento] <> x_data Or !Periodo <> x_periodo Or ![Tipo do Movimento] <> x_tipo_movimento Then
                        Exit Do
                    End If
                    BuscaDespesaCaixa = BuscaDespesaCaixa + !valor
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function BuscaNotasFirma(x_data As Date, x_periodo As String, x_tipo_movimento As String) As Currency
    BuscaNotasFirma = 0
    With tbl_movimento_nota
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, x_data, x_periodo, 0, 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Data do Abastecimento] <> x_data Or !Periodo <> x_periodo Then
                        Exit Do
                    End If
                    If ![Tipo do Movimento] = x_tipo_movimento Then
                        BuscaNotasFirma = BuscaNotasFirma + tbl_movimento_nota![Valor Total]
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function BuscaChequePreDatado(x_data As Date, x_periodo As String, x_tipo_movimento As String) As Currency
    BuscaChequePreDatado = 0
    If tbl_movimento_cheque.RecordCount > 0 Then
        tbl_movimento_cheque.Seek ">", g_empresa, x_data, "          ", "      "
        If Not tbl_movimento_cheque.NoMatch Then
            Do Until tbl_movimento_cheque.EOF
                If tbl_movimento_cheque!Empresa <> g_empresa Or tbl_movimento_cheque![Data de Emissao] <> x_data Then
                    Exit Do
                End If
                If tbl_movimento_cheque!Periodo = x_periodo And tbl_movimento_cheque![Tipo do Movimento] = x_tipo_movimento Then
                    BuscaChequePreDatado = BuscaChequePreDatado + tbl_movimento_cheque!valor
                End If
                tbl_movimento_cheque.MoveNext
            Loop
        End If
    End If
End Function
Function BuscaChequeAvista(x_data As Date, x_periodo As String, x_tipo_movimento As String) As Currency
    BuscaChequeAvista = 0
    If x_tipo_movimento = 3 Then
        Exit Function
    End If
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, x_data, " ", " ", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Data de Emissao] <> x_data Then
                        Exit Do
                    End If
                    If !Periodo = x_periodo And ![Tipo do Movimento] = x_tipo_movimento Then
                        BuscaChequeAvista = BuscaChequeAvista + !valor
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function BuscaCartaoCredito(x_data As Date, x_periodo As String, x_tipo_movimento As String, x_codigo As Integer) As Currency
    BuscaCartaoCredito = 0
    With tbl_movimento_cartao_credito
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, x_data, x_periodo, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Data de Emissao] <> x_data Then
                        Exit Do
                    End If
                    If !Periodo = cbo_periodo And ![Tipo do Movimento] = x_tipo_movimento And ![Codigo do Cartao] = x_codigo Then
                        BuscaCartaoCredito = BuscaCartaoCredito + !valor
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function BuscaAfericao(x_data As Date, x_periodo As String, x_tipo_movimento As String, x_transferencia As Boolean) As Currency
    If x_tipo_movimento = 3 Then
        Exit Function
    End If
    BuscaAfericao = 0
    If Mid(cbo_tipo_movimento, 1, 1) = 1 Then
        With tbl_movimento_afericao
            .Seek ">=", g_empresa, CDate(x_data), x_periodo, 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or !Data <> x_data Then
                        Exit Do
                    End If
                    If !Periodo = x_periodo And !Transferencia = x_transferencia Then
                        BuscaAfericao = BuscaAfericao + ![Valor Total]
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
End Function
Private Sub CalculaTotal()
    lbl_total = Format(fValidaValor2(txt_visa) + fValidaValor2(txt_dinners) + fValidaValor2(txt_amex) + fValidaValor2(txt_hipercheque) + fValidaValor2(txt_cheque_predatado) + fValidaValor2(txt_cheque_avista) + fValidaValor2(txt_dinheiro) + fValidaValor2(txt_nota) + fValidaValor2(txt_assalto) + fValidaValor2(txt_afericao) + fValidaValor2(txt_transferencia) + fValidaValor2(txt_despesa), "###,##0.00")
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
    If tbl_configuracao.RecordCount = 8 Then
        If l_qtd_periodo = 3 Then
            l_qtd_periodo = 4
        End If
    End If
End Sub
Private Sub AtualTabe()
    With tbl_movimento_historico
        l_data = msk_data
        l_periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
        l_ilha = Val(txt_numero_ilha)
        l_tipo_movimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
        !Empresa = g_empresa
        !Data = msk_data
        !Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
        ![Numero da Ilha] = Val(txt_numero_ilha)
        ![Tipo do Movimento] = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
        !Visa = fValidaValor2(txt_visa)
        !Dinners = fValidaValor2(txt_dinners)
        !Amex = fValidaValor2(txt_amex)
        ![Cheque Pre-Datado] = fValidaValor2(txt_cheque_predatado)
        ![Cheque A Vista] = fValidaValor2(txt_cheque_avista)
        !Hipercheque = fValidaValor2(txt_hipercheque)
        !Dinheiro = fValidaValor2(txt_dinheiro)
        !Nota = fValidaValor2(txt_nota)
        !Assalto = fValidaValor2(txt_assalto)
        !Afericao = fValidaValor2(txt_afericao)
        !Transferencia = fValidaValor2(txt_transferencia)
        ![Despesa do Caixa] = fValidaValor2(txt_despesa)
        !total = fValidaValor2(lbl_total)
        ![Codigo do Funcionario] = dbcbo_funcionario.BoundText
    End With
End Sub
Private Sub AtualTela()
    Dim i As Integer
    With tbl_movimento_historico
        l_data = !Data
        l_periodo = !Periodo
        l_ilha = ![Numero da Ilha]
        l_tipo_movimento = ![Tipo do Movimento]
        msk_data = Format(!Data, "dd/mm/yyyy")
        cbo_periodo.ListIndex = -1
        For i = 0 To cbo_periodo.ListCount - 1
            If cbo_periodo.ItemData(i) = !Periodo Then
                cbo_periodo.ListIndex = i
                Exit For
            End If
        Next
        txt_numero_ilha = ![Numero da Ilha]
        cbo_tipo_movimento.ListIndex = -1
        For i = 0 To cbo_tipo_movimento.ListCount - 1
            If cbo_tipo_movimento.ItemData(i) = ![Tipo do Movimento] Then
                cbo_tipo_movimento.ListIndex = i
                Exit For
            End If
        Next
        txt_visa = Format(!Visa, "###,##0.00")
        txt_dinners = Format(!Dinners, "###,##0.00")
        txt_amex = Format(!Amex, "###,##0.00")
        txt_hipercheque = Format(!Hipercheque, "###,##0.00")
        txt_cheque_predatado = Format(![Cheque Pre-Datado], "###,##0.00")
        txt_cheque_avista = Format(![Cheque A Vista], "###,##0.00")
        txt_dinheiro = Format(!Dinheiro, "###,##0.00")
        txt_nota = Format(!Nota, "###,##0.00")
        txt_assalto = Format(!Assalto, "###,##0.00")
        txt_despesa = Format(![Despesa do Caixa], "###,##0.00")
        txt_afericao = Format(!Afericao, "###,##0.00")
        txt_transferencia = Format(!Transferencia, "###,##0.00")
        CalculaTotal
        txt_funcionario = ![Codigo do Funcionario]
        dbcbo_funcionario.BoundText = ![Codigo do Funcionario]
    End With
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.Seek "<", g_empresa, CDate("31/12/2500"), 9, 9, 9
        If Not tbl_movimento_historico.NoMatch Then
            If tbl_movimento_historico!Empresa = g_empresa Then
                AtualTela
                BuscaDados = True
                Exit Function
            End If
        End If
    End If
    LimpaTela
End Function
Function BuscaRegistro(x_data As Date, x_periodo As String, x_ilha As Integer, x_tipo_movimento As String) As Boolean
    BuscaRegistro = False
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.Seek "=", g_empresa, x_data, x_periodo, x_ilha, x_tipo_movimento
        If Not tbl_movimento_historico.NoMatch Then
            AtualTela
            BuscaRegistro = True
        End If
    End If
End Function
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
Private Sub Finaliza()
    tbl_configuracao.Close
    tbl_funcionario.Close
    tbl_movimento_afericao.Close
    tbl_movimento_cartao_credito.Close
    tbl_movimento_cheque.Close
    tbl_movimento_cheque_avista.Close
    tbl_movimento_despesa_caixa.Close
    tbl_movimento_historico.Close
    tbl_movimento_nota.Close
End Sub
Private Sub GravaComposicao(xCodigo As Integer, xValor As Currency)
    If xValor = 0 Then
        Exit Sub
    End If
    With tbl_movimento_historico
        MovimentoComposicaoCaixa.Empresa = !Empresa
        MovimentoComposicaoCaixa.Data = !Data
        MovimentoComposicaoCaixa.Periodo = !Periodo
        MovimentoComposicaoCaixa.NumeroIlha = ![Numero da Ilha]
        MovimentoComposicaoCaixa.TipoMovimento = ![Tipo do Movimento]
        MovimentoComposicaoCaixa.CodigoFuncionario = ![Codigo do Funcionario]
        MovimentoComposicaoCaixa.CodigoComposicao = xCodigo
        MovimentoComposicaoCaixa.valor = xValor
        If Not MovimentoComposicaoCaixa.Incluir Then
            MsgBox "Registro não foi gravado!", vbInformation, "Erro Interno"
        End If
    End With
End Sub
Function PeriodoAnteriorAberto() As Boolean
    Dim x_data As Date
    Dim x_data_teste As Date
    Dim x_periodo As String
    Dim x_ilha As Integer
    Dim x_tipo_movimento As String
    Dim x_valor As Currency
    PeriodoAnteriorAberto = False
    x_data = msk_data
    x_data_teste = msk_data
    x_periodo = Val(cbo_periodo)
    x_ilha = Val(txt_numero_ilha)
    x_tipo_movimento = Val(cbo_tipo_movimento)
    Do Until (x_data - x_data_teste) = 2
        If x_tipo_movimento > 1 Then
            x_tipo_movimento = x_tipo_movimento - 1
        Else
            If x_periodo > 1 Then
                x_periodo = x_periodo - 1
                x_tipo_movimento = 3
            Else
                x_data_teste = x_data_teste - 1
                x_periodo = 4
                x_tipo_movimento = 3
            End If
        End If
        If tbl_movimento_historico.RecordCount > 0 Then
            tbl_movimento_historico.Seek "=", g_empresa, x_data_teste, x_periodo, x_ilha, x_tipo_movimento
            If tbl_movimento_historico.NoMatch Then
                'Cheque Pré-Datado
                x_valor = BuscaChequePreDatado(x_data_teste, x_periodo, x_tipo_movimento)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cheque pré-datado: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Cheque À Vista
                x_valor = BuscaChequeAvista(x_data_teste, x_periodo, x_tipo_movimento)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cheque a vista: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Nota Firma
                x_valor = BuscaNotasFirma(x_data_teste, x_periodo, x_tipo_movimento)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em nota abastecimento: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Amex / Sollo
                x_valor = BuscaCartaoCredito(x_data_teste, x_periodo, x_tipo_movimento, 3)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cartão Amex/Sollo: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Visa
                x_valor = BuscaCartaoCredito(x_data_teste, x_periodo, x_tipo_movimento, 1)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cartão Visa: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'CrediCard / Dinners
                x_valor = BuscaCartaoCredito(x_data_teste, x_periodo, x_tipo_movimento, 2)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cartão Credicard/Dinners: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'HiperCheque
                x_valor = BuscaCartaoCredito(x_data_teste, x_periodo, x_tipo_movimento, 4)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Tipo de Movimento: " & x_tipo_movimento & Chr(10) & "Total em cartão Hipercheque: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Aferição
                x_valor = BuscaAfericao(x_data_teste, x_periodo, x_tipo_movimento, False)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Total em Aferição: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
                'Transferência
                x_valor = BuscaAfericao(x_data_teste, x_tipo_movimento, x_periodo, True)
                If x_valor > 0 Then
                    PeriodoAnteriorAberto = True
                    MsgBox "Não existe histórico!" & Chr(10) & "Data: " & x_data_teste & Chr(10) & "Período: " & x_periodo & Chr(10) & "Total em Transferência: " & Format(x_valor, "###,###,##0.00"), vbInformation, "Erro de Fechamento!"
                End If
            End If
        End If
    Loop
End Function
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
    cbo_tipo_movimento.AddItem "3 Caixa da Borracharia/Lavagem"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
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
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cheque_predatado.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_LostFocus()
    If PeriodoAnteriorAberto Then
        msk_data.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_cheque_predatado.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.MovePrevious
        If Not tbl_movimento_historico.BOF Then
            If tbl_movimento_historico!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        tbl_movimento_historico.MoveNext
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento) Then
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
Private Sub LimpaTela()
    msk_data = "__/__/____"
    cbo_periodo.ListIndex = -1
    cbo_tipo_movimento.ListIndex = -1
    txt_visa = ""
    txt_dinners = ""
    txt_amex = ""
    txt_hipercheque = ""
    txt_cheque_predatado = ""
    txt_cheque_avista = ""
    txt_dinheiro = ""
    txt_nota = ""
    txt_assalto = ""
    txt_despesa = ""
    txt_afericao = ""
    txt_transferencia = ""
    lbl_total = ""
    txt_funcionario = ""
    dbcbo_funcionario.BoundText = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_funcionario) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            tbl_movimento_historico.Edit
            tbl_movimento_historico.Delete
            LimpaTela
            If Not BuscaDados Then
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If BuscaProximoCaixa Then
        cbo_periodo.SetFocus
    Else
        msk_data.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                tbl_movimento_historico.AddNew
                AtualTabe
                tbl_movimento_historico.Update
            ElseIf lOpcao = 2 Then
                tbl_movimento_historico.Edit
                AtualTabe
                tbl_movimento_historico.Update
            End If
            lOpcao = 0
            Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento)
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_movimento_historico.name, "Historicoo"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data do movimento.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not ValidaPeriodo Then
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf fValidaValor2(lbl_total) = 0 Then
        MsgBox "Informe o valor de algum campo.", 64, "Atenção!"
        txt_dinheiro.SetFocus
    ElseIf Val(dbcbo_funcionario.BoundText) = 0 Then
        MsgBox "Escolha o funcionario.", 64, "Atenção!"
        dbcbo_funcionario.SetFocus
    ElseIf Not Val(txt_numero_ilha) > 0 Then
        MsgBox "O número da ilha deve ser maior que 0.", 64, "Atenção!"
        txt_numero_ilha.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaPeriodo() As Boolean
    ValidaPeriodo = True
    If cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", 64, "Atenção!"
        ValidaPeriodo = False
    ElseIf cbo_periodo.ListIndex = 3 And cbo_tipo_movimento.ListIndex <> 1 Then
        If (MsgBox("Tipo de movimento não aceito para este período." & Chr(10) & Chr(10) & "Deseja cadastrar mesmo assim?", vbYesNo + vbDefaultButton2, "Atenção!")) = 7 Then
            ValidaPeriodo = False
        End If
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    With tbl_movimento_historico
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
    consulta_historico.Show 1
    If Len(g_string) > 0 Then
        l_data = RetiraGString(1)
        l_periodo = RetiraGString(2)
        l_ilha = RetiraGString(3)
        l_tipo_movimento = RetiraGString(4)
        Call BuscaRegistro(l_data, l_periodo, l_ilha, l_tipo_movimento)
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.Seek ">", g_empresa, CDate("01/01/1900"), 0, 0, 0
        If Not tbl_movimento_historico.NoMatch Then
            If tbl_movimento_historico!Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.MoveNext
        If Not tbl_movimento_historico.EOF Then
            If tbl_movimento_historico!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        tbl_movimento_historico.MovePrevious
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If tbl_movimento_historico.RecordCount > 0 Then
        tbl_movimento_historico.Seek "<", g_empresa, CDate("31/12/2500"), 9, 9, 9
        If Not tbl_movimento_historico.NoMatch Then
            If tbl_movimento_historico!Empresa = g_empresa Then
                AtualTela
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub Command1_Click()
    Dim sql As String
    Exit Sub
    sql = "Update Movimento_Historico "
    sql = sql & "Set Afericao = 0, Transferencia = 0"
    bd_sgp.Execute sql
End Sub
Private Sub Command2_Click()
    With tbl_movimento_historico
        .MoveFirst
        Do Until .EOF
            Call GravaComposicao(1, !Dinheiro)
            Call GravaComposicao(2, ![Cheque A Vista])
            Call GravaComposicao(3, ![Cheque Pre-Datado])
            Call GravaComposicao(4, !Nota)
            Call GravaComposicao(5, ![Despesa do Caixa])
            Call GravaComposicao(6, !Assalto)
            Call GravaComposicao(7, !Afericao)
            Call GravaComposicao(8, !Transferencia)
            Call GravaComposicao(9, !Visa)
            Call GravaComposicao(10, !Dinners)
            Call GravaComposicao(11, !Amex)
            .MoveNext
        Loop
    End With
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub TabelaFuncionarioRefresh()
    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = " & Chr(34) & "A" & Chr(34) & " And [Periodo] < 5 Order By [Nome]"
    dta_funcionario.Refresh
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If dbcbo_funcionario.BoundText <> "" Then
        txt_funcionario = dbcbo_funcionario.BoundText
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_movimento_historico = 0
    End If
    If flag_movimento_historico = 0 Then
        AtualizaConstantes
        TabelaFuncionarioRefresh
        lOpcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        flag_movimento_historico = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_historico = 1
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
    Set tbl_movimento_afericao = bd_sgp.OpenTable("Movimento_Afericao")
    Set tbl_movimento_cartao_credito = bd_sgp.OpenTable("Movimento_Cartao_Credito")
    Set tbl_movimento_cheque = bd_sgp.OpenTable("Movimento_Cheque")
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    Set tbl_movimento_despesa_caixa = bd_sgp.OpenTable("Movimento_Despesa_Caixa")
    Set tbl_movimento_historico = bd_sgp.OpenTable("Movimento_Historico")
    Set tbl_movimento_nota = bd_sgp.OpenTable("Movimento_Nota_Abastecimento")
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_afericao.Index = "id_data"
    tbl_movimento_cartao_credito.Index = "id_data_emissao"
    tbl_movimento_cheque.Index = "id_data_emissao"
    tbl_movimento_cheque_avista.Index = "id_digitacao"
    tbl_movimento_despesa_caixa.Index = "id_data"
    tbl_movimento_historico.Index = "id_data"
    tbl_movimento_nota.Index = "id_data_abastecimento"
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    l_data = "01/01/1900"
    l_periodo = "0"
    l_tipo_movimento = "0"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_afericao_GotFocus()
    txt_afericao = BuscaAfericao(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), False)
End Sub
Private Sub txt_afericao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_transferencia.SetFocus
    End If
End Sub
Private Sub txt_afericao_LostFocus()
    txt_afericao = Format(txt_afericao, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_amex_GotFocus()
    txt_amex = BuscaCartaoCredito(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), 3)
End Sub
Private Sub txt_amex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_visa.SetFocus
    End If
End Sub
Private Sub txt_amex_LostFocus()
    txt_amex = Format(txt_amex, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_assalto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_despesa.SetFocus
    End If
End Sub
Private Sub txt_assalto_LostFocus()
    txt_assalto = Format(txt_assalto, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_cheque_avista_GotFocus()
    txt_cheque_avista = BuscaChequeAvista(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    txt_cheque_avista.SelStart = 0
    txt_cheque_avista.SelLength = Len(txt_cheque_avista)
End Sub
Private Sub txt_cheque_avista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_dinheiro.SetFocus
    End If
End Sub
Private Sub txt_cheque_avista_LostFocus()
    txt_cheque_avista = Format(txt_cheque_avista, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_cheque_predatado_GotFocus()
    txt_cheque_predatado = BuscaChequePreDatado(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    txt_cheque_predatado.SelStart = 0
    txt_cheque_predatado.SelLength = Len(txt_cheque_predatado)
End Sub
Private Sub txt_cheque_predatado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_cheque_avista.SetFocus
    End If
End Sub
Private Sub txt_cheque_predatado_LostFocus()
    txt_cheque_predatado = Format(txt_cheque_predatado, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_despesa_GotFocus()
    txt_despesa = BuscaDespesaCaixa(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
End Sub
Private Sub txt_despesa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub txt_despesa_LostFocus()
    txt_despesa = Format(txt_despesa, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_dinheiro_GotFocus()
    txt_dinheiro.SelStart = 0
    txt_dinheiro.SelLength = Len(txt_dinheiro)
End Sub
Private Sub txt_dinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_nota.SetFocus
    End If
End Sub
Private Sub txt_dinheiro_LostFocus()
    txt_dinheiro = Format(txt_dinheiro, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_dinners_GotFocus()
    txt_dinners = BuscaCartaoCredito(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), 2)
End Sub
Private Sub txt_dinners_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_amex.SetFocus
    End If
End Sub
Private Sub txt_dinners_LostFocus()
    txt_dinners = Format(txt_dinners, "###,##0.00")
    CalculaTotal
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
            If tbl_funcionario!Situacao = "A" Then
                dbcbo_funcionario.BoundText = Val(txt_funcionario)
                cmd_ok.SetFocus
                Exit Sub
            Else
                MsgBox "O funcionário " & Trim(tbl_funcionario!Nome) & " está inativo.", 64, "Aviso de Verificação!"
                txt_funcionario.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastro.", 64, "Aviso de Verificação!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_hipercheque_GotFocus()
    txt_hipercheque = BuscaCartaoCredito(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), 4)
End Sub
Private Sub txt_hipercheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_afericao.SetFocus
    End If
End Sub
Private Sub txt_hipercheque_LostFocus()
    txt_hipercheque = Format(txt_hipercheque, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_nota_GotFocus()
    If cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) > -1 Then
        txt_nota = BuscaNotasFirma(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    End If
    txt_nota.SelStart = 0
    txt_nota.SelLength = Len(txt_nota)
End Sub
Private Sub txt_nota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If g_nome_empresa = "AUTO POSTO BATISTA MENDES LTDA" Then
            txt_funcionario.SetFocus
        Else
            txt_dinners.SetFocus
        End If
    End If
End Sub
Private Sub txt_nota_LostFocus()
    txt_nota = Format(txt_nota, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_numero_ilha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_transferencia_GotFocus()
    txt_transferencia = BuscaAfericao(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), True)
End Sub
Private Sub txt_transferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_assalto.SetFocus
    End If
End Sub
Private Sub txt_transferencia_LostFocus()
    txt_transferencia = Format(txt_transferencia, "###,##0.00")
    CalculaTotal
End Sub
Private Sub txt_visa_GotFocus()
    txt_visa = BuscaCartaoCredito(CDate(msk_data), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), 1)
End Sub
Private Sub txt_visa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_hipercheque.SetFocus
    End If
End Sub
Private Sub txt_visa_LostFocus()
    txt_visa = Format(txt_visa, "###,##0.00")
    CalculaTotal
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    Dim x_tipo_movimento As String
    BuscaProximoCaixa = False
    With tbl_movimento_historico
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate("31/12/2500"), 9, 9, 9
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    msk_data = !Data
                    x_periodo = !Periodo - 1
                    x_tipo_movimento = ![Tipo do Movimento] - 1
                    x_tipo_movimento = x_tipo_movimento + 1
                    
                    If x_tipo_movimento = 2 Then
                        x_tipo_movimento = 0
                        x_periodo = x_periodo + 1
                    End If
                    
                    
                    If x_periodo > (l_qtd_periodo - 1) Then
                        msk_data = !Data + 1
                        x_periodo = 0
                        x_tipo_movimento = 0
                    End If
                    If x_periodo = 3 Then
                        x_tipo_movimento = 1
                    End If
                    cbo_periodo.ListIndex = x_periodo
                    cbo_tipo_movimento.ListIndex = x_tipo_movimento
                    BuscaProximoCaixa = True
                    Exit Function
                End If
            End If
        End If
        msk_data = g_data_def - 1
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
    End With
End Function

