VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form movimento_cheque_avista 
   Caption         =   "Movimentação de Cheques à Vista"
   ClientHeight    =   3135
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "movimento_cheque_avista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cheque_avista.frx":030A
   ScaleHeight     =   3135
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cheque_avista.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_cheque_avista.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_cheque_avista.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_cheque_avista.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cheque_avista.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox msk_valor 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1500
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_emissao 
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
      Begin VB.Frame frmCodigoBarra 
         Caption         =   "Código de Barra"
         ForeColor       =   &H80000002&
         Height          =   1515
         Left            =   3120
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txt_codigo_barra_1 
            Height          =   285
            Left            =   180
            MaxLength       =   8
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txt_codigo_barra_2 
            Height          =   285
            Left            =   2100
            MaxLength       =   10
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txt_codigo_barra_3 
            Height          =   285
            Left            =   180
            MaxLength       =   12
            TabIndex        =   17
            Top             =   1050
            Width           =   1455
         End
         Begin VB.CommandButton cmd_ok2 
            Caption         =   "O&K"
            Height          =   375
            Left            =   2580
            Picture         =   "movimento_cheque_avista.frx":7472
            TabIndex        =   18
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Código de Barra &1"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Código de Barra &2"
            Height          =   195
            Index           =   1
            Left            =   2100
            TabIndex        =   14
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Código de Barra &3"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   16
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   10
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "&Tipo de Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Período"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   26
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_cheque_avista.frx":8A7C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cheque_avista.frx":9F76
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cheque_avista.frx":B470
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cheque_avista.frx":C8E2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cheque_avista.frx":DE64
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_cheque_avista.frx":F46E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   2160
      Width           =   795
   End
End
Attribute VB_Name = "movimento_cheque_avista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lPeriodo As String
Dim lTipoMovimento As String
Dim lOrdem As Integer
Dim lGravados As Integer
Dim lDados As String
Dim lCodigoBarra1 As String
Dim lCodigoBarra2 As String
Dim lCodigoBarra3 As String
Dim lQtdPeriodo As Integer
Dim lLeitoraCheque As Boolean

Private Configuracao As New cConfiguracao
Private MovChequeAvista As New cMovimentoChequeAvista
Private Sub AtualizaConstantes()
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
        If Mid(Configuracao.OutrasConfiguracoes, 2, 1) = "S" Then
            lLeitoraCheque = True
        Else
            lLeitoraCheque = False
        End If
    Else
        lQtdPeriodo = 1
        lLeitoraCheque = False
    End If
End Sub
Private Sub AtualTabe()
    MovChequeAvista.Empresa = g_empresa
    MovChequeAvista.DataEmissao = msk_data_emissao.Text
    MovChequeAvista.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    MovChequeAvista.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovChequeAvista.valor = fValidaValor2(msk_valor)
    MovChequeAvista.OrdemDigitacao = lOrdem
    MovChequeAvista.CodigoBarra1 = lCodigoBarra1
    MovChequeAvista.CodigoBarra2 = lCodigoBarra2
    MovChequeAvista.CodigoBarra3 = lCodigoBarra3
    MovChequeAvista.DataEmissao = MovChequeAvista.DataEmissao
End Sub
Private Sub MostraDadosInicial()
    Dim i As Integer
    msk_data_emissao = lData
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
    lData = MovChequeAvista.DataEmissao
    lPeriodo = MovChequeAvista.Periodo
    lTipoMovimento = MovChequeAvista.TipoMovimento
    lOrdem = MovChequeAvista.OrdemDigitacao
    lCodigoBarra1 = MovChequeAvista.CodigoBarra1
    lCodigoBarra2 = MovChequeAvista.CodigoBarra2
    lCodigoBarra3 = MovChequeAvista.CodigoBarra3
    msk_data_emissao = Format(MovChequeAvista.DataEmissao, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovChequeAvista.Periodo - 1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovChequeAvista.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    msk_valor = Format(MovChequeAvista.valor, "###,##0.00")
    lbl_total.Caption = Format(MovChequeAvista.TotalPeriodo(g_empresa, lData, lPeriodo, lTipoMovimento), "###,##0.00")
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub Finaliza()
    Set Configuracao = Nothing
    Set MovChequeAvista = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub InformaCodigoBarra()
    frmCodigoBarra.Top = 300
    frmCodigoBarra.Left = 3120
    frmCodigoBarra.Visible = True
    txt_codigo_barra_1 = lCodigoBarra1
    txt_codigo_barra_2 = lCodigoBarra2
    txt_codigo_barra_3 = lCodigoBarra3
    txt_codigo_barra_1.SetFocus
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
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
    If MovChequeAvista.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If MovChequeAvista.LocalizarUltimo(g_empresa) Then
        AtivaBotoes
        AtualTela
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
        msk_data_emissao = "__/__/____"
        cbo_periodo.ListIndex = -1
        cbo_tipo_movimento.ListIndex = -1
    End If
    msk_valor = ""
    If lLeitoraCheque Then
        lCodigoBarra1 = ""
        lCodigoBarra2 = ""
        lCodigoBarra3 = ""
    Else
        lCodigoBarra1 = "00000000"
        lCodigoBarra2 = "0000000000"
        lCodigoBarra3 = "000000000000"
    End If
End Sub
Private Sub cmd_excluir_Click()
    If msk_data_emissao.Text <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If MovChequeAvista.Excluir(g_empresa, lData, lPeriodo, lTipoMovimento, lOrdem) Then
                LimpaTela
                If MovChequeAvista.LocalizarUltimo(g_empresa) Then
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
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    If lGravados = 0 Then
        If BuscaProximoCaixa Then
            msk_valor.SetFocus
        Else
            msk_data_emissao.SetFocus
        End If
    Else
        msk_valor.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                AtualTabe
                If MovChequeAvista.Incluir Then
                    lData = MovChequeAvista.DataEmissao
                    lPeriodo = MovChequeAvista.Periodo
                    lTipoMovimento = MovChequeAvista.TipoMovimento
                    lOrdem = MovChequeAvista.OrdemDigitacao
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
                lGravados = 1
            ElseIf lOpcao = 2 Then
                AtualTabe
                If MovChequeAvista.Alterar(g_empresa, lData, lPeriodo, lTipoMovimento, lOrdem) Then
                    lData = MovChequeAvista.DataEmissao
                    lPeriodo = MovChequeAvista.Periodo
                    lTipoMovimento = MovChequeAvista.TipoMovimento
                    lOrdem = MovChequeAvista.OrdemDigitacao
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
            End If
            If MovChequeAvista.LocalizarCodigo(g_empresa, lData, lPeriodo, lTipoMovimento, lOrdem) Then
                AtualTela
            Else
                LimpaTela
                MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
            End If
            If lOpcao = 1 Then
                lOpcao = 0
                cmd_novo_Click
            Else
                lOpcao = 0
                cmd_novo.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_movimento_cheque_avista.Name, "Cheque Avistao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    Dim dias As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_emissao) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data_emissao.SetFocus
    ElseIf Not cbo_periodo > "" Then
        MsgBox "Informe o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Informe o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf Not fValidaValor2(msk_valor) > 0 Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Atenção!"
        msk_valor.SetFocus
    ElseIf Not ValidaCodigoBarra Then
        MsgBox "Informe o código de barra.", vbInformation, "Atenção!"
        InformaCodigoBarra
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
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovChequeAvista.Empresa < g_cfg_empresa_i Or MovChequeAvista.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovChequeAvista.DataEmissao < g_cfg_data_i Or MovChequeAvista.DataEmissao > g_cfg_data_f Then
            x_flag = False
        ElseIf MovChequeAvista.Periodo < g_cfg_periodo_i Or MovChequeAvista.Periodo > g_cfg_periodo_f Then
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
    If msk_data_emissao < g_cfg_data_i Or msk_data_emissao > g_cfg_data_f Then
        MsgBox "A data de emissão deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data_emissao.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_ok2_Click()
    frmCodigoBarra.Visible = False
    lCodigoBarra1 = txt_codigo_barra_1
    lCodigoBarra2 = txt_codigo_barra_2
    lCodigoBarra3 = txt_codigo_barra_3
    cmd_ok.SetFocus
End Sub
Private Sub cmd_pesquisa_Click()
    consulta_cheque_avista.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lTipoMovimento = RetiraGString(3)
        lOrdem = RetiraGString(4)
        If MovChequeAvista.LocalizarCodigo(g_empresa, lData, lPeriodo, lTipoMovimento, lOrdem) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovChequeAvista.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovChequeAvista.LocalizarProximo Then
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
    If MovChequeAvista.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaConstantes
        lGravados = 0
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If MovChequeAvista.LocalizarUltimo(g_empresa) Then
            AtivaBotoes
            AtualTela
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
    frmCodigoBarra.Visible = False
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
    frmCodigoBarra.Visible = False
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
    PreencheCboPeriodo
    PreencheCboTipoMovimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_emissao_GotFocus()
    If Not IsDate(msk_data_emissao) Then
        msk_data_emissao = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
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
    If lOpcao = 1 And IsDate(msk_data_emissao) And CDate(msk_data_emissao) <> lData Then
        BuscaPeriodo
    End If
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
        If msk_valor <> "" Then
            If lLeitoraCheque And CCur(msk_valor) > 0 Then
                LeituraCheque
                If Not ValidaCodigoBarra Then
                    InformaCodigoBarra
                End If
            End If
        End If
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    If Val(msk_valor) > 0 Then
        msk_valor = Format(msk_valor, "###,##0.00")
    End If
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub BuscaPeriodo()
    cbo_periodo.ListIndex = 0
    cbo_tipo_movimento.ListIndex = 0
    If MovChequeAvista.LocalizarUltimo(g_empresa) Then
        If CDate(msk_data_emissao.Text) = MovChequeAvista.DataEmissao Then
            If MovChequeAvista.Periodo < 4 Then
                cbo_periodo.ListIndex = MovChequeAvista.Periodo
            End If
        End If
    End If
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    If MovChequeAvista.LocalizarUltimo(g_empresa) Then
        If MovChequeAvista.Empresa = g_empresa Then
            msk_data_emissao.Text = MovChequeAvista.DataEmissao
            x_periodo = MovChequeAvista.Periodo
            If MovChequeAvista.Periodo >= lQtdPeriodo Then
                msk_data_emissao.Text = MovChequeAvista.DataEmissao + 1
                x_periodo = 0
            End If
            cbo_periodo.ListIndex = x_periodo
            cbo_tipo_movimento.ListIndex = 0
            BuscaProximoCaixa = True
            Exit Function
        End If
        msk_data_emissao.Text = g_data_def - 1
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
    End If
End Function
Private Sub txt_codigo_barra_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_codigo_barra_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_barra_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_codigo_barra_3.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_barra_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
