VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form movimento_cheque_avista 
   Caption         =   "Movimenta��o de Cheques � Vista"
   ClientHeight    =   3135
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cheque_avista.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cheque_avista.frx":0446
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
      Picture         =   "movimento_cheque_avista.frx":1720
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
      Picture         =   "movimento_cheque_avista.frx":29FA
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
      Picture         =   "movimento_cheque_avista.frx":3CD4
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro espec�fico."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cheque_avista.frx":4FAE
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
      Begin Threed.SSFrame frm_codigo_barra 
         Height          =   1515
         Left            =   3120
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   2672
         _StockProps     =   14
         Caption         =   "C�digo de Barra"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         ShadowStyle     =   1
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
            TabIndex        =   18
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "C�digo de Barra &1"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "C�digo de Barra &2"
            Height          =   195
            Index           =   1
            Left            =   2100
            TabIndex        =   14
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "C�digo de Barra &3"
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
         Caption         =   "&Data de Emiss�o"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Per�odo"
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
         Picture         =   "movimento_cheque_avista.frx":6288
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cheque_avista.frx":7562
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o �ltimo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cheque_avista.frx":883C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cheque_avista.frx":9B16
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o pr�ximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cheque_avista.frx":ADF0
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
      Picture         =   "movimento_cheque_avista.frx":C0CA
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
Dim lTotal As Currency
Dim lDados As String
Dim lCodigoBarra1 As String
Dim lCodigoBarra2 As String
Dim lCodigoBarra3 As String
Dim lQtdPeriodo As Integer
Dim lLeitoraCheque As Boolean
Dim tbl_configuracao As Table
Private MovChequeAvista As cMovimentoChequeAvista
Private Sub AtualizaConstantes()
    tbl_configuracao.Index = "id_codigo"
    tbl_configuracao.Seek "=", g_empresa
    If Not tbl_configuracao.NoMatch Then
        lQtdPeriodo = tbl_configuracao![Quantidade de Periodos]
        If Mid(tbl_configuracao![Outras Configuracoes], 2, 1) = "S" Then
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
    MovChequeAvista.[Data de Emissao] = msk_data_emissao
    MovChequeAvista.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    MovChequeAvista.[Tipo do Movimento] = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovChequeAvista.valor = fValidaValor2(msk_valor)
    MovChequeAvista.[Ordem da Digitacao] = lOrdem
    MovChequeAvista.[Codigo de Barra 1] = lCodigoBarra1
    MovChequeAvista.[Codigo de Barra 2] = lCodigoBarra2
    MovChequeAvista.[Codigo de Barra 3] = lCodigoBarra3
    MovChequeAvista.Data = MovChequeAvista.[Data de Emissao]
    lPeriodo = MovChequeAvista.Periodo
    lTipoMovimento = MovChequeAvista.[Tipo do Movimento]
    lOrdem = MovChequeAvista.[Ordem da Digitacao]
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_movimento_cheque_avista
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
            .Seek "<", g_empresa, CDate("31/12/2500"), "Z", "Z", 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    BuscaDados = True
                    Exit Function
                End If
            End If
        End If
        lblTotal = ""
        lGravados = 0
        LimpaTela
    End With
End Function
Function BuscaRegistro(x_data As Date, x_periodo As String, x_tipo_movimento As String, x_ordem As Integer) As Boolean
    BuscaRegistro = False
    If tbl_movimento_cheque_avista.RecordCount > 0 Then
        tbl_movimento_cheque_avista.Seek "=", g_empresa, x_data, x_periodo, x_tipo_movimento, x_ordem
        If Not tbl_movimento_cheque_avista.NoMatch Then
            If tbl_movimento_cheque_avista!Empresa = g_empresa Then
                AtualTela
                BuscaRegistro = True
            End If
        End If
    End If
End Function
Private Sub BuscaOrdemDigitacao()
    Dim x_tipo_movimento As String * 1
    x_tipo_movimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    lOrdem = 0
    If tbl_movimento_cheque_avista.RecordCount > 0 Then
        tbl_movimento_cheque_avista.Seek "<", g_empresa, CDate(msk_data_emissao), cbo_periodo, x_tipo_movimento, 9999
        If Not tbl_movimento_cheque_avista.NoMatch Then
            If tbl_movimento_cheque_avista!Empresa = g_empresa And tbl_movimento_cheque_avista![Data de Emissao] = CDate(msk_data_emissao) And tbl_movimento_cheque_avista!Periodo = cbo_periodo And tbl_movimento_cheque_avista![Tipo do Movimento] = x_tipo_movimento Then
                lOrdem = tbl_movimento_cheque_avista![Ordem da Digitacao]
            End If
        End If
    End If
    lOrdem = Val(lOrdem) + 1
End Sub
Private Sub Totaliza()
    lTotal = 0
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, CDate(msk_data_emissao), cbo_periodo, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Data de Emissao] <> CDate(msk_data_emissao) Then
                        Exit Do
                    ElseIf !Periodo = cbo_periodo And ![Tipo do Movimento] = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) Then
                        lTotal = lTotal + !valor
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    lblTotal = Format(lTotal, "###,##0.00")
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
    cbo_tipo_movimento.AddItem "1 - Caixa de Combust�veis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de �leos/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Cheque Inclus�o"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = MovChequeAvista.[Data de Emissao]
    lPeriodo = MovChequeAvista.Periodo
    lTipoMovimento = MovChequeAvista.[Tipo do Movimento]
    lOrdem = MovChequeAvista.[Ordem da Digitacao]
    lCodigoBarra1 = MovChequeAvista.[Codigo de Barra 1]
    lCodigoBarra2 = MovChequeAvista.[Codigo de Barra 2]
    lCodigoBarra3 = MovChequeAvista.[Codigo de Barra 3]
    msk_data_emissao = Format(MovChequeAvista.[Data de Emissao], "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovChequeAvista.Periodo - 1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovChequeAvista.[Tipo do Movimento] Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    msk_valor = Format(MovChequeAvista.valor, "###,##0.00")
    Totaliza
    .Seek "=", g_empresa, lData, lPeriodo, lTipoMovimento, lOrdem
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub Finaliza()
    tbl_configuracao.Close
    'tbl_movimento_cheque_avista.Close
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub InformaCodigoBarra()
    frm_codigo_barra.Top = 300
    frm_codigo_barra.Left = 3120
    frm_codigo_barra.Visible = True
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
    If tbl_movimento_cheque_avista.RecordCount > 0 Then
        tbl_movimento_cheque_avista.MovePrevious
        If Not tbl_movimento_cheque_avista.BOF Then
            If tbl_movimento_cheque_avista!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "In�cio de Arquivo.", vbInformation, "Aten��o!"
        tbl_movimento_cheque_avista.MoveNext
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(lData, lPeriodo, lTipoMovimento, lOrdem) Then
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
        MsgBox "Cheque N�o Inserido!"
    ElseIf x = 1 Then
        Open "\VB5\SGP\DATA\DR10.RET" For Input As #1
        Line Input #1, lDados
        Close #1
        lCodigoBarra1 = Mid(lDados, 2, 8)
        lCodigoBarra2 = Mid(lDados, 11, 10)
        lCodigoBarra3 = Mid(lDados, 22, 12)
    Else
        MsgBox "Erro n�o identificado! " & x
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
    If tbl_movimento_cheque_avista![Ordem da Digitacao] > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclus�o de Registro!")) = 6 Then
            tbl_movimento_cheque_avista.Edit
            tbl_movimento_cheque_avista.Delete
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
                BuscaOrdemDigitacao
                tbl_movimento_cheque_avista.AddNew
                AtualTabe
                tbl_movimento_cheque_avista.Update
                lGravados = 1
            ElseIf lOpcao = 2 Then
                tbl_movimento_cheque_avista.Edit
                AtualTabe
                tbl_movimento_cheque_avista.Update
            End If
            Call BuscaRegistro(lData, lPeriodo, lTipoMovimento, lOrdem)
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
    ErroArquivo tbl_movimento_cheque_avista.Name, "Cheque Avistao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    Dim dias As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_emissao) Then
        MsgBox "Informe a data de emiss�o.", vbInformation, "Aten��o!"
        msk_data_emissao.SetFocus
    ElseIf Not cbo_periodo > "" Then
        MsgBox "Informe o per�odo.", vbInformation, "Aten��o!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Informe o tipo de movimento.", vbInformation, "Aten��o!"
        cbo_tipo_movimento.SetFocus
    ElseIf Not fValidaValor2(msk_valor) > 0 Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Aten��o!"
        msk_valor.SetFocus
    ElseIf Not ValidaCodigoBarra Then
        MsgBox "Informe o c�digo de barra.", vbInformation, "Aten��o!"
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
    With tbl_movimento_cheque_avista
        If g_nivel_acesso > 4 Then
            If !Empresa < g_cfg_empresa_i Or !Empresa > g_cfg_empresa_f Then
                x_flag = False
            ElseIf ![Data de Emissao] < g_cfg_data_i Or ![Data de Emissao] > g_cfg_data_f Then
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
    If msk_data_emissao < g_cfg_data_i Or msk_data_emissao > g_cfg_data_f Then
        MsgBox "A data de emiss�o deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digita��o N�o Autorizada!"
        msk_data_emissao.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O per�odo deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digita��o N�o Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_ok2_Click()
    frm_codigo_barra.Visible = False
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
        Call BuscaRegistro(lData, lPeriodo, lTipoMovimento, lOrdem)
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovChequeAvista.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "N�o h� registros nesta empresa.", vbInformation, "Erro de Verifica��o!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If tbl_movimento_cheque_avista.RecordCount > 0 Then
        tbl_movimento_cheque_avista.MoveNext
        If Not tbl_movimento_cheque_avista.EOF Then
            If tbl_movimento_cheque_avista!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Fim de Arquivo.", vbInformation, "Aten��o!"
        tbl_movimento_cheque_avista.MovePrevious
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If tbl_movimento_cheque_avista.RecordCount > 0 Then
        tbl_movimento_cheque_avista.Seek "<", g_empresa, CDate("31/12/2500"), "9", "9", 9999
        If Not tbl_movimento_cheque_avista.NoMatch Then
            If tbl_movimento_cheque_avista!Empresa = g_empresa Then
                AtualTela
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "N�o h� registros nesta empresa.", vbInformation, "Erro de Verifica��o!"
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
    frm_codigo_barra.Visible = False
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
    frm_codigo_barra.Visible = False
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_configuracao = bd_sgp.OpenTable("configuracao")
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    tbl_movimento_cheque_avista.Index = "id_digitacao"
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
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate(msk_data_emissao), "9", "9", 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    If CDate(msk_data_emissao) = ![Data de Emissao] Then
                        If !Periodo < 4 Then
                            cbo_periodo.ListIndex = !Periodo
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate("31/12/2500"), "9", "9", 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    msk_data_emissao = ![Data de Emissao]
                    x_periodo = !Periodo
                    If !Periodo >= lQtdPeriodo Then
                        msk_data_emissao = ![Data de Emissao] + 1
                        x_periodo = 0
                    End If
                    cbo_periodo.ListIndex = x_periodo
                    cbo_tipo_movimento.ListIndex = 0
                    BuscaProximoCaixa = True
                    Exit Function
                End If
            End If
        End If
        msk_data_emissao = g_data_def - 1
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
    End With
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
