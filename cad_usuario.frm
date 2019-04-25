VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_usuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   4395
   ClientLeft      =   1395
   ClientTop       =   2385
   ClientWidth     =   7035
   Icon            =   "cad_usuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_usuario.frx":030A
   ScaleHeight     =   4395
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_usuario.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cria um novo registro."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_usuario.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Altera o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_usuario.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_usuario.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3420
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6795
      Begin VB.ComboBox cbo_nivel_acesso 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   3675
      End
      Begin VB.TextBox txt_senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   2700
         Width           =   1035
      End
      Begin VB.TextBox txt_nome 
         Height          =   315
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin MSMask.MaskEdBox msk_hora_saida 
         Height          =   315
         Left            =   1740
         TabIndex        =   8
         Top             =   1440
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_hora_entrada 
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   1020
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption opt_situacao 
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   11
         Top             =   1920
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Inativo"
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
      Begin Threed.SSOption opt_situacao 
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   10
         Top             =   1920
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Ativo"
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
      Begin VB.Label Label4 
         Caption         =   "Senha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nível de Acesso"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Situação"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Hora de Saida"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Hora de Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Usuário"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Código do Usuário"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   22
      Top             =   3300
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_usuario.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_usuario.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_usuario.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_usuario.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cad_usuario.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_usuario.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3420
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Private Usuario As New cUsuario
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub PreencheCboNivelAcesso()
    cbo_nivel_acesso.Clear
    cbo_nivel_acesso.AddItem "Desenvolvimento"
    cbo_nivel_acesso.ItemData(cbo_nivel_acesso.NewIndex) = 1
    cbo_nivel_acesso.AddItem "Digitação      "
    cbo_nivel_acesso.ItemData(cbo_nivel_acesso.NewIndex) = 5
    cbo_nivel_acesso.AddItem "Diretoria      "
    cbo_nivel_acesso.ItemData(cbo_nivel_acesso.NewIndex) = 2
    cbo_nivel_acesso.AddItem "Gerência       "
    cbo_nivel_acesso.ItemData(cbo_nivel_acesso.NewIndex) = 3
    cbo_nivel_acesso.AddItem "Operação       "
    cbo_nivel_acesso.ItemData(cbo_nivel_acesso.NewIndex) = 4
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        Usuario.Codigo = "" & txt_codigo.Text
    End If
    Usuario.Nome = "" & txt_nome.Text
    Usuario.HoraEntrada = "" & msk_hora_entrada.Text
    Usuario.HoraSaida = "" & msk_hora_saida.Text
    If opt_situacao(0) Then
        Usuario.Situacao = "A"
    Else
        Usuario.Situacao = "I"
    End If
    Usuario.Senha = "" & Kriptografa(txt_senha.Text)
    Usuario.TipoAcesso = cbo_nivel_acesso.ItemData(cbo_nivel_acesso.ListIndex)
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lCodigo = Usuario.Codigo
    txt_codigo = Usuario.Codigo
    txt_nome = Usuario.Nome
    msk_hora_entrada = Format(Usuario.HoraEntrada, "hh:mm")
    msk_hora_saida = Format(Usuario.HoraSaida, "hh:mm")
    If Usuario.Situacao = "A" Then
        opt_situacao(0) = True
    Else
        opt_situacao(1) = True
    End If
    For i = 0 To cbo_nivel_acesso.ListCount - 1
        cbo_nivel_acesso.ListIndex = i
        If cbo_nivel_acesso.ItemData(i) = Usuario.TipoAcesso Then
            Exit For
        End If
    Next
    txt_senha = DesKriptografa(Usuario.Senha)
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set Usuario = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If Usuario.LocalizarUltimo Then
        txt_codigo.Text = Usuario.Codigo + 1
    End If
End Sub
Private Sub cbo_nivel_acesso_GotFocus()
    SendMessageLong cbo_nivel_acesso.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_nivel_acesso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha.SetFocus
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
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Usuario.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Usuario.LocalizarCodigo(lCodigo) Then
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
    txt_codigo = ""
    txt_nome = ""
    msk_hora_entrada = "__:__"
    msk_hora_saida = "__:__"
    opt_situacao(0) = True
    cbo_nivel_acesso.ListIndex = -1
    txt_senha = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Usuario.Excluir(txt_codigo.Text) Then
                LimpaTela
                If Usuario.LocalizarUltimo Then
                    AtualTela
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Usuario.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Usuario.Alterar(lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Usuario.LocalizarCodigo(lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_usuario.Name, "Usuarioo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código do usuário.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe o nome do usuário.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf Mid(msk_hora_entrada, 1, 1) = "_" Or Mid(msk_hora_entrada, 2, 1) = "_" Or Mid(msk_hora_entrada, 4, 1) = "_" Or Mid(msk_hora_entrada, 5, 1) = "_" Then
        MsgBox "Informe a hora de entrada corretamento." & Chr(10) & "Exemplo 08:00", vbInformation, "Atenção!"
        msk_hora_entrada.SetFocus
    ElseIf Mid(msk_hora_saida, 1, 1) = "_" Or Mid(msk_hora_saida, 2, 1) = "_" Or Mid(msk_hora_saida, 4, 1) = "_" Or Mid(msk_hora_saida, 5, 1) = "_" Then
        MsgBox "Informe a hora de saida corretamento." & Chr(10) & "Exemplo 18:00", vbInformation, "Atenção!"
        msk_hora_saida.SetFocus
    ElseIf opt_situacao(0) = False And opt_situacao(1) = False Then
        MsgBox "Escolha a situação.", vbInformation, "Atenção!"
        opt_situacao(0).SetFocus
    ElseIf cbo_nivel_acesso.ListIndex = -1 Then
        MsgBox "Selecione o nível de acesso.", vbInformation, "Atenção!"
        cbo_nivel_acesso.SetFocus
    ElseIf txt_senha = "" Then
        MsgBox "Informe a senha do usuário.", vbInformation, "Atenção!"
        txt_senha.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Usuario.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Usuario.LocalizarProximo Then
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
    If Usuario.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Usuario.LocalizarUltimo Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagCadastro = 0
    End If
End Sub
Private Sub Form_Deactivate()
    lFlagCadastro = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 Then
        KeyCode = 0
        cmd_excluir_Click
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
    PreencheCboNivelAcesso
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_hora_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_saida.SetFocus
    End If
End Sub
Private Sub msk_hora_saida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        opt_situacao(0).SetFocus
    End If
End Sub
Private Sub opt_situacao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_nivel_acesso.SetFocus
    End If
End Sub
Private Sub txt_codigo_GotFocus()
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Usuario.LocalizarCodigo(txt_codigo.Text) Then
            MsgBox "Já usuário cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_entrada.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If Usuario.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Já existe usuário cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_senha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
