VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_mala_direta 
   Caption         =   "Cadastro de Mala Direta"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_dados 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.TextBox txt_cpf_cnpj 
         Height          =   285
         Left            =   1800
         MaxLength       =   18
         TabIndex        =   6
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   26
         Top             =   4560
         Width           =   4815
      End
      Begin VB.TextBox txt_pessoa_contato 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   24
         Top             =   4200
         Width           =   4815
      End
      Begin VB.TextBox txt_fax 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   22
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txt_telefone_2 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   20
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txt_telefone_1 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txt_endereco 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txt_nome_razao_social 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txt_uf 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   14
         Top             =   2400
         Width           =   375
      End
      Begin MSMask.MaskEdBox msk_data_nascimento 
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   4920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   2760
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "__.___-___"
         Mask            =   "##.###-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "E-&mail"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Pess&oa p/ Contato"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "&Fax"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Telefone &2"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "C&PF/CNPJ"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "&Data Nascimento"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Telefone &1"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Endereço"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Nome/Razão Social"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Código"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "C&idade"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "&Sigla do Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "C.&E.P."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cadastro_mala_direta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cria um novo registro."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cadastro_mala_direta.frx":1692
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Altera o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cadastro_mala_direta.frx":2B8C
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cadastro_mala_direta.frx":421E
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cadastro_mala_direta.frx":5690
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5520
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4860
      TabIndex        =   36
      Top             =   5400
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cadastro_mala_direta.frx":6D22
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cadastro_mala_direta.frx":821C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cadastro_mala_direta.frx":9716
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cadastro_mala_direta.frx":AB88
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5340
      Picture         =   "cadastro_mala_direta.frx":C10A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6240
      Picture         =   "cadastro_mala_direta.frx":D714
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5520
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_mala_direta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Private MalaDireta As cMalaDireta
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        MalaDireta.Codigo = Val(txt_codigo.Text)
    End If
    MalaDireta.NomeRazaoSocial = txt_nome_razao_social.Text
    MalaDireta.CPFCNPJ = txt_cpf_cnpj.Text
    MalaDireta.Endereco = txt_endereco.Text
    MalaDireta.Bairro = txt_bairro.Text
    MalaDireta.Cidade = txt_cidade.Text
    MalaDireta.UnidadeFederativa = txt_uf.Text
    MalaDireta.CEP = msk_cep.Text
    MalaDireta.Telefone1 = fDesmascaraTelefone(txt_telefone_1.Text)
    MalaDireta.Telefone2 = fDesmascaraTelefone(txt_telefone_2.Text)
    MalaDireta.Fax = fDesmascaraTelefone(txt_fax.Text)
    MalaDireta.PessoaparaContato = txt_pessoa_contato.Text
    MalaDireta.Email = txt_email.Text
    If msk_data_nascimento.Text = "__/__/____" Then
        MalaDireta.DataNascimento = "00:00:00"
    Else
        MalaDireta.DataNascimento = msk_data_nascimento.Text
    End If
End Sub
Private Sub AtualTela()
    lCodigo = MalaDireta.Codigo
    txt_codigo.Text = MalaDireta.Codigo
    txt_nome_razao_social.Text = MalaDireta.NomeRazaoSocial
    txt_cpf_cnpj = MalaDireta.CPFCNPJ
    txt_endereco.Text = MalaDireta.Endereco
    txt_bairro.Text = MalaDireta.Bairro
    txt_cidade.Text = MalaDireta.Cidade
    txt_uf.Text = MalaDireta.UnidadeFederativa
    msk_cep = MalaDireta.CEP
    txt_telefone_1.Text = fMascaraTelefone(MalaDireta.Telefone1)
    txt_telefone_2.Text = fMascaraTelefone(MalaDireta.Telefone2)
    txt_fax.Text = fMascaraTelefone(MalaDireta.Fax)
    txt_pessoa_contato.Text = MalaDireta.PessoaparaContato
    txt_email.Text = MalaDireta.Email
    If MalaDireta.DataNascimento = "00:00:00" Then
        msk_data_nascimento = "__/__/____"
    Else
        msk_data_nascimento = Format(MalaDireta.DataNascimento, "dd/mm/yyyy")
    End If
    frm_dados.Enabled = False
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
Private Sub Finaliza()
    Set MalaDireta = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If MalaDireta.LocalizarUltimo Then
        txt_codigo.Text = MalaDireta.Codigo + 1
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
    txt_nome_razao_social.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MalaDireta.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If MalaDireta.LocalizarCodigo(lCodigo) Then
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
    txt_codigo.Text = ""
    txt_nome_razao_social.Text = ""
    txt_cpf_cnpj.Text = ""
    txt_endereco.Text = ""
    txt_bairro.Text = ""
    txt_cidade.Text = ""
    txt_uf.Text = ""
    msk_cep = "__.___-___"
    txt_telefone_1.Text = ""
    txt_telefone_2.Text = ""
    txt_fax.Text = ""
    txt_pessoa_contato.Text = ""
    txt_email.Text = ""
    msk_data_nascimento.Text = "__/__/____"
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If MalaDireta.Excluir(txt_codigo.Text) Then
                LimpaTela
                If MalaDireta.LocalizarUltimo Then
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
    txt_nome_razao_social.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If MalaDireta.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not MalaDireta.Alterar(lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call MalaDireta.LocalizarCodigo(lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_MalaDireta.Name, "MalaDiretao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo.Text) > 0 Then
        MsgBox "Informe o código.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome_razao_social.Text = "" Then
        MsgBox "Informe o nome/razão social.", vbInformation, "Atenção!"
        txt_nome_razao_social.SetFocus
    ElseIf txt_cpf_cnpj = "" Then
        MsgBox "Informe o CPF/CNPJ.", vbInformation, "Atenção!"
        txt_cpf_cnpj.SetFocus
    ElseIf txt_endereco.Text = "" Then
        MsgBox "Informe o endereço.", vbInformation, "Atenção!"
        txt_endereco.SetFocus
    ElseIf txt_bairro.Text = "" Then
        MsgBox "Informe o bairro.", vbInformation, "Atenção!"
        txt_bairro.SetFocus
    ElseIf txt_cidade.Text = "" Then
        MsgBox "Informe a cidade.", vbInformation, "Atenção!"
        txt_cidade.SetFocus
    ElseIf txt_uf.Text = "" Then
        MsgBox "Informe a sígla do estado.", vbInformation, "Atenção!"
        txt_uf.SetFocus
    ElseIf Val(msk_cep) < 10000000 Then
        MsgBox "Informe um CEP válido.", vbInformation, "Atenção!"
        msk_cep.SetFocus
    ElseIf msk_data_nascimento.Text <> "__/__/____" And Not IsDate(msk_data_nascimento.Text) Then
        MsgBox "Informe uma data de nascimento válida.", vbInformation, "Atenção!"
        msk_data_nascimento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_mala_direta.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        Call MalaDireta.LocalizarCodigo(lCodigo)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MalaDireta.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MalaDireta.LocalizarProximo Then
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
    If MalaDireta.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If MalaDireta.LocalizarUltimo Then
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
    Set MalaDireta = New cMalaDireta
    Set MalaDireta.Conexao = Conectar.Conexao
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub msk_data_nascimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub msk_cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone_1.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cpf_cnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_endereco.SetFocus
    End If
End Sub
Private Sub txt_email_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_nascimento.SetFocus
    End If
End Sub
Private Sub txt_fax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_pessoa_contato.SetFocus
    End If
End Sub
Private Sub txt_pessoa_contato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_email.SetFocus
    End If
End Sub
Private Sub txt_telefone_1_GotFocus()
    txt_telefone_1.Text = fDesmascaraTelefone(txt_telefone_1.Text)
End Sub
Private Sub txt_telefone_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_bairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cidade.SetFocus
    End If
End Sub
Private Sub txt_cidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_uf.SetFocus
    End If
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome_razao_social.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If MalaDireta.LocalizarCodigo(CLng(txt_codigo.Text)) Then
            MsgBox "Já existe Mala Direta cadastrada com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_endereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_bairro.SetFocus
    End If
End Sub
Private Sub txt_nome_razao_social_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cpf_cnpj.SetFocus
    End If
End Sub
Private Sub txt_nome_razao_social_LostFocus()
    If lOpcao = 1 And txt_nome_razao_social.Text <> "" Then
        If MalaDireta.LocalizarNome(txt_nome_razao_social.Text) Then
            If (MsgBox("Já existe Mala Direta cadastrada com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome_razao_social.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_telefone_1_LostFocus()
    txt_telefone_1.Text = fMascaraTelefone(txt_telefone_1.Text)
End Sub
Private Sub txt_telefone_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_fax.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cep.SetFocus
    End If
End Sub

