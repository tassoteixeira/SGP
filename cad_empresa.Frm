VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_empresa 
   Caption         =   "Cadastro de Empresas"
   ClientHeight    =   5025
   ClientLeft      =   180
   ClientTop       =   2955
   ClientWidth     =   6885
   Icon            =   "cad_empresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_empresa.frx":030A
   ScaleHeight     =   5025
   ScaleWidth      =   6885
   Begin VB.Frame frm_dados 
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.TextBox txt_inscricao_estadual 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2700
         Width           =   2535
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_endereco 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txt_estado 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txt_distribuidora 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   20
         Top             =   3120
         Width           =   4815
      End
      Begin VB.CheckBox chk_inativo 
         Caption         =   "&Empresa Inativa"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   3540
         Width           =   1515
      End
      Begin MSMask.MaskEdBox msk_cgc 
         Height          =   300
         Left            =   1680
         TabIndex        =   16
         Top             =   2400
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         Format          =   "__.___.___/____-__"
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   300
         Left            =   5340
         TabIndex        =   14
         Top             =   2040
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "__.___-___"
         Mask            =   "##.###-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Inscrição Estadual"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "C.G.C."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Razão Social"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Distribuidora"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "C.E.P."
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   2100
         Width           =   975
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_empresa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cria um novo registro."
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_empresa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Altera o registro atual."
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_empresa.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_empresa.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4080
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4560
      TabIndex        =   28
      Top             =   3960
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_empresa.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_empresa.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_empresa.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_empresa.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5940
      Picture         =   "cad_empresa.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5040
      Picture         =   "cad_empresa.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4080
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Private Empresa As New cEmpresa
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    Empresa.Codigo = Val(txt_codigo.Text)
    Empresa.Nome = txt_nome.Text
    Empresa.Endereco = txt_endereco.Text
    Empresa.Bairro = txt_bairro.Text
    Empresa.Cidade = txt_cidade.Text
    Empresa.Estado = txt_estado.Text
    Empresa.CEP = msk_cep
    Empresa.CGC = msk_cgc
    Empresa.InscricaoEstadual = txt_inscricao_estadual.Text
    Empresa.Distribuidora = txt_distribuidora.Text
    Empresa.Inativo = chk_inativo.Value
End Sub
Private Sub AtualTela()
    lCodigo = Empresa.Codigo
    txt_codigo.Text = Empresa.Codigo
    txt_nome.Text = Empresa.Nome
    txt_endereco.Text = Empresa.Endereco
    txt_bairro.Text = Empresa.Bairro
    txt_cidade.Text = Empresa.Cidade
    txt_estado.Text = Empresa.Estado
    msk_cep.Text = Empresa.CEP
    msk_cgc.Text = Mid(Empresa.CGC, 1, 14)
    txt_inscricao_estadual.Text = Empresa.InscricaoEstadual
    txt_distribuidora.Text = Empresa.Distribuidora
    If Empresa.Inativo Then
        chk_inativo.Value = 1
    Else
        chk_inativo.Value = 0
    End If
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set Empresa = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If Empresa.LocalizarUltimo Then
        txt_codigo.Text = Empresa.Codigo + 1
    End If
End Sub
Private Sub chk_inativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
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
    If Empresa.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Empresa.LocalizarCodigo(lCodigo) Then
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
    txt_endereco = ""
    txt_bairro = ""
    txt_cidade = ""
    txt_estado = ""
    msk_cep = "__.___-___"
    msk_cgc = "__.___.___/____-__"
    txt_inscricao_estadual = ""
    txt_distribuidora = ""
    chk_inativo.Value = 0
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Empresa.Excluir(Val(txt_codigo.Text)) Then
                LimpaTela
                If Empresa.LocalizarUltimo Then
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
            If Empresa.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Empresa.Alterar(lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Empresa.LocalizarCodigo(lCodigo)
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_empresa.Name, "Empresaa"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código da empresa.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe a razão social da empresa.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf txt_endereco = "" Then
        MsgBox "Informe o endereco.", vbInformation, "Atenção!"
        txt_endereco.SetFocus
    ElseIf txt_bairro = "" Then
        MsgBox "Informe o bairro.", vbInformation, "Atenção!"
        txt_bairro.SetFocus
    ElseIf txt_cidade = "" Then
        MsgBox "Informe a cidade.", vbInformation, "Atenção!"
        txt_cidade.SetFocus
    ElseIf txt_estado = "" Then
        MsgBox "Informe o estado.", vbInformation, "Atenção!"
        txt_estado.SetFocus
    ElseIf Val(msk_cep) < 10000000 Then
        MsgBox "Informe o cep.", vbInformation, "Atenção!"
        msk_cep.SetFocus
    ElseIf msk_cgc = "" Then
        MsgBox "Informe o CGC.", vbInformation, "Atenção!"
        msk_cgc.SetFocus
    ElseIf txt_inscricao_estadual = "" Then
        MsgBox "Informe a Inscrição Estadual.", vbInformation, "Atenção!"
        txt_inscricao_estadual.SetFocus
    ElseIf txt_distribuidora = "" Then
        MsgBox "Informe a distribuidora.", vbInformation, "Atenção!"
        txt_distribuidora.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Empresa.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Empresa.LocalizarProximo Then
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
    If Empresa.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Empresa.LocalizarUltimo Then
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
End Sub
Private Sub msk_cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cgc.SetFocus
    End If
End Sub
Private Sub msk_cgc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
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
        txt_estado.SetFocus
    End If
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
        If Empresa.LocalizarCodigo(Val(txt_codigo.Text)) Then
            MsgBox "Já existe empresa cadastrada com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_distribuidora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_inativo.SetFocus
    End If
End Sub
Private Sub txt_endereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_bairro.SetFocus
    End If
End Sub
Private Sub txt_estado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cep.SetFocus
    End If
End Sub
Private Sub txt_inscricao_estadual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_distribuidora.SetFocus
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_endereco.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If Empresa.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Já existe empresa cadastrada com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
